import { describe, expect, it } from "vitest";
import type { Caller } from "../src/contracts.js";
import type { StoredSync, SyncJob, SyncStore } from "../src/sync-contracts.js";
import { SyncService, syncCreateSchema, syncSettings, syncUpdateSchema } from "../src/sync.js";
import { ApiError } from "../src/model.js";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
class Store implements SyncStore {
  active: StoredSync | null = null;
  rows = new Map<string, SyncJob>();
  revision = 0;
  async current() { return structuredClone(this.active); }
  async get(id: string) { return structuredClone(this.rows.get(id) ?? null); }
  async save(job: SyncJob, etag?: string) {
    if (etag !== this.active?.etag) return false;
    this.active = { job: structuredClone(job), etag: String(++this.revision) };
    this.rows.set(job.id, structuredClone(job)); return true;
  }
  async acknowledge(id: string, revision: number) {
    const job = await this.get(id);
    if (!job || job.revision !== revision) return job;
    job.mirrorPending = false;
    this.rows.set(id, job);
    if (this.active?.job.id === id) this.active.job = structuredClone(job);
    return job;
  }
  async prune() { return 0; }
  async list() { return { items: [...this.rows.values()] }; }
}
function setup() {
  let time = Date.parse("2026-10-28T13:00:00Z");
  const store = new Store();
  const caller: Caller = { tenantId: id(1), objectId: id(2) };
  const other: Caller = { ...caller, objectId: id(3) };
  const settings = { enabled: true, siteId: `example.sharepoint.com,${id(4)},${id(5)}`, policyListId: id(6),
    balanceListId: id(7), jobListId: id(8), defaultConcurrency: 3, maxConcurrency: 8, retentionDays: 30 };
  const service = new SyncService(store, settings, (who) => {
    if (who.tenantId !== caller.tenantId || ![caller.objectId, other.objectId].includes(who.objectId)) throw new ApiError(403, "admin", "Admin required.");
  }, () => new Date(time));
  const create = (operationId = id(9)) => service.create(caller, { operationId, concurrency: 3 });
  return { store, service, caller, other, settings, create, advance: () => { time += 91_000; } };
}
const policy = { id: id(10), displayName: "Policy", type: "#microsoft.graph.selectedGroupsSpendingPolicy" };

describe("durable administrator sync tasks", () => {
  it("creates idempotently and prevents overlapping tenant work", async () => {
    const { service, caller, create } = setup();
    expect(await create()).toMatchObject({ state: "queued", mirrorPending: true });
    expect(await create()).toMatchObject({ id: id(9) });
    await expect(create(id(11))).rejects.toMatchObject({ code: "sync-active" });
    await expect(service.create(caller, { operationId: id(9), concurrency: 4 })).rejects.toMatchObject({ code: "sync-operation-conflict" });
  });
  it("checks admin and tenant on configuration, reads, writes and maintenance", async () => {
    const { service, caller, create } = setup(); await create();
    const denied = [{ ...caller, objectId: id(99) }, { ...caller, tenantId: id(98) }];
    for (const who of denied) {
      expect(() => service.configuration(who)).toThrow("Admin");
      await expect(service.list(who)).rejects.toMatchObject({ status: 403 });
      await expect(service.get(who, id(9))).rejects.toMatchObject({ status: 403 });
      await expect(service.update(who, id(9), { action: "cancel" })).rejects.toMatchObject({ status: 403 });
      await expect(service.maintain(who, { jobId: id(9), leaseId: id(99) })).rejects.toMatchObject({ status: 403 });
    }
  });
  it("allows any admin to take over only after expiry and fences out the previous browser", async () => {
    const { service, caller, other, create, advance } = setup(); await create();
    const first = await service.update(caller, id(9), { action: "claim" });
    await service.update(caller, id(9), { action: "inventory", leaseId: first.lease!.id, policies: [policy] });
    await service.update(caller, id(9), { action: "policy", leaseId: first.lease!.id, policy: { ...policy, state: "running" } });
    await expect(service.update(other, id(9), { action: "claim" })).rejects.toMatchObject({ code: "sync-leased" });
    advance();
    const second = await service.update(other, id(9), { action: "claim" });
    expect(second.lease?.id).not.toBe(first.lease?.id);
    expect(second.actors).toEqual([caller.objectId, other.objectId]);
    expect(second.policies[0].state).toBe("pending");
    await expect(service.update(caller, id(9), { action: "heartbeat", leaseId: first.lease!.id })).rejects.toMatchObject({ code: "sync-lease-lost" });
    await expect(service.maintain(other, { jobId: id(9), leaseId: second.lease!.id })).rejects.toMatchObject({ code: "sync-active" });
  });
  it("preserves successful outcomes when retrying failed policies", async () => {
    const { service, caller, create } = setup(); await create();
    const claimed = await service.update(caller, id(9), { action: "claim" }); const leaseId = claimed.lease!.id;
    const other = { ...policy, id: id(11) };
    await service.update(caller, id(9), { action: "inventory", leaseId, policies: [policy, other] });
    await service.update(caller, id(9), { action: "policy", leaseId, policy: { ...policy, state: "completed", users: 2, exclusions: 1 } });
    await service.update(caller, id(9), { action: "policy", leaseId, policy: { ...other, state: "failed", error: "Graph HTTP 404" } });
    expect((await service.update(caller, id(9), { action: "finish", leaseId })).state).toBe("partially-failed");
    const retry = await service.update(caller, id(9), { action: "retry" });
    expect(retry.state).toBe("paused");
    expect(retry.policies.map((row) => row.state)).toEqual(["completed", "pending"]);
    expect(retry.policies[1].error).toBeUndefined();
  });
  it("rejects premature completion, duplicate inventory and single-policy scope mismatch", async () => {
    const { service, caller } = setup();
    await service.create(caller, { operationId: id(9), concurrency: 1, policyId: id(10) });
    const job = await service.update(caller, id(9), { action: "claim" }); const leaseId = job.lease!.id;
    await expect(service.update(caller, id(9), { action: "finish", leaseId })).rejects.toMatchObject({ code: "sync-incomplete" });
    await expect(service.update(caller, id(9), { action: "inventory", leaseId, policies: [] })).rejects.toMatchObject({ code: "sync-inventory" });
    await service.update(caller, id(9), { action: "inventory", leaseId, policies: [policy] });
    await expect(service.update(caller, id(9), { action: "inventory", leaseId, policies: [policy] })).rejects.toMatchObject({ code: "sync-inventory-exists" });
    await expect(service.update(caller, id(9), { action: "finish", leaseId })).rejects.toMatchObject({ code: "sync-incomplete" });
  });
  it("mirrors operational history without changing the new current execution", async () => {
    const { service, caller, create } = setup(); const old = await create();
    const cancelled = await service.update(caller, old.id, { action: "cancel" });
    await create(id(11));
    const mirrored = await service.update(caller, old.id, { action: "mirrored", revision: cancelled.revision });
    expect(mirrored.mirrorPending).toBe(false);
    expect((await service.current(caller))?.id).toBe(id(11));
  });
  it("pauses explicitly for access loss and requires a new lease on resume", async () => {
    const { service, caller, other, create } = setup(); await create();
    const first = await service.update(caller, id(9), { action: "claim" });
    const paused = await service.update(caller, id(9), { action: "pause", leaseId: first.lease!.id, accessRequired: true });
    expect(paused).toMatchObject({ state: "access-required", lease: undefined });
    expect((await service.update(other, id(9), { action: "claim" })).executorId).toBe(other.objectId);
  });
  it("locks terminal-job maintenance against new runs and rejects stale mirror revisions", async () => {
    const { service, caller, create, store } = setup(); await create();
    const cancelled = await service.update(caller, id(9), { action: "cancel" });
    const maintenance = await service.update(caller, id(9), { action: "claim-maintenance" });
    expect((await service.update(caller, id(9), { action: "mirrored", revision: cancelled.revision })).mirrorPending).toBe(true);
    await expect(create(id(11))).rejects.toMatchObject({ code: "sync-active" });
    expect(await service.maintain(caller, { jobId: id(9), leaseId: maintenance.lease!.id })).toEqual({ removed: 0 });
    await service.update(caller, id(9), { action: "release-maintenance", leaseId: maintenance.lease!.id });
    await create(id(11));
    expect(store.active?.job.id).toBe(id(11));
  });
});
describe("sync configuration and input bounds", () => {
  it("defaults to disabled, concurrency 3/8, and explicit runtime retention", () => {
    expect(syncSettings({})).toMatchObject({ enabled: false, defaultConcurrency: 3, maxConcurrency: 8, retentionDays: 30 });
    expect(() => syncSettings({ SYNC_SITE_ID: "only-one-setting" })).toThrow("all four");
    expect(() => syncSettings({ SYNC_DEFAULT_CONCURRENCY: "4", SYNC_MAX_CONCURRENCY: "3" })).toThrow("CONCURRENCY");
    expect(() => syncSettings({ SYNC_MAX_CONCURRENCY: "9" })).toThrow("CONCURRENCY");
  });
  it("rejects out-of-range concurrency and token-shaped/unknown fields", () => {
    for (const value of [0, 9, 1.5]) expect(syncCreateSchema.safeParse({ operationId: id(9), concurrency: value }).success).toBe(false);
    expect(syncCreateSchema.safeParse({ operationId: id(9), concurrency: 3, accessToken: "never-store" }).success).toBe(false);
    expect(syncUpdateSchema.safeParse({ action: "pause", leaseId: id(10), accessRequired: true, token: "never-store" }).success).toBe(false);
  });
});
