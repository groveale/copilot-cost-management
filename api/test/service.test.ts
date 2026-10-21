import { beforeEach, describe, expect, it } from "vitest";
import type { Caller, Metadata, Page, RequestRecord, Submission } from "../src/contracts.js";
import { decisionSchema, fulfilmentUpdateSchema, seedSchema, submissionSchema } from "../src/model.js";
import { RequestService, type Store } from "../src/service.js";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const requester: Caller = { tenantId: id(1), objectId: id(2) };
const owner: Caller = { tenantId: id(1), objectId: id(3) };
const input: Submission = { operationId: id(9), justification: "Additional capacity for our contract review project.", grantKind: "billingPeriod" };
const metadata: Metadata = {
  version: 1, tenantId: id(1), serviceId: "cowork",
  budgets: [{
    id: "legal", displayName: "Legal", businessArea: "Legal", costCenter: "example", owners: [owner.objectId],
    tiers: { standard: { policyId: id(4), groupId: id(5) }, enhanced: { policyId: id(6), groupId: id(7) } },
  }],
};
class MemoryStore implements Store {
  config = structuredClone(metadata);
  index = new Map([[requester.objectId, "legal"]]);
  records = new Map<string, RequestRecord>();
  async metadata() { return structuredClone(this.config); }
  async routing(caller: Caller) { return this.index.get(caller.objectId) ?? null; }
  async get(tenant: string, key: string) {
    const record = this.records.get(key);
    return record?.tenantId === tenant ? structuredClone(record) : null;
  }
  async create(record: RequestRecord) {
    if (this.records.has(record.id)) return false;
    this.records.set(record.id, { ...structuredClone(record), etag: '"1"' });
    return true;
  }
  async replace(record: RequestRecord, etag: string) {
    if (this.records.get(record.id)?.etag !== etag) return false;
    this.records.set(record.id, { ...structuredClone(record), etag: `"${Number(etag.slice(1, -1)) + 1}"` });
    return true;
  }
  async list(tenant: string, requesterId?: string, budgets?: string[]): Promise<Page<RequestRecord>> {
    return { items: [...this.records.values()].filter((r) => r.tenantId === tenant &&
      (!requesterId || r.requesterId === requesterId) && (!budgets || budgets.includes(r.budgetId))).map((r) => structuredClone(r)) };
  }
}
let store: MemoryStore;
let now: Date;
let service: RequestService;
beforeEach(() => {
  store = new MemoryStore(); now = new Date("2026-12-31T23:59:59Z");
  service = new RequestService(store, () => now);
});
describe("persistent request rules", () => {
  it("captures exclusive UTC month end on the server and exposes no owner metadata to requesters", async () => {
    const saved = await service.submit(requester, input);
    expect(saved.expiresAt).toBe("2027-01-01T00:00:00.000Z");
    expect(saved.requesterId).toBe(requester.objectId);
    expect((await service.context(requester)).ownedBudgets).toEqual([]);
    expect((await service.context(owner)).ownedBudgets).toHaveLength(1);
  });
  it("has no expiry for permanent requests", async () => {
    expect((await service.submit(requester, { ...input, grantKind: "permanent" })).expiresAt).toBeNull();
  });
  it("replays identical submissions, including concurrent retries, and rejects conflicting reuse", async () => {
    const [first, retry] = await Promise.all([service.submit(requester, input), service.submit(requester, input)]);
    expect(first).toEqual(retry); expect(store.records.size).toBe(1);
    expect(await service.submit(requester, input)).toEqual(first);
    await expect(service.submit(requester, { ...input, grantKind: "permanent" })).rejects.toMatchObject({ code: "idempotency-conflict" });
  });
  it("scopes idempotency and history to tenant and authenticated requester", async () => {
    await service.submit(requester, input);
    expect((await service.history(owner)).items).toEqual([]);
    expect((await service.history({ ...requester, tenantId: id(99) })).items).toEqual([]);
    expect((await service.history(requester)).items).toHaveLength(1);
  });
  it("rejects missing, dangling and ownerless routes without creating an orphan", async () => {
    for (const budget of [undefined, "missing", "legal"]) {
      store.index.clear(); if (budget) store.index.set(requester.objectId, budget);
      store.config.budgets[0].owners = [];
      await expect(service.submit(requester, input)).rejects.toMatchObject({ code: "needs-routing" });
      expect((await service.context(requester)).route.state).toBe("needs-routing");
    }
    expect(store.records.size).toBe(0);
  });
  it("allows an owner to request and decide their own request for the MVP, but rejects non-owners", async () => {
    store.config.budgets[0].owners = [requester.objectId];
    const saved = await service.submit(requester, input);
    await expect(service.decide({ ...owner, objectId: id(80) }, saved.id, { status: "approved", etag: saved.etag })).rejects.toMatchObject({ status: 404 });
    expect((await service.context(requester)).route.state).toBe("ready");
    expect((await service.inbox(requester)).items).toHaveLength(1);
    const decided = await service.decide(requester, saved.id, { status: "approved", etag: saved.etag });
    expect(decided.status).toBe("approved");
    expect(decided.decisionBy).toBe(requester.objectId);
    expect((await service.history(requester)).items[0]).toEqual(decided);
  });
  it("still enforces routing changes and expiry for self-approval", async () => {
    store.config.budgets[0].owners = [requester.objectId];
    const saved = await service.submit(requester, input);
    store.index.delete(requester.objectId);
    await expect(service.decide(requester, saved.id, { status: "approved", etag: saved.etag })).rejects.toMatchObject({ code: "needs-routing" });
    store.index.set(requester.objectId, "legal");
    now = new Date(saved.expiresAt!);
    await expect(service.decide(requester, saved.id, { status: "approved", etag: saved.etag })).rejects.toMatchObject({ code: "expired" });
  });
  it("rechecks current ownership and requester routing", async () => {
    const saved = await service.submit(requester, input);
    store.config.budgets[0].owners = [id(80)];
    await expect(service.inbox(owner)).rejects.toMatchObject({ status: 403 });
    await expect(service.decide(owner, saved.id, { status: "approved", etag: saved.etag })).rejects.toMatchObject({ status: 404 });
    store.config.budgets[0].owners = [owner.objectId];
    store.index.delete(requester.objectId);
    await expect(service.decide(owner, saved.id, { status: "approved", etag: saved.etag })).rejects.toMatchObject({ code: "needs-routing" });
  });
  it("blocks decisions at the exact expiry boundary", async () => {
    const saved = await service.submit(requester, input);
    now = new Date(saved.expiresAt!);
    await expect(service.decide(owner, saved.id, { status: "approved", etag: saved.etag })).rejects.toMatchObject({ code: "expired" });
  });
  it("permits only one concurrent decision and stores its audit fields", async () => {
    const saved = await service.submit(requester, input);
    const results = await Promise.allSettled([
      service.decide(owner, saved.id, { status: "approved", etag: saved.etag }),
      service.decide(owner, saved.id, { status: "rejected", reason: "Capacity not needed.", etag: saved.etag }),
    ]);
    expect(results.filter((r) => r.status === "fulfilled")).toHaveLength(1);
    const record = (await service.history(requester)).items[0];
    expect(record.decisionBy).toBe(owner.objectId); expect(record.decisionAt).toBe(now.toISOString());
    expect(record.status).toBe("approved");
    await expect(service.decide(owner, saved.id, { status: "approved", etag: record.etag })).rejects.toMatchObject({ code: "already-decided" });
  });
  it("surfaces storage failures instead of claiming submission succeeded", async () => {
    store.create = async () => { throw new Error("Identity access denied"); };
    await expect(service.submit(requester, input)).rejects.toThrow("Identity access denied");
    expect(store.records.size).toBe(0);
  });
});
describe("strict input boundary", () => {
  it("rejects forged routing, identity and approval fields", () => {
    for (const field of ["requesterId", "tenantId", "budgetId", "approverId", "status", "expiresAt"]) {
      expect(submissionSchema.safeParse({ ...input, [field]: id(40) }).success).toBe(false);
    }
  });

  describe("durable Standard-to-Enhanced fulfilment", () => {
    async function approved(permanent = true) {
      const saved = await service.submit(requester, { ...input, ...(permanent ? { grantKind: "permanent" as const } : {}) });
      return service.decide(owner, saved.id, { status: "approved", etag: saved.etag });
    }
    async function claim(record: RequestRecord) {
      return service.fulfil(owner, record.id, { action: "claim", etag: record.etag });
    }
    it.each([true, false])("freezes the server's policy/group mapping for permanent=%s", async (permanent) => {
      const record = await approved(permanent);
      expect(record.fulfilment).toEqual({ state: "pending", ...metadata.budgets[0].tiers });
      expect((await service.history(requester)).items[0]).toEqual(record);
    });
    it("does not allow a requester, another tenant or a non-owner to apply an approval", async () => {
      const record = await approved();
      for (const caller of [requester, { ...owner, tenantId: id(80) }, { ...owner, objectId: id(80) }]) {
        await expect(service.fulfil(caller, record.id, { action: "claim", etag: record.etag })).rejects.toMatchObject({ status: 404 });
      }
    });
    it("does not apply pending or rejected requests", async () => {
      const record = await service.submit(requester, input);
      await expect(claim(record)).rejects.toMatchObject({ code: "not-approved" });
      const rejected = await service.decide(owner, record.id, { status: "rejected", reason: "No additional capacity needed.", etag: record.etag });
      await expect(claim(rejected)).rejects.toMatchObject({ code: "not-approved" });
    });
    it("allows explicit fulfilment of a legacy approved record without changing its decision", async () => {
      const record = await approved();
      delete record.fulfilment;
      store.records.set(record.id, structuredClone(record));
      const started = await claim(record);
      expect(started.fulfilment).toMatchObject({ state: "applying", standard: metadata.budgets[0].tiers.standard });
      expect(started.decisionAt).toBe(record.decisionAt);
    });
    it("grants only one concurrent lease and rejects stale updates", async () => {
      const record = await approved();
      const outcomes = await Promise.allSettled([claim(record), claim(record)]);
      expect(outcomes.filter((value) => value.status === "fulfilled")).toHaveLength(1);
      const current = (await service.history(requester)).items[0];
      await expect(claim(current)).rejects.toMatchObject({ code: "fulfilment-busy" });
      await expect(service.fulfil(owner, current.id, { action: "complete", etag: record.etag, leaseId: current.fulfilment!.lease!.id, verified: true }))
        .rejects.toMatchObject({ code: "fulfilment-conflict" });
    });
    it("renews before writes, verifies completion and preserves an idempotent completed result", async () => {
      let record = await claim(await approved());
      const leaseId = record.fulfilment!.lease!.id;
      const original = record.fulfilment!.lease!.until;
      now = new Date(now.getTime() + 60_000);
      record = await service.fulfil(owner, record.id, { action: "heartbeat", etag: record.etag, leaseId });
      expect(Date.parse(record.fulfilment!.lease!.until)).toBeGreaterThan(Date.parse(original));
      await expect(service.fulfil(owner, record.id, { action: "complete", etag: record.etag, leaseId })).rejects.toMatchObject({ code: "verification-required" });
      record = await service.fulfil(owner, record.id, { action: "complete", etag: record.etag, leaseId, verified: true });
      expect(record.fulfilment).toMatchObject({ state: "applied", appliedBy: owner.objectId, appliedAt: now.toISOString() });
      expect(record.fulfilment?.lease).toBeUndefined();
      expect(await service.fulfil(owner, record.id, { action: "claim", etag: "stale" })).toEqual(record);
    });
    it("permits release after a failure and retries with a new lease", async () => {
      let record = await claim(await approved());
      const oldLease = record.fulfilment!.lease!.id;
      record = await service.fulfil(owner, record.id, { action: "release", etag: record.etag, leaseId: oldLease });
      expect(record.fulfilment?.state).toBe("needs-attention");
      record = await claim(record);
      expect(record.fulfilment?.lease?.id).not.toBe(oldLease);
      await expect(service.fulfil(owner, record.id, { action: "release", etag: record.etag, leaseId: oldLease })).rejects.toMatchObject({ code: "fulfilment-lease" });
    });
    it("recovers a closed-browser lease only after expiry and rejects its old worker", async () => {
      let record = await claim(await approved());
      const oldLease = record.fulfilment!.lease!.id;
      now = new Date(record.fulfilment!.lease!.until);
      record = await claim(record);
      expect(record.fulfilment?.lease?.id).not.toBe(oldLease);
      await expect(service.fulfil(owner, record.id, { action: "complete", etag: record.etag, leaseId: oldLease, verified: true }))
        .rejects.toMatchObject({ code: "fulfilment-lease" });
    });
    it("rechecks ownership, requester routing and frozen mappings before a write heartbeat", async () => {
      const record = await claim(await approved());
      const update = { action: "heartbeat" as const, etag: record.etag, leaseId: record.fulfilment!.lease!.id };
      store.config.budgets[0].owners = [id(80)];
      await expect(service.fulfil(owner, record.id, update)).rejects.toMatchObject({ status: 404 });
      store.config.budgets[0].owners = [owner.objectId];
      store.index.delete(requester.objectId);
      await expect(service.fulfil(owner, record.id, update)).rejects.toMatchObject({ code: "needs-routing" });
      store.index.set(requester.objectId, "legal");
      store.config.budgets[0].tiers.enhanced.groupId = id(88);
      await expect(service.fulfil(owner, record.id, update)).rejects.toMatchObject({ code: "fulfilment-mapping-changed" });
    });
    it("prevents new claims and writes at the temporary expiry boundary, while retaining reviewable partial state", async () => {
      let record = await claim(await approved(false));
      const leaseId = record.fulfilment!.lease!.id;
      now = new Date(record.expiresAt!);
      await expect(service.fulfil(owner, record.id, { action: "heartbeat", etag: record.etag, leaseId })).rejects.toMatchObject({ code: "expired" });
      record = await service.fulfil(owner, record.id, { action: "release", etag: record.etag, leaseId });
      expect(record.fulfilment?.state).toBe("needs-attention");
      await expect(claim(record)).rejects.toMatchObject({ code: "expired" });
    });
    it("can record an already-verified move just after the temporary boundary without authorizing further writes", async () => {
      const record = await claim(await approved(false));
      now = new Date(record.expiresAt!);
      const completed = await service.fulfil(owner, record.id, {
        action: "complete", etag: record.etag, leaseId: record.fulfilment!.lease!.id, verified: true,
      });
      expect(completed.fulfilment?.state).toBe("applied");
      expect(completed.expiresAt).toBe(record.expiresAt);
    });
    it("accepts only specific ETags, current lease IDs and explicit completion verification", () => {
      expect(fulfilmentUpdateSchema.safeParse({ action: "claim", etag: "e" }).success).toBe(true);
      for (const update of [
        { action: "claim", etag: "*" }, { action: "complete", etag: "e" },
        { action: "complete", etag: "e", leaseId: id(40) },
        { action: "claim", etag: "e", enhancedGroupId: id(40) },
        { action: "claim", etag: "e", requesterId: id(40) },
      ]) expect(fulfilmentUpdateSchema.safeParse(update).success).toBe(false);
    });
  });
  it("enforces exact justification and rejection thresholds", () => {
    for (const length of [19, 2001]) expect(submissionSchema.safeParse({ ...input, justification: "x".repeat(length) }).success).toBe(false);
    for (const length of [20, 2000]) expect(submissionSchema.safeParse({ ...input, justification: "x".repeat(length) }).success).toBe(true);
    for (const length of [0, 9, 1001]) expect(decisionSchema.safeParse({ status: "rejected", etag: "e", reason: "x".repeat(length) }).success).toBe(false);
    for (const length of [10, 1000]) expect(decisionSchema.safeParse({ status: "rejected", etag: "e", reason: "x".repeat(length) }).success).toBe(true);
    expect(decisionSchema.safeParse({ status: "approved", etag: "*" }).success).toBe(false);
  });
  it("blocks ambiguous index records and unknown budget mappings at import", () => {
    for (const requesterIndex of [
      [{ objectId: requester.objectId, budgetId: "missing" }],
      [{ objectId: requester.objectId, budgetId: "legal" }, { objectId: requester.objectId, budgetId: "legal" }],
    ]) expect(seedSchema.safeParse({ metadata, requesterIndex }).success).toBe(false);
  });
});
