import { describe, expect, it } from "vitest";
import type { SyncJob, SyncSettings } from "../../api/src/sync-contracts";
import { GraphClient, GraphError } from "./graph";
import { SharePointSync } from "./sync-sharepoint";
import { observationsFor, aggregateObservations, type PolicySnapshot } from "./sync-collector";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const settings: SyncSettings = { enabled: true, siteId: `example.sharepoint.com,${id(2)},${id(3)}`, policyListId: id(4),
  balanceListId: id(5), jobListId: id(6), defaultConcurrency: 3, maxConcurrency: 8, retentionDays: 30 };
const policy = { id: id(7), displayName: "Policy", type: "#microsoft.graph.selectedGroupsSpendingPolicy", state: "pending" as const };
const job: SyncJob = { version: 1, revision: 1, id: id(8), tenantId: id(1), initiatorId: id(9), executorId: id(9), actors: [id(9)],
  concurrency: 3, state: "running", phase: "Collecting", createdAt: "2026-10-28T13:00:00.123Z", updatedAt: "2026-10-28T13:00:00.123Z",
  inventoryComplete: true, policies: [policy], mirrorPending: true, lease: { id: id(10), expiresAt: "2026-10-28T14:00:00Z" } };
interface Row { id: string; eTag: string; fields: Record<string, unknown> }
class Lists extends GraphClient {
  rows = new Map<string, Row[]>();
  counter = 0;
  failObservation = false;
  loseResponse = false;
  constructor() { super(async () => "unused"); }
  private listId(path: string) { return path.match(/\/lists\/([a-f0-9-]+)/)![1]; }
  private data(path: string) { const key = this.listId(path); if (!this.rows.has(key)) this.rows.set(key, []); return this.rows.get(key)!; }
  override async list(path: string) {
    const url = new URL(path, "https://graph.microsoft.com");
    const filter = url.searchParams.get("$filter")!;
    const [, name, raw] = /fields\/(\w+) eq '(.*)'/.exec(filter)!;
    return structuredClone(this.data(path).filter((row) => row.fields[name] === raw.replace(/''/g, "'")));
  }
  override async post(path: string, body: unknown) {
    const fields = (body as { fields: Record<string, unknown> }).fields;
    if (this.failObservation && fields.ObservationKey) throw new GraphError("Write failed", 503, "Unavailable", null);
    const key = fields.ObservationKey ?? fields.PolicyKey ?? fields.JobId;
    if (this.data(path).some((row) => (row.fields.ObservationKey ?? row.fields.PolicyKey ?? row.fields.JobId) === key)) {
      throw new GraphError("Duplicate key", 409, "Conflict", null);
    }
    const row = { id: String(++this.counter), eTag: `"${this.counter}:1"`, fields: structuredClone(fields) };
    this.data(path).push(row);
    if (this.loseResponse && fields.ObservationKey) { this.loseResponse = false; throw new TypeError("Connection lost after committing"); }
    return { data: row, etag: row.eTag };
  }
  override async patch(path: string, body: unknown, _signal: AbortSignal, etag: string) {
    const itemId = /\/items\/(\d+)/.exec(path)![1];
    const row = this.data(path).find((row) => row.id === itemId)!;
    if (row.eTag !== etag) throw new GraphError("Changed", 412, "PreconditionFailed", null);
    Object.assign(row.fields, body); row.eTag += "x";
    return { data: row, etag: row.eTag };
  }
  override async deleteListItem(path: string, _signal: AbortSignal, etag: string) {
    const data = this.data(path), itemId = /\/items\/(\d+)/.exec(path)![1];
    const index = data.findIndex((row) => row.id === itemId && row.eTag === etag);
    if (index < 0) throw new Error("Delete ETag mismatch.");
    data.splice(index, 1); return { data: null, etag: null };
  }
}
function setup() {
  const graph = new Lists(), reporting = new SharePointSync(graph, settings, job.tenantId);
  const observations = observationsFor({ id: id(11), displayName: "Private user", username: "private@example.test", groups: [id(12)] },
    policy.id, [{ serviceId: "cowork", spendingPolicyId: policy.id, consumedQuantity: 7, unit: "credit" }]);
  const snapshot: PolicySnapshot = { policy: { id: policy.id, displayName: policy.displayName }, services: [], groups: [{ id: id(12), displayName: null }],
    observations, totals: aggregateObservations(observations), users: 1, exclusions: 0, unsupported: false,
    startedAt: job.createdAt, completedAt: job.updatedAt };
  return { graph, reporting, snapshot, signal: new AbortController().signal };
}
describe("SharePoint generation publication", () => {
  it("publishes a complete generation, recovers a lost job acknowledgement, and removes only superseded rows", async () => {
    const { graph, reporting, snapshot, signal } = setup();
    const fence = await reporting.fence(policy, job, signal);
    await reporting.publish(fence, snapshot, policy, job, null, signal, async () => {});
    expect(await reporting.recover(policy, job, signal)).toMatchObject({ state: "completed", users: 1 });
    const first = graph.rows.get(settings.policyListId)![0].fields.ActiveSnapshotId;
    const next = { ...job, id: id(20), lease: { ...job.lease!, id: id(21) } };
    await reporting.publish(await reporting.fence(policy, next, signal), snapshot, policy, next, null, signal, async () => {});
    expect(graph.rows.get(settings.balanceListId)).toHaveLength(2);
    expect(graph.rows.get(settings.policyListId)![0].fields.ActiveSnapshotId).not.toBe(first);
    await reporting.cleanupPolicy(policy.id, signal);
    expect(graph.rows.get(settings.balanceListId)).toHaveLength(1);
    expect(graph.rows.get(settings.balanceListId)![0].fields.JobId).toBe(next.id);
  });
  it("preserves the prior successful snapshot when a new observation write fails", async () => {
    const { graph, reporting, snapshot, signal } = setup();
    await reporting.publish(await reporting.fence(policy, job, signal), snapshot, policy, job, null, signal, async () => {});
    const before = graph.rows.get(settings.policyListId)![0].fields.ActiveSnapshotId;
    const next = { ...job, id: id(20) }; graph.failObservation = true;
    await expect(reporting.publish(await reporting.fence(policy, next, signal), snapshot, policy, next, null, signal, async () => {})).rejects.toThrow("Write failed");
    expect(graph.rows.get(settings.policyListId)![0].fields.ActiveSnapshotId).toBe(before);
  });
  it("rejects stale executor publication after another admin fences the policy", async () => {
    const { graph, reporting, snapshot, signal } = setup();
    const first = await reporting.fence(policy, job, signal);
    await reporting.fence(policy, { ...job, lease: { ...job.lease!, id: id(22) } }, signal);
    await expect(reporting.publish(first, snapshot, policy, job, null, signal, async () => {})).rejects.toThrow("fence changed");
    expect(graph.rows.get(settings.policyListId)![0].fields.ActiveSnapshotId).toBeUndefined();
  });
  it("reconciles an uncertain create by exact deterministic-key readback, without adding duplicate balances", async () => {
    const { graph, reporting, snapshot, signal } = setup(); graph.loseResponse = true;
    await reporting.publish(await reporting.fence(policy, job, signal), snapshot, policy, job, null, signal, async () => {});
    expect(graph.rows.get(settings.balanceListId)).toHaveLength(1);
  });
  it("does not publish when the API lease recheck fails", async () => {
    const { graph, reporting, snapshot, signal } = setup();
    await expect(reporting.publish(await reporting.fence(policy, job, signal), snapshot, policy, job, null, signal,
      async () => { throw new Error("Lease lost"); })).rejects.toThrow("Lease lost");
    expect(graph.rows.get(settings.policyListId)![0].fields.ActiveSnapshotId).toBeUndefined();
  });
  it("uses a fresh conditional-write ETag after ACL-only verification without weakening the publication fence", async () => {
    const { graph, reporting, snapshot, signal } = setup();
    const fence = await reporting.fence(policy, job, signal);
    await reporting.publish(fence, snapshot, policy, job, null, signal, async () => {
      graph.rows.get(settings.policyListId)![0].eTag += "acl-change";
    });
    expect(await reporting.recover(policy, job, signal)).toMatchObject({ state: "completed" });
  });
  it.each(["FenceId", "ActiveSnapshotId", "Configuration"])("rejects a %s change during permission verification", async (field) => {
    const { graph, reporting, snapshot, signal } = setup();
    const fence = await reporting.fence(policy, job, signal);
    await expect(reporting.publish(fence, snapshot, policy, job, null, signal, async () => {
      const row = graph.rows.get(settings.policyListId)![0];
      row.fields[field] = field === "Configuration" ? "{}" : id(90);
      row.eTag += "concurrent-change";
    })).rejects.toThrow("fence or content changed");
    expect(graph.rows.get(settings.policyListId)![0].fields.ActiveSnapshotId).toBe(field === "ActiveSnapshotId" ? id(90) : undefined);
  });
  it("keeps operational job history free of user balance and authentication data", async () => {
    const { graph, reporting, signal } = setup();
    await reporting.mirror(job, signal);
    const mirror = JSON.stringify(graph.rows.get(settings.jobListId));
    expect(mirror).not.toContain("private@example.test");
    expect(mirror).not.toContain("consumedQuantity");
    expect(mirror).not.toContain(job.lease!.id);
    expect(graph.rows.get(settings.jobListId)![0].fields.Revision).toBe(1);
    await expect(reporting.mirror({ ...job, revision: 0 }, signal)).rejects.toThrow("newer");
  });
  it("marks absent policies stale without deleting their last snapshot", async () => {
    const { graph, reporting, snapshot, signal } = setup();
    await reporting.publish(await reporting.fence(policy, job, signal), snapshot, policy, job, null, signal, async () => {});
    const before = graph.rows.get(settings.policyListId)![0].fields.ActiveSnapshotId;
    await reporting.markAbsent([], signal);
    expect(graph.rows.get(settings.policyListId)![0].fields).toMatchObject({ Presence: "stale", ActiveSnapshotId: before });
    expect(graph.rows.get(settings.balanceListId)).toHaveLength(1);
  });
  it("writes missing balances without inventing zero consumption or a false access decision", async () => {
    const { graph, reporting, snapshot, signal } = setup();
    snapshot.observations = observationsFor({ id: id(11), displayName: null, username: null, groups: [id(12)] }, policy.id, []);
    snapshot.totals = []; snapshot.exclusions = 1;
    await reporting.publish(await reporting.fence(policy, job, signal), snapshot, policy, job, null, signal, async () => {});
    expect(graph.rows.get(settings.balanceListId)![0].fields).toMatchObject({
      ServiceId: null, PolicyMatch: "no-balance", ConsumedQuantity: null, ConsumptionAllowed: "not-reported",
    });
  });
  it("publishes unsupported-policy inventory with null consumption, never a fabricated zero total", async () => {
    const { graph, reporting, snapshot, signal } = setup();
    snapshot.unsupported = true; snapshot.observations = []; snapshot.totals = []; snapshot.users = 0; snapshot.exclusions = 1;
    await reporting.publish(await reporting.fence(policy, job, signal), snapshot, policy, job, null, signal, async () => {});
    expect(graph.rows.get(settings.policyListId)![0].fields.Totals).toBeNull();
    expect(await reporting.recover(policy, job, signal)).toMatchObject({ state: "excluded", exclusions: 1 });
    expect(graph.rows.get(settings.balanceListId)).toBeUndefined();
  });
});
