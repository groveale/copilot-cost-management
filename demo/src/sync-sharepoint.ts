import type { SyncJob, SyncPolicy, SyncPolicyUpdate, SyncSettings } from "../../api/src/sync-contracts";
import type { Metadata } from "../../api/src/contracts";
import { guid, object, text } from "./config";
import { type Fields } from "./details";
import { GraphClient, GraphError } from "./graph";
import { parallelMap, type PolicySnapshot } from "./sync-collector";

interface Item { id: string; etag: string; fields: Fields }
export interface PolicyFence { item: Item; leaseId: string; key: string }
function listItem(value: unknown): Item {
  const row = object(value, "SharePoint item");
  const id = text(row.id, "SharePoint item ID");
  if (!/^\d+$/.test(id)) throw new Error("Invalid SharePoint list item ID.");
  return { id, etag: text(row.eTag ?? row["@odata.etag"], "SharePoint item ETag"), fields: object(row.fields, "SharePoint fields") };
}
function jsonField(value: unknown): string {
  const json = JSON.stringify(value);
  if (json.length > 60000) throw new Error("Reporting data exceeds a SharePoint text field limit. No data was truncated.");
  return json;
}
async function digest(value: unknown): Promise<string> {
  const hash = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(JSON.stringify(value)));
  return [...new Uint8Array(hash)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}
const policyKey = (tenantId: string, policyId: string) => `${guid(tenantId, "Tenant ID")}:${guid(policyId, "Policy ID")}`;
function sameField(name: string, actual: unknown, expected: unknown) {
  if (actual == null && expected == null) return true;
  if (["CreatedAt", "UpdatedAt", "FinishedAt", "StartedAt", "ObservedAt", "SourceTime", "LastInventoryAt"].includes(name) &&
      typeof actual === "string" && typeof expected === "string") {
    return Math.floor(Date.parse(actual) / 1000) === Math.floor(Date.parse(expected) / 1000);
  }
  return actual === expected;
}
export class SharePointSync {
  constructor(private readonly graph: GraphClient, readonly settings: SyncSettings, private readonly tenantId: string) {}
  private root(list: string) {
    return `/v1.0/sites/${encodeURIComponent(this.settings.siteId)}/lists/${guid(list, "List ID")}`;
  }
  private async items(list: string, field: string, value: string, signal: AbortSignal) {
    const query = new URLSearchParams({ "$expand": "fields", "$filter": `fields/${field} eq '${value.replace(/'/g, "''")}'` });
    return (await this.graph.list(`${this.root(list)}/items?${query}`, signal)).map(listItem);
  }
  private async find(list: string, field: string, value: string, signal: AbortSignal): Promise<Item | undefined> {
    const rows = await this.items(list, field, value, signal);
    if (rows.length > 1) throw new Error(`SharePoint ${field} is not unique. Repair the destination schema.`);
    return rows[0];
  }
  private async write(list: string, field: string, key: string, fields: Fields, signal: AbortSignal, existing?: Item): Promise<Item> {
    for (const [name, value] of Object.entries(fields)) {
      if (typeof value === "number" && !Number.isFinite(value)) throw new Error(`Invalid numeric reporting field ${name}.`);
      if (typeof value === "string" && value.length > (["Configuration", "Totals", "Coverage", "Groups", "Outcomes", "Actors"].includes(name) ? 60000 : 255)) {
        throw new Error(`Reporting field ${name} exceeds its SharePoint limit. No value was truncated.`);
      }
    }
    const path = `${this.root(list)}/items`;
    try {
      if (existing) await this.graph.patch(`${path}/${existing.id}/fields`, fields, signal, existing.etag);
      else await this.graph.post(path, { fields }, signal);
    } catch (error) {
      signal.throwIfAborted();
      // Deterministic keys permit verification after a lost write response.
      if (!(error instanceof TypeError) && !(error instanceof GraphError && [409, 412, 429, 500, 502, 503, 504].includes(error.status))) throw error;
      const observed = await this.find(list, field, key, signal);
      if (observed && Object.entries(fields).every(([name, value]) => sameField(name, observed.fields[name], value))) return observed;
      throw error;
    }
    const saved = await this.find(list, field, key, signal);
    if (!saved) throw new Error("SharePoint acknowledged a write but the saved item could not be verified.");
    const mismatches = Object.keys(fields).filter((name) => !sameField(name, saved.fields[name], fields[name]));
    if (mismatches.length) {
      throw new Error(`SharePoint item write could not be verified. Fields did not match: ${mismatches.join(", ")}. The write may have succeeded; refresh task status before resuming.`);
    }
    return saved;
  }
  async validate(signal: AbortSignal): Promise<string[]> {
    const requirements = [
      [this.settings.policyListId, "PolicyKey", "Title", "ActiveSnapshotId", "FenceId", "Configuration", "Totals", "Coverage", "TenantId",
        "PolicyId", "PolicyType", "JobId", "Presence", "StartedAt", "ObservedAt", "LastInventoryAt", "SchemaVersion"],
      [this.settings.balanceListId, "ObservationKey", "Title", "SnapshotKey", "PolicyKey", "SnapshotId", "UserId", "ServiceId",
        "TenantId", "PolicyId", "JobId", "UserPrincipalName", "Groups", "ReportedPolicyId", "PolicyMatch",
        "ConsumedQuantity", "Unit", "RemainingKind", "RemainingValue", "ConsumptionAllowed", "AccessReason", "SourceTime", "ObservedAt", "SchemaVersion"],
      [this.settings.jobListId, "JobId", "Title", "TenantId", "State", "UpdatedAt", "Revision", "Outcomes", "ScopePolicyId",
        "Concurrency", "InitiatorId", "ExecutorId", "Actors", "CreatedAt", "FinishedAt", "Phase", "PolicyCount", "CompletedCount", "FailedCount", "SchemaVersion"],
    ];
    for (const [list, key, ...required] of requirements) {
      const columns = (await this.graph.list(`${this.root(list)}/columns`, signal)).map((row) => object(row, "Column"));
      for (const name of [key, ...required]) if (!columns.some((column) => column.name === name)) throw new Error(`SharePoint list is missing column ${name}. Rerun setup.`);
      const unique = columns.find((column) => column.name === key)!;
      if (unique.indexed !== true || unique.enforceUniqueValues !== true) throw new Error(`SharePoint ${key} requires a unique index.`);
      for (const name of list === this.settings.balanceListId ? ["PolicyKey", "SnapshotKey"] : ["TenantId"]) {
        if (!columns.some((column) => column.name === name && column.indexed === true)) throw new Error(`SharePoint ${name} requires an index.`);
      }
    }
    return ["All three SharePoint list schemas and indexes: verified", "Destination writes: checked when mirroring the job and publishing each policy"];
  }
  async links(signal: AbortSignal): Promise<{ label: string; url: string }[]> {
    const links: { label: string; url: string }[] = [];
    for (const [label, list] of [["Spending policies", this.settings.policyListId], ["User balances", this.settings.balanceListId], ["Sync jobs", this.settings.jobListId]]) {
      const row = object(await this.graph.get(`${this.root(list)}?$select=webUrl`, signal), "SharePoint list");
      const url = new URL(text(row.webUrl, "SharePoint list URL"));
      if (url.protocol !== "https:" || url.hostname.toLowerCase() !== this.settings.siteId.split(",")[0].toLowerCase() ||
          url.username || url.password) throw new Error("SharePoint returned an unexpected reporting URL.");
      links.push({ label, url: url.href });
    }
    return links;
  }
  async mirror(job: SyncJob, signal: AbortSignal): Promise<void> {
    const existing = await this.find(this.settings.jobListId, "JobId", job.id, signal);
    if (existing && typeof existing.fields.Revision === "number" && existing.fields.Revision > job.revision) {
      throw new Error("A newer SharePoint job summary already exists. Refresh before reconciling.");
    }
    await this.write(this.settings.jobListId, "JobId", job.id, {
      Title: job.id, JobId: job.id, TenantId: job.tenantId, SchemaVersion: 1, Revision: job.revision, State: job.state,
      ScopePolicyId: job.policyId ?? null, Concurrency: job.concurrency, InitiatorId: job.initiatorId,
      ExecutorId: job.executorId ?? null, Actors: jsonField(job.actors), CreatedAt: job.createdAt, UpdatedAt: job.updatedAt,
      FinishedAt: job.finishedAt ?? null, Phase: job.phase, PolicyCount: job.policies.length,
      CompletedCount: job.policies.filter((policy) => ["completed", "excluded"].includes(policy.state)).length,
      FailedCount: job.policies.filter((policy) => policy.state === "failed").length,
      Outcomes: jsonField(job.policies.map(({ id, state, users, exclusions, error }) => ({ id, state, users, exclusions, error }))),
    }, signal, existing);
  }
  async recover(policy: SyncPolicy, job: SyncJob, signal: AbortSignal): Promise<SyncPolicyUpdate | undefined> {
    const item = await this.find(this.settings.policyListId, "PolicyKey", policyKey(job.tenantId, policy.id), signal);
    if (!item || item.fields.JobId !== job.id || !item.fields.ActiveSnapshotId) return undefined;
    const coverage = object(JSON.parse(text(item.fields.Coverage, "Coverage")), "Coverage");
    if (typeof coverage.users !== "number" || typeof coverage.exclusions !== "number" || typeof coverage.groups !== "number") {
      throw new Error("Published snapshot has invalid coverage metadata.");
    }
    return { ...policy, state: coverage.unsupported ? "excluded" : "completed", users: coverage.users,
      groups: coverage.groups, exclusions: coverage.exclusions };
  }
  async fence(policy: SyncPolicy, job: SyncJob, signal: AbortSignal): Promise<PolicyFence> {
    if (!job.lease) throw new Error("An execution lease is required.");
    const key = policyKey(job.tenantId, policy.id);
    let item = await this.find(this.settings.policyListId, "PolicyKey", key, signal);
    if (!item) item = await this.write(this.settings.policyListId, "PolicyKey", key, {
      Title: policy.displayName, PolicyKey: key, TenantId: job.tenantId, PolicyId: policy.id, SchemaVersion: 1,
    }, signal);
    item = await this.write(this.settings.policyListId, "PolicyKey", key, {
      FenceId: job.lease.id,
    }, signal, item);
    return { item, leaseId: job.lease.id, key };
  }
  async invalidate(job: SyncJob, signal: AbortSignal) {
    for (const policy of job.policies.filter((row) => row.state === "running")) {
      const key = policyKey(job.tenantId, policy.id);
      const item = await this.find(this.settings.policyListId, "PolicyKey", key, signal);
      if (item) await this.write(this.settings.policyListId, "PolicyKey", key, { FenceId: crypto.randomUUID() }, signal, item);
    }
  }
  async publish(fence: PolicyFence, snapshot: PolicySnapshot, policy: SyncPolicy, job: SyncJob,
    metadata: Metadata | null, signal: AbortSignal, checkLease: () => Promise<void>): Promise<void> {
    const snapshotId = crypto.randomUUID();
    const snapshotKey = `${fence.key}:${snapshotId}`;
    await parallelMap(snapshot.observations, job.concurrency, async (row) => {
      signal.throwIfAborted();
      const key = await digest([snapshotKey, row.userId, row.serviceId]);
      await this.write(this.settings.balanceListId, "ObservationKey", key, {
        Title: row.displayName ?? row.userId, ObservationKey: key, SnapshotKey: snapshotKey, PolicyKey: fence.key,
        TenantId: job.tenantId, PolicyId: policy.id, SnapshotId: snapshotId, JobId: job.id, UserId: row.userId,
        ServiceId: row.serviceId, UserPrincipalName: row.username, Groups: jsonField(row.groups),
        ReportedPolicyId: row.reportedPolicyId, PolicyMatch: row.match, ConsumedQuantity: row.consumed, Unit: row.unit,
        RemainingKind: row.remainingKind, RemainingValue: row.remaining,
        ConsumptionAllowed: row.allowed === null ? "not-reported" : String(row.allowed),
        AccessReason: row.reason, SourceTime: row.sourceTime, ObservedAt: row.observedAt, SchemaVersion: 1,
      }, signal);
    });
    const before = await this.find(this.settings.policyListId, "PolicyKey", fence.key, signal);
    if (!before || before.fields.FenceId !== fence.leaseId || before.etag !== fence.item.etag) {
      throw new Error("SharePoint publication fence changed. The previous snapshot was preserved.");
    }
    await checkLease();
    const current = await this.find(this.settings.policyListId, "PolicyKey", fence.key, signal);
    const protectedFields = ["Title", "PolicyKey", "TenantId", "PolicyId", "SchemaVersion", "FenceId",
      "PolicyType", "ActiveSnapshotId", "JobId", "Configuration", "Totals", "Coverage", "Presence", "StartedAt", "ObservedAt", "LastInventoryAt"];
    if (!current || !protectedFields.every((name) => sameField(name, current.fields[name], before.fields[name]))) {
      throw new Error("SharePoint publication fence or content changed while verifying permissions. The previous snapshot was preserved.");
    }
    const mappings = metadata?.budgets.flatMap((budget) => (["standard", "enhanced"] as const)
      .filter((tier) => budget.tiers[tier].policyId.toLowerCase() === policy.id)
      .map((tier) => ({ budgetId: budget.id, tier, businessArea: budget.businessArea, costCenter: budget.costCenter, owners: budget.owners }))) ?? [];
    await this.write(this.settings.policyListId, "PolicyKey", fence.key, {
      Title: text(snapshot.policy.displayName, "Policy name"),
      PolicyType: typeof snapshot.policy["@odata.type"] === "string" ? snapshot.policy["@odata.type"] : policy.type,
      ActiveSnapshotId: snapshotId,
      JobId: job.id, Configuration: jsonField({ policy: snapshot.policy, services: snapshot.services, groups: snapshot.groups, mappings }),
      Totals: snapshot.unsupported ? null : jsonField(snapshot.totals),
      Coverage: jsonField({ users: snapshot.users, groups: snapshot.groups.length, exclusions: snapshot.exclusions,
        observations: snapshot.observations.length,
        observedUsers: new Set(snapshot.observations.filter((row) => row.serviceId !== null).map((row) => row.userId)).size,
        matchedUsers: new Set(snapshot.observations.filter((row) => row.match === "matched").map((row) => row.userId)).size,
        matchedObservations: snapshot.observations.filter((row) => row.match === "matched").length,
        missingBalanceUsers: snapshot.observations.filter((row) => row.match === "no-balance").length,
        mismatchedObservations: snapshot.observations.filter((row) => row.match === "mismatch").length,
        unsupported: snapshot.unsupported, reason: snapshot.unsupported ? "Policy type has no supported group-scoped collection" : null }),
      Presence: "present", StartedAt: snapshot.startedAt, ObservedAt: snapshot.completedAt,
    }, signal, current);
  }
  async cleanupPolicy(policyId: string, signal: AbortSignal, guard: () => Promise<void> = async () => {}) {
    const key = policyKey(this.tenantId, policyId);
    const item = await this.find(this.settings.policyListId, "PolicyKey", key, signal);
    if (!item) return;
    const rows = await this.items(this.settings.balanceListId, "PolicyKey", key, signal);
    for (const row of rows) {
      if (row.fields.SnapshotId === item.fields.ActiveSnapshotId) continue;
      await guard();
      await this.graph.deleteListItem(`${this.root(this.settings.balanceListId)}/items/${row.id}`, signal, row.etag);
    }
  }
  async markAbsent(policies: SyncPolicy[], signal: AbortSignal) {
    const present = new Set(policies.map((row) => row.id));
    const rows = await this.items(this.settings.policyListId, "TenantId", this.tenantId, signal);
    for (const row of rows) if (!present.has(String(row.fields.PolicyId))) {
      await this.write(this.settings.policyListId, "PolicyKey", text(row.fields.PolicyKey, "Policy key"),
        { Presence: "stale", LastInventoryAt: new Date().toISOString() }, signal, row);
    }
  }
  async pruneJobs(signal: AbortSignal, keepJobId: string, guard: () => Promise<void> = async () => {}) {
    const cutoff = Date.now() - this.settings.retentionDays * 86400000;
    const rows = await this.items(this.settings.jobListId, "TenantId", this.tenantId, signal);
    for (const row of rows) if (row.fields.JobId !== keepJobId && ["completed", "completed-with-exclusions", "partially-failed", "failed", "cancelled"].includes(String(row.fields.State)) &&
        Date.parse(String(row.fields.UpdatedAt)) < cutoff) {
      await guard();
      await this.graph.deleteListItem(`${this.root(this.settings.jobListId)}/items/${row.id}`, signal, row.etag);
    }
  }
}
