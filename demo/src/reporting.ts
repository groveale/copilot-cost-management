import type { ReportingSettings } from "../../api/src/reporting-contracts";
import { guid, object, text } from "./config";
import type { Fields } from "./details";
import { GraphClient } from "./graph";
import { aggregateObservations, parallelMap, type BalanceObservation, type PolicySnapshot } from "./sync-collector";

export interface SnapshotPolicy {
  id: string;
  displayName: string;
  activeSnapshotId: string | null;
  observedAt: string | null;
  startedAt: string | null;
  stale: boolean;
  configuration: { policy: Fields; services: Fields[]; groups: { id: string; displayName: string | null }[]; mappings: Fields[] } | null;
  coverage: { users: number; groups: number; exclusions: number; observations: number; unsupported: boolean; reason: string | null } | null;
  totals: PolicySnapshot["totals"] | null;
  jobId: string | null;
}
export interface SnapshotReading { policy: SnapshotPolicy; observations: BalanceObservation[] }
const optional = (value: unknown, label: string) => value === undefined || value === null || value === "" ? null : text(value, label);
function timestamp(value: unknown, label: string) {
  const raw = text(value, label);
  if (!/(Z|[+-]\d\d:\d\d)$/.test(raw) || !Number.isFinite(Date.parse(raw))) throw new Error(`Invalid ${label}.`);
  return new Date(raw).toISOString();
}
function quantity(value: unknown, label: string): number {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) throw new Error(`Invalid ${label}.`);
  return value;
}
function count(value: unknown, label: string) {
  const result = quantity(value, label);
  if (!Number.isSafeInteger(result)) throw new Error(`Invalid ${label}.`);
  return result;
}
function json(value: unknown, label: string): unknown {
  const raw = text(value, label);
  if (raw.length > 60000) throw new Error(`${label} exceeds the reporting schema limit.`);
  try { return JSON.parse(raw); }
  catch { throw new Error(`Invalid ${label} JSON. Ask an administrator to resync this policy.`); }
}
function array(value: unknown, label: string): unknown[] {
  if (!Array.isArray(value)) throw new Error(`${label} must be an array.`);
  return value;
}
function fields(value: unknown): Fields { return object(object(value, "SharePoint item").fields, "SharePoint fields"); }

export class SharePointReporting {
  private readonly tenantId: string;
  constructor(private readonly graph: GraphClient, readonly settings: ReportingSettings, tenantId: string) {
    this.tenantId = guid(tenantId, "Tenant ID");
    if (settings.enabled !== true) throw new Error("SharePoint reporting is not configured. Ask an administrator to configure sync.");
    if (typeof settings.isAdmin !== "boolean" || typeof settings.securityReady !== "boolean" ||
        !Array.isArray(settings.policyIds) || settings.policyIds.some((id) => guid(id, "Authorized policy ID") !== id)) {
      throw new Error("Invalid SharePoint reporting access configuration.");
    }
    if (!settings.isAdmin && !settings.securityReady) throw new Error("SharePoint owner permissions have not been verified. Ask an administrator to reconcile permissions.");
  }
  private root(list: string) {
    return `/v1.0/sites/${encodeURIComponent(text(this.settings.siteId, "Reporting site ID"))}/lists/${guid(list, "Reporting list ID")}/items`;
  }
  private async rows(list: string, field: "TenantId" | "PolicyKey" | "SnapshotKey", key: string, signal: AbortSignal) {
    const query = new URLSearchParams({ "$expand": "fields", "$filter": `fields/${field} eq '${key.replace(/'/g, "''")}'` });
    return (await this.graph.list(`${this.root(list)}?${query}`, signal)).map(fields);
  }
  private policy(row: Fields): SnapshotPolicy {
    const id = guid(row.PolicyId, "Snapshot policy ID");
    if (row.SchemaVersion !== 1 || row.TenantId !== this.tenantId || row.PolicyKey !== `${this.tenantId}:${id}`) {
      throw new Error("SharePoint policy schema or tenant/key does not match the configured destination.");
    }
    if (!this.settings.isAdmin && !this.settings.policyIds.includes(id)) throw new Error("SharePoint returned a policy outside your authorized workspace.");
    const active = optional(row.ActiveSnapshotId, "Active snapshot ID");
    const base = { id, displayName: optional(row.DisplayName ?? row.Title, "Policy name") ?? id,
      stale: row.Presence === "stale" };
    if (!active) return { ...base, activeSnapshotId: null, observedAt: null, startedAt: null,
      configuration: null, totals: null, coverage: null, jobId: null };
    guid(active, "Active snapshot ID");
    if (row.Presence !== "present" && row.Presence !== "stale") throw new Error("Invalid policy presence marker.");
    const configuration = object(json(row.Configuration, "Policy configuration"), "Policy configuration");
    const policy = object(configuration.policy, "Stored policy");
    if (guid(policy.id, "Stored policy ID") !== id) throw new Error("Stored configuration belongs to a different policy.");
    const coverage = object(json(row.Coverage, "Snapshot coverage"), "Snapshot coverage");
    if (typeof coverage.unsupported !== "boolean") throw new Error("Invalid snapshot coverage status.");
    const totals = coverage.unsupported ? null : array(json(row.Totals, "Snapshot totals"), "Snapshot totals").map((value) => {
      const total = object(value, "Snapshot total");
      return { serviceId: text(total.serviceId, "Total service"), unit: text(total.unit, "Total unit"),
        consumed: quantity(total.consumed, "Total consumption"), users: count(total.users, "Total users") };
    });
    return { ...base, activeSnapshotId: active, observedAt: timestamp(row.ObservedAt, "Snapshot collection time"),
      startedAt: timestamp(row.StartedAt, "Snapshot start time"), jobId: guid(row.JobId, "Snapshot job ID"),
      configuration: { policy, services: array(configuration.services, "Stored services").map((entry) => object(entry, "Stored service")),
        groups: array(configuration.groups, "Stored groups").map((entry) => {
          const group = object(entry, "Stored group");
          return { id: guid(group.id, "Stored group ID"), displayName: optional(group.displayName, "Stored group name") };
        }), mappings: array(configuration.mappings, "Stored mappings").map((entry) => object(entry, "Stored mapping")) },
      coverage: { users: count(coverage.users, "Discovered users"), groups: count(coverage.groups, "Discovered groups"),
        exclusions: count(coverage.exclusions, "Snapshot exclusions"), observations: count(coverage.observations, "Snapshot observations"),
        unsupported: coverage.unsupported, reason: optional(coverage.reason, "Coverage reason") }, totals };
  }
  async inventory(signal: AbortSignal): Promise<SnapshotPolicy[]> {
    const rows = this.settings.isAdmin ? await this.rows(this.settings.policyListId, "TenantId", this.tenantId, signal) :
      (await parallelMap(this.settings.policyIds, 3, (id) =>
        this.rows(this.settings.policyListId, "PolicyKey", `${this.tenantId}:${id}`, signal))).flat();
    const policies = rows.map((row) => this.policy(row));
    if (new Set(policies.map((row) => row.id)).size !== policies.length) throw new Error("SharePoint contains duplicate policy keys.");
    return policies.sort((a, b) => a.displayName.localeCompare(b.displayName));
  }
  private async master(policyId: string, signal: AbortSignal): Promise<SnapshotPolicy> {
    const rows = await this.rows(this.settings.policyListId, "PolicyKey", `${this.tenantId}:${policyId}`, signal);
    if (rows.length !== 1) throw new Error(rows.length ? "SharePoint contains duplicate policy keys." :
      "No SharePoint policy snapshot is available, or SharePoint has not granted you access. Ask an administrator to sync and verify permissions.");
    return this.policy(rows[0]);
  }
  async read(policyId: string, signal: AbortSignal): Promise<SnapshotReading> {
    const id = guid(policyId, "Policy ID");
    if (!this.settings.isAdmin && !this.settings.policyIds.includes(id)) throw new Error("You are not authorized to read this policy.");
    for (let attempt = 0; attempt < 2; attempt++) {
      signal.throwIfAborted();
      const policy = await this.master(id, signal);
      if (!policy.activeSnapshotId) return { policy, observations: [] };
      const key = `${this.tenantId}:${id}:${policy.activeSnapshotId}`;
      const rows = await this.rows(this.settings.balanceListId, "SnapshotKey", key, signal);
      const after = await this.master(id, signal);
      if (after.activeSnapshotId !== policy.activeSnapshotId || after.jobId !== policy.jobId) continue;
      const observations = rows.map((row) => this.observation(row, policy, key));
      if (observations.length !== policy.coverage!.observations) throw new Error("The published snapshot is incomplete or inaccessible. Refresh or ask an administrator to resync; missing balances are not zero.");
      const totals = aggregateObservations(observations);
      if (policy.totals !== null) {
        const sorted = (values: PolicySnapshot["totals"]) => [...values].sort((a, b) => `${a.serviceId}:${a.unit}`.localeCompare(`${b.serviceId}:${b.unit}`));
        const actual = sorted(totals), expected = sorted(policy.totals);
        const matches = actual.length === expected.length && actual.every((total, index) => {
          const stored = expected[index];
          // Parallel writes can change summation order for fractional quantities.
          const tolerance = Number.EPSILON * Math.max(1, total.consumed, stored.consumed) * Math.max(2, observations.length);
          return total.serviceId === stored.serviceId && total.unit === stored.unit && total.users === stored.users &&
            Math.abs(total.consumed - stored.consumed) <= tolerance;
        });
        if (!matches) throw new Error("Stored snapshot totals do not match its observations.");
      }
      return { policy: { ...policy, stale: after.stale }, observations };
    }
    throw new Error("The active SharePoint snapshot changed during reading. Refresh after sync finishes.");
  }
  private observation(row: Fields, policy: SnapshotPolicy, key: string): BalanceObservation {
    if (row.SchemaVersion !== 1 || row.TenantId !== this.tenantId || row.PolicyId !== policy.id ||
        row.PolicyKey !== `${this.tenantId}:${policy.id}` || row.SnapshotKey !== key ||
        row.SnapshotId !== policy.activeSnapshotId || row.JobId !== policy.jobId) throw new Error("Balance row belongs to another snapshot or schema.");
    const serviceId = optional(row.ServiceId, "Observation service"), reported = optional(row.ReportedPolicyId, "Reported policy ID");
    const reportedPolicyId = reported ? guid(reported, "Reported policy ID") : null;
    const match = row.PolicyMatch;
    if (match !== "matched" && match !== "mismatch" && match !== "no-balance") throw new Error("Invalid observation policy-match marker.");
    if (match === "no-balance" ? serviceId !== null || reportedPolicyId !== null :
      serviceId === null || reportedPolicyId === null || (match === "matched") !== (reportedPolicyId === policy.id)) throw new Error("Inconsistent observation policy attribution.");
    const remainingKind = row.RemainingKind;
    if (remainingKind !== "limited" && remainingKind !== "unlimited" && remainingKind !== "missing") throw new Error("Invalid remaining quantity marker.");
    const allowed = row.ConsumptionAllowed;
    if (allowed !== "true" && allowed !== "false" && allowed !== "not-reported") throw new Error("Invalid consumption-allowed marker.");
    if (match === "no-balance" && (row.ConsumedQuantity != null || row.Unit != null ||
        remainingKind !== "missing" || row.RemainingValue != null || allowed !== "not-reported")) throw new Error("A missing-balance row contains contradictory quantities or access flags.");
    return { userId: guid(row.UserId, "Snapshot user ID"), displayName: optional(row.DisplayName, "User display name"),
      username: optional(row.UserPrincipalName, "User principal name"),
      groups: array(json(row.Groups, "User groups"), "User groups").map((id) => guid(id, "User group ID")),
      serviceId, reportedPolicyId, match, consumed: match === "no-balance" ? null : quantity(row.ConsumedQuantity, "Consumption"),
      unit: match === "no-balance" ? null : text(row.Unit, "Balance unit"), remainingKind,
      remaining: remainingKind === "limited" ? quantity(row.RemainingValue, "Remaining value") : null,
      allowed: allowed === "not-reported" ? null : allowed === "true", reason: optional(row.AccessReason, "Access reason"),
      sourceTime: row.SourceTime ? timestamp(row.SourceTime, "Source timestamp") : null,
      observedAt: timestamp(row.ObservedAt, "Observation timestamp") };
  }
}
