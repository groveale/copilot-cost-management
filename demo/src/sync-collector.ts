import type { SyncPolicy } from "../../api/src/sync-contracts";
import { guid, object, text } from "./config";
import { serviceRows, type Fields } from "./details";
import { GraphClient } from "./graph";

export const costBase = "/beta/copilot/costManagement";
export const selectedPolicyType = "#microsoft.graph.selectedGroupsSpendingPolicy";
export interface SyncUser { id: string; displayName: string | null; username: string | null; groups: string[] }
export interface BalanceObservation {
  userId: string;
  displayName: string | null;
  username: string | null;
  groups: string[];
  serviceId: string | null;
  reportedPolicyId: string | null;
  match: "matched" | "mismatch" | "no-balance";
  consumed: number | null;
  unit: string | null;
  remainingKind: "limited" | "unlimited" | "missing";
  remaining: number | null;
  allowed: boolean | null;
  reason: string | null;
  sourceTime: string | null;
  observedAt: string;
}
export interface PolicySnapshot {
  policy: Fields;
  services: Fields[];
  groups: { id: string; displayName: string | null }[];
  observations: BalanceObservation[];
  totals: { serviceId: string; unit: string; consumed: number; users: number }[];
  users: number;
  exclusions: number;
  unsupported: boolean;
  startedAt: string;
  completedAt: string;
}
const optionalText = (value: unknown) => typeof value === "string" && value.trim() ? value.trim() : null;
function quantity(value: unknown, label: string) {
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) throw new Error(`Invalid ${label}.`);
  return value;
}
export function policyDescriptor(value: unknown): SyncPolicy {
  const row = object(value, "Policy");
  return { id: guid(row.id, "Policy ID"), displayName: text(row.displayName, "Policy name"),
    type: optionalText(row["@odata.type"]) ?? "unknown", state: "pending" };
}
export async function policyInventory(graph: GraphClient, signal: AbortSignal, policyId?: string): Promise<SyncPolicy[]> {
  if (policyId) return [policyDescriptor(await graph.get(`${costBase}/spendingPolicies/${guid(policyId, "Policy ID")}`, signal))];
  const policies = (await graph.list(`${costBase}/spendingPolicies`, signal)).map(policyDescriptor);
  if (new Set(policies.map((row) => row.id)).size !== policies.length) throw new Error("Policy inventory contained duplicate IDs.");
  return policies;
}
function members(rows: unknown[], groupId: string, users: Map<string, SyncUser>) {
  for (const input of rows) {
    const row = object(input, "Group member");
    const kind = text(row["@odata.type"], "Member object type");
    if (kind !== "#microsoft.graph.user") {
      if (!["#microsoft.graph.group", "#microsoft.graph.device", "#microsoft.graph.servicePrincipal",
        "#microsoft.graph.orgContact"].includes(kind)) throw new Error("Group returned an unrecognized member type.");
      continue;
    }
    const id = guid(row.id, "Member ID");
    const existing = users.get(id);
    if (existing) { if (!existing.groups.includes(groupId)) existing.groups.push(groupId); }
    else users.set(id, { id, displayName: optionalText(row.displayName), username: optionalText(row.userPrincipalName), groups: [groupId] });
  }
}
export function observationsFor(user: SyncUser, policyId: string, rows: Fields[], observedAt = new Date().toISOString()): BalanceObservation[] {
  const base = { userId: user.id, displayName: user.displayName, username: user.username, groups: user.groups, observedAt };
  if (!rows.length) return [{ ...base, serviceId: null, reportedPolicyId: null, match: "no-balance", consumed: null, unit: null,
    remainingKind: "missing", remaining: null, allowed: null, reason: null, sourceTime: null }];
  const services = new Set<string>();
  return rows.map((row) => {
    const serviceId = text(row.serviceId, "Service ID");
    if (services.has(serviceId)) throw new Error("Multiple balance rows for the same user/service are ambiguous.");
    services.add(serviceId);
    const reportedPolicyId = guid(row.spendingPolicyId, "Reported policy ID");
    const consumed = quantity(row.consumedQuantity, "consumed quantity");
    const unit = text(row.unit, "Balance unit");
    let remainingKind: BalanceObservation["remainingKind"] = "missing";
    let remaining: number | null = null;
    if (row.remainingQuantity !== undefined && row.remainingQuantity !== null) {
      const value = object(row.remainingQuantity, "Remaining quantity");
      if (value.unit !== unit) throw new Error("Remaining quantity uses a different unit.");
      if (value["@odata.type"] === "#microsoft.graph.unlimitedCostManagementQuantity") remainingKind = "unlimited";
      else if (value["@odata.type"] === "#microsoft.graph.limitedCostManagementQuantity") {
        remainingKind = "limited"; remaining = quantity(value.value, "remaining quantity");
      } else throw new Error("Unrecognized remaining quantity.");
    }
    let sourceTime: string | null = null;
    if (row.dataStatus !== undefined && row.dataStatus !== null) {
      const status = object(row.dataStatus, "Data status");
      if (status.lastModifiedDateTime !== undefined && status.lastModifiedDateTime !== null) {
        const raw = text(status.lastModifiedDateTime, "Source timestamp");
        if (!/Z$|[+-]\d\d:\d\d$/.test(raw) || !Number.isFinite(Date.parse(raw))) throw new Error("Invalid source timestamp.");
        sourceTime = new Date(raw).toISOString();
      }
    }
    if (row.isConsumptionAllowed !== undefined && row.isConsumptionAllowed !== null && typeof row.isConsumptionAllowed !== "boolean") {
      throw new Error("Invalid consumption-allowed flag.");
    }
    return { ...base, serviceId, reportedPolicyId, match: reportedPolicyId === policyId ? "matched" : "mismatch",
      consumed, unit, remainingKind, remaining, allowed: typeof row.isConsumptionAllowed === "boolean" ? row.isConsumptionAllowed : null,
      reason: optionalText(row.accessDecisionReason), sourceTime };
  });
}
export function aggregateObservations(rows: BalanceObservation[]): PolicySnapshot["totals"] {
  const totals = new Map<string, PolicySnapshot["totals"][number]>();
  const seen = new Set<string>();
  for (const row of rows) {
    const observation = JSON.stringify([row.userId, row.serviceId]);
    if (seen.has(observation)) throw new Error("Duplicate user/service observation cannot be aggregated.");
    seen.add(observation);
    if (row.match !== "matched") continue;
    const serviceId = text(row.serviceId, "Service ID"), unit = text(row.unit, "Unit");
    const key = JSON.stringify([serviceId, unit]);
    const total = totals.get(key) ?? { serviceId, unit, consumed: 0, users: 0 };
    total.consumed += quantity(row.consumed, "consumed quantity");
    if (!Number.isFinite(total.consumed)) throw new Error("Consumption total overflowed.");
    total.users++; totals.set(key, total);
  }
  return [...totals.values()];
}
export async function parallelMap<T, R>(values: T[], concurrency: number, operation: (value: T) => Promise<R>): Promise<R[]> {
  const result = new Array<R>(values.length);
  let next = 0, failure: unknown;
  let failed = false;
  await Promise.all(Array.from({ length: Math.min(concurrency, values.length) }, async () => {
    while (!failed) {
      const index = next++;
      if (index >= values.length) return;
      try { result[index] = await operation(values[index]); }
      catch (error) { failed = true; failure = error; }
    }
  }));
  if (failed) throw failure;
  return result;
}
export async function collectPolicy(graph: GraphClient, descriptor: SyncPolicy, concurrency: number, signal: AbortSignal,
  progress: (phase: string) => void = () => {}): Promise<PolicySnapshot> {
  const startedAt = new Date().toISOString();
  const policy = object(await graph.get(`${costBase}/spendingPolicies/${descriptor.id}`, signal), "Policy");
  if (guid(policy.id, "Policy ID") !== descriptor.id) throw new Error("Graph returned a different policy.");
  const unsupported = policy["@odata.type"] !== selectedPolicyType;
  const knownType = !unsupported || policy["@odata.type"] === "#microsoft.graph.allUsersSpendingPolicy";
  const services = knownType ? await serviceRows(graph, `${costBase}/spendingPolicies/${descriptor.id}/serviceSettings`, signal) : [];
  if (unsupported) return { policy, services, groups: [], observations: [], totals: [], users: 0, exclusions: 1,
    unsupported: true, startedAt, completedAt: new Date().toISOString() };
  progress("Reading assigned groups and transitive users");
  const assigned = await graph.list(`${costBase}/spendingPolicies/${descriptor.id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups`, signal);
  const groups = assigned.map((value) => {
    const row = object(value, "Assigned group");
    return { id: guid(row.id, "Group ID"), displayName: optionalText(row.displayName) };
  });
  if (new Set(groups.map((group) => group.id)).size !== groups.length) throw new Error("Repeated assigned group ID.");
  const users = new Map<string, SyncUser>();
  for (const group of groups) members(await graph.list(`/v1.0/groups/${group.id}/transitiveMembers`, signal), group.id, users);
  progress(`Reading balances for ${users.size} unique users`);
  const batches = await parallelMap([...users.values()], concurrency, async (user) =>
    observationsFor(user, descriptor.id, await serviceRows(graph, `${costBase}/userBalances/${user.id}/serviceBalances`, signal)));
  signal.throwIfAborted();
  const observations = batches.flat();
  return { policy, services, groups, observations, totals: aggregateObservations(observations), users: users.size,
    exclusions: observations.filter((row) => row.match !== "matched").length, unsupported: false,
    startedAt, completedAt: new Date().toISOString() };
}

export async function probeCollectionAccess(graph: GraphClient, policies: SyncPolicy[], signal: AbortSignal): Promise<string[]> {
  const result = ["Policy inventory: verified"];
  const selected = policies.find((policy) => policy.type === selectedPolicyType);
  if (!selected) return [...result, "Assigned-group, membership and balance access: untested (no supported policy)"];
  const groups = await graph.list(`${costBase}/spendingPolicies/${selected.id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups`, signal);
  result.push("Assigned groups: verified");
  if (!groups.length) return [...result, "Membership and balance access: untested (no assigned group)"];
  const groupId = guid(object(groups[0], "Assigned group").id, "Group ID");
  const users = new Map<string, SyncUser>();
  members(await graph.list(`/v1.0/groups/${groupId}/transitiveMembers`, signal), groupId, users);
  result.push("Transitive membership: verified");
  const user = [...users.values()][0];
  if (!user) return [...result, "Balance access: untested (no discovered user)"];
  await serviceRows(graph, `${costBase}/userBalances/${user.id}/serviceBalances`, signal);
  return [...result, "User/service balances: verified"];
}
