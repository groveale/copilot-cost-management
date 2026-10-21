import { guid, object, text, type BudgetConfig, type TierName } from "./config";
import { GraphClient } from "./graph";

export const tierNames: TierName[] = ["standard", "enhanced"];
export type Fields = Record<string, unknown>;
export type Resource<T> =
  | { state: "loading" }
  | { state: "ready"; data: T; fetchedAt: string }
  | { state: "error"; error: string };
export interface Entity {
  id: string;
  fields: Fields;
}
export interface Member extends Entity {
  kind: string;
}
export interface TierDetails {
  policy: Resource<Entity>;
  services: Resource<Fields[]>;
  assignedGroups: Resource<Entity[]>;
}
export interface GroupDetails {
  directory: Resource<Entity>;
  members: Resource<Member[]>;
}
export interface BudgetSnapshot {
  budgetId: string;
  requestedAt: string;
  tiers: Record<TierName, TierDetails>;
  groups: Record<string, GroupDetails>;
  balances: Record<string, Resource<Fields | null>>;
}

const base = "/beta/copilot/costManagement";
const loading = (): Resource<never> => ({ state: "loading" });
const emptyGroup = (): GroupDetails => ({ directory: loading(), members: loading() });
const emptyTier = (): TierDetails => ({ policy: loading(), services: loading(), assignedGroups: loading() });

export function initialSnapshot(budget: BudgetConfig): BudgetSnapshot {
  return {
    budgetId: budget.id,
    requestedAt: new Date().toISOString(),
    tiers: { standard: emptyTier(), enhanced: emptyTier() },
    groups: Object.fromEntries(tierNames.map((name) => [budget.tiers[name].groupId, emptyGroup()])),
    balances: {},
  };
}

function entity(value: unknown, label: string, expectedId?: string): Entity {
  const fields = object(value, label);
  const id = guid(fields.id, `${label} ID`);
  if (expectedId && id !== expectedId) throw new Error(`${label} returned a different ID than requested.`);
  return { id, fields };
}

export async function serviceRows(graph: GraphClient, path: string, signal: AbortSignal): Promise<Fields[]> {
  const response = object(await graph.get(path, signal), "Service response");
  let rows: unknown[];
  if (Array.isArray(response.value)) {
    rows = response.value;
    const next = response["@odata.nextLink"];
    if (next !== undefined && next !== null) {
      rows = [...rows, ...await graph.list(text(next, "Service continuation link"), signal)];
    }
  } else if (typeof response.serviceId === "string") {
    rows = [response];
  } else {
    throw new Error("Unrecognized service response: expected a value array or an explicit service row.");
  }
  return rows.map((value) => {
    const row = object(value, "Service row");
    text(row.serviceId, "Service ID");
    return row;
  });
}

export async function loadUserBalance(graph: GraphClient, userId: string, serviceId: string, signal: AbortSignal): Promise<Fields | null> {
  const id = guid(userId, "User object ID");
  const rows = await serviceRows(graph, `${base}/userBalances/${id}/serviceBalances`, signal);
  const matches = rows.filter((row) => row.serviceId === serviceId);
  if (matches.length > 1) throw new Error(`Graph returned multiple ${serviceId} rows for the same user.`);
  return matches[0] ?? null;
}

export async function loadSpendingPolicy(graph: GraphClient, policyId: unknown, signal: AbortSignal): Promise<Entity> {
  const id = guid(policyId, "Spending policy ID");
  return entity(await graph.get(`${base}/spendingPolicies/${id}`, signal), "Spending policy", id);
}

export async function loadBudgetDetails(
  budget: BudgetConfig,
  serviceId: string,
  graph: GraphClient,
  signal: AbortSignal,
  report: (snapshot: BudgetSnapshot) => void,
): Promise<void> {
  let snapshot = initialSnapshot(budget);
  report(snapshot);
  const publish = () => {
    signal.throwIfAborted();
    report(snapshot);
  };
  const read = async <T>(operation: () => Promise<T>): Promise<Resource<T>> => {
    try {
      const data = await operation();
      signal.throwIfAborted();
      return { state: "ready", data, fetchedAt: new Date().toISOString() };
    } catch (error) {
      signal.throwIfAborted();
      return { state: "error", error: error instanceof Error ? error.message : String(error) };
    }
  };
  const updateTier = (name: TierName, patch: Partial<TierDetails>) => {
    snapshot = { ...snapshot, tiers: { ...snapshot.tiers, [name]: { ...snapshot.tiers[name], ...patch } } };
    publish();
  };
  for (const name of tierNames) {
    const tier = budget.tiers[name];
    updateTier(name, {
      policy: await read(() => loadSpendingPolicy(graph, tier.policyId, signal)),
    });
    updateTier(name, {
      services: await read(() => serviceRows(graph, `${base}/spendingPolicies/${tier.policyId}/serviceSettings`, signal)),
    });
    const assignedGroups = await read(async () => {
      const rows = await graph.list(`${base}/spendingPolicies/${tier.policyId}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups`, signal);
      return rows.map((row) => entity(row, "Assigned group"));
    });
    if (assignedGroups.state === "ready") {
      for (const group of assignedGroups.data) {
        if (!snapshot.groups[group.id]) {
          snapshot = { ...snapshot, groups: { ...snapshot.groups, [group.id]: emptyGroup() } };
        }
      }
    }
    updateTier(name, { assignedGroups });
  }
  const users = new Set<string>();
  for (const id of Object.keys(snapshot.groups)) {
    const directory = await read(async () => entity(
      await graph.get(`/v1.0/groups/${id}`, signal), "Directory group", id));
    snapshot = { ...snapshot, groups: { ...snapshot.groups, [id]: { ...snapshot.groups[id], directory } } };
    publish();
    const members = await read(async () => {
      const rows = await graph.list(`/v1.0/groups/${id}/members`, signal);
      const unique = new Map<string, Member>();
      for (const row of rows) {
        const member = entity(row, "Group member");
        const type = member.fields["@odata.type"];
        const kind = typeof type === "string" ? type.replace("#microsoft.graph.", "") : "unknown";
        unique.set(member.id, { ...member, kind });
      }
      return [...unique.values()];
    });
    if (members.state === "ready") {
      for (const member of members.data) if (member.kind === "user") users.add(member.id);
    }
    snapshot = { ...snapshot, groups: { ...snapshot.groups, [id]: { ...snapshot.groups[id], members } } };
    publish();
  }
  snapshot = { ...snapshot, balances: Object.fromEntries([...users].map((id) => [id, loading()])) };
  publish();
  for (const id of users) {
    const balance = await read(() => loadUserBalance(graph, id, serviceId, signal));
    snapshot = { ...snapshot, balances: { ...snapshot.balances, [id]: balance } };
    publish();
  }
}

export function stopLoading(snapshot: BudgetSnapshot, reason: string): BudgetSnapshot {
  const stop = <T>(resource: Resource<T>): Resource<T> =>
    resource.state === "loading" ? { state: "error", error: reason } : resource;
  return {
    ...snapshot,
    tiers: {
      standard: {
        policy: stop(snapshot.tiers.standard.policy),
        services: stop(snapshot.tiers.standard.services),
        assignedGroups: stop(snapshot.tiers.standard.assignedGroups),
      },
      enhanced: {
        policy: stop(snapshot.tiers.enhanced.policy),
        services: stop(snapshot.tiers.enhanced.services),
        assignedGroups: stop(snapshot.tiers.enhanced.assignedGroups),
      },
    },
    groups: Object.fromEntries(Object.entries(snapshot.groups).map(([id, group]) =>
      [id, { directory: stop(group.directory), members: stop(group.members) }])),
    balances: Object.fromEntries(Object.entries(snapshot.balances).map(([id, balance]) => [id, stop(balance)])),
  };
}

export function formatField(value: unknown): string {
  if (value === undefined || value === null || value === "") return "Not reported";
  if (typeof value === "boolean") return value ? "Yes" : "No";
  if (typeof value === "string" || typeof value === "number") return String(value);
  return JSON.stringify(value);
}

function unitLabel(unit: unknown): string {
  return unit === "copilotCredit" ? "Copilot credits" :
    typeof unit === "string" && unit ? unit : "(unit not reported)";
}

export function formatQuantity(value: unknown): string {
  if (value === undefined || value === null) return "Not reported";
  if (typeof value !== "object" || Array.isArray(value)) return "Invalid quantity returned";
  const quantity = object(value, "Quantity");
  if (quantity["@odata.type"] === "#microsoft.graph.unlimitedCostManagementQuantity") {
    return `Unlimited ${unitLabel(quantity.unit)}`;
  }
  if (quantity["@odata.type"] !== "#microsoft.graph.limitedCostManagementQuantity") {
    return `Unsupported quantity type: ${formatField(quantity["@odata.type"])}`;
  }
  if (typeof quantity.value !== "number" || !Number.isFinite(quantity.value) || quantity.value < 0) {
    return "Invalid finite quantity returned";
  }
  return `${quantity.value.toLocaleString()} ${unitLabel(quantity.unit)}`;
}

export function formatConsumed(row: Fields): string {
  if (row.consumedQuantity === undefined || row.consumedQuantity === null) return "Not reported";
  if (typeof row.consumedQuantity !== "number" || !Number.isFinite(row.consumedQuantity) || row.consumedQuantity < 0) {
    return "Invalid consumed quantity returned";
  }
  return `${row.consumedQuantity.toLocaleString()} ${unitLabel(row.unit)}`;
}

export function formatTime(value: unknown): string {
  if (value === undefined || value === null) return "Not reported";
  if (typeof value !== "string" || !Number.isFinite(Date.parse(value))) return `Invalid timestamp: ${formatField(value)}`;
  return new Date(value).toLocaleString();
}

export function sourceTimestamp(row: Fields): unknown {
  const status = row.dataStatus;
  if (status === undefined || status === null) return undefined;
  return object(status, "Balance dataStatus").lastModifiedDateTime;
}

export function tierGroupIds(budget: BudgetConfig, name: TierName, snapshot: BudgetSnapshot): string[] {
  const assigned = snapshot.tiers[name].assignedGroups;
  return [...new Set([
    budget.tiers[name].groupId,
    ...(assigned.state === "ready" ? assigned.data.map((group) => group.id) : []),
  ])];
}

export function memberPolicyNote(budget: BudgetConfig, snapshot: BudgetSnapshot, memberId: string, row: Fields): string | null {
  const names = tierNames.filter((name) => tierGroupIds(budget, name, snapshot).some((id) => {
    const members = snapshot.groups[id]?.members;
    return members?.state === "ready" && members.data.some((member) => member.id === memberId);
  }));
  if (names.length > 1) return "Member appears in both tiers. This reading is not a tier spend total.";
  const reported = row.spendingPolicyId;
  if (typeof reported !== "string" || !reported.trim()) return "Active policy is not reported; tier attribution is unavailable.";
  if (names.length === 1 && reported.toLowerCase() !== budget.tiers[names[0]].policyId) {
    return "Reported policy differs from this membership tier. Consumption is shown as a user reading, not attributed to this tier.";
  }
  if (row.unit !== "copilotCredit") return "Credit unit is missing or unexpected; this reading is not attributed to the tier.";
  return null;
}
