export type TierName = "standard" | "enhanced";
export interface TierConfig {
  policyId: string;
  groupId: string;
}
export interface BudgetConfig {
  id: string;
  displayName: string;
  businessArea: string;
  costCenter: string;
  owners: string[];
  plannedUsd?: number;
  approvedCreditCeiling?: number;
  periodLabel?: string;
  tiers: Record<TierName, TierConfig>;
}
export interface DemoConfig {
  version: 1;
  tenantId: string;
  clientId: string;
  apiBaseUrl?: string;
  apiScope?: string;
  serviceId: string;
  budgets: BudgetConfig[];
  requesterIndex?: RequesterMapping[];
  exampleRequests?: ExampleRequest[];
}

export type GrantKind = "permanent" | "billingPeriod";
export interface RequesterMapping {
  username: string;
  budgetId: string;
}
export interface ExampleRequest {
  id: string;
  requesterUsername: string;
  budgetId: string;
  justification: string;
  grantKind: GrantKind;
  requestedAt: string;
}

export function normalizeUsername(value: string): string {
  return value.trim().toLowerCase();
}

export function budgetsForOwner(
  config: DemoConfig,
  account: { username: string; tenantId: string } | null,
): BudgetConfig[] {
  if (!account || account.tenantId.toLowerCase() !== config.tenantId.toLowerCase()) return [];
  const username = normalizeUsername(account.username);
  if (!username) return [];
  return config.budgets.filter((budget) =>
    budget.owners.some((owner) => normalizeUsername(owner) === username));
}

function isObject(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

export function object(value: unknown, label: string): Record<string, unknown> {
  if (!isObject(value)) {
    throw new Error(`${label} must be an object.`);
  }
  return value;
}

export function text(value: unknown, label: string): string {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`${label} must be non-empty text.`);
  }
  return value.trim();
}

export function guid(value: unknown, label: string): string {
  const result = text(value, label).toLowerCase();
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(result) ||
      result === "00000000-0000-0000-0000-000000000000") {
    throw new Error(`${label} must be a non-zero GUID.`);
  }
  return result;
}

export function parseConfig(input: unknown): DemoConfig {
  const data = object(input, "Configuration");
  if (data.version !== 1) throw new Error("Configuration version must be 1.");
  if (!Array.isArray(data.budgets)) throw new Error("budgets must be an array.");
  const budgetIds = new Set<string>();
  const policyIds = new Set<string>();
  const groupIds = new Set<string>();
  const budgets = data.budgets.map((input, index): BudgetConfig => {
    const label = `budgets[${index}]`;
    const budget = object(input, label);
    const id = text(budget.id, `${label}.id`);
    if (budgetIds.has(id)) throw new Error(`Duplicate budget ID: ${id}.`);
    budgetIds.add(id);
    if (!Array.isArray(budget.owners)) throw new Error(`${label}.owners must be an array.`);
    const tiers = object(budget.tiers, `${label}.tiers`);
    if (Object.keys(tiers).length !== 2 || !("standard" in tiers) || !("enhanced" in tiers)) {
      throw new Error(`${label}.tiers must contain exactly standard and enhanced.`);
    }
    const parseTier = (name: TierName): TierConfig => {
      const tier = object(tiers[name], `${label}.${name}`);
      const policyId = guid(tier.policyId, `${label}.${name}.policyId`);
      const groupId = guid(tier.groupId, `${label}.${name}.groupId`);
      if (policyIds.has(policyId)) throw new Error(`Policy ${policyId} is mapped more than once.`);
      if (groupIds.has(groupId)) throw new Error(`Group ${groupId} is mapped more than once.`);
      policyIds.add(policyId);
      groupIds.add(groupId);
      return { policyId, groupId };
    };
    if (budget.plannedUsd !== undefined &&
        (typeof budget.plannedUsd !== "number" || !Number.isFinite(budget.plannedUsd) || budget.plannedUsd < 0)) {
      throw new Error(`${label}.plannedUsd must be a non-negative finite number.`);
    }
    return {
      id,
      displayName: text(budget.displayName, `${label}.displayName`),
      businessArea: text(budget.businessArea, `${label}.businessArea`),
      costCenter: text(budget.costCenter, `${label}.costCenter`),
      owners: budget.owners.map((owner, i) => text(owner, `${label}.owners[${i}]`)),
      plannedUsd: budget.plannedUsd,
      periodLabel: budget.periodLabel === undefined ? undefined : text(budget.periodLabel, `${label}.periodLabel`),
      tiers: { standard: parseTier("standard"), enhanced: parseTier("enhanced") },
    };
  });
  const requesterIndex: RequesterMapping[] = [];
  const mappingKeys = new Set<string>();
  if (data.requesterIndex !== undefined && !Array.isArray(data.requesterIndex)) {
    throw new Error("requesterIndex must be an array.");
  }
  for (const input of data.requesterIndex ?? []) {
    const mapping = object(input, "Requester mapping");
    const username = normalizeUsername(text(mapping.username, "Requester username"));
    if (!/^[^@\s]+@[^@\s]+$/.test(username)) throw new Error("Requester username must be a sign-in UPN.");
    const budgetId = text(mapping.budgetId, "Requester budget ID");
    if (!budgetIds.has(budgetId)) throw new Error(`Requester index references unknown budget ${budgetId}.`);
    const key = JSON.stringify([username, budgetId]);
    if (mappingKeys.has(key)) throw new Error(`Duplicate requester mapping for ${username} and ${budgetId}.`);
    mappingKeys.add(key);
    requesterIndex.push({ username, budgetId });
  }
  const exampleRequests: ExampleRequest[] = [];
  const requestIds = new Set<string>();
  if (data.exampleRequests !== undefined && !Array.isArray(data.exampleRequests)) {
    throw new Error("exampleRequests must be an array.");
  }
  for (const input of data.exampleRequests ?? []) {
    const request = object(input, "Example request");
    const id = text(request.id, "Example request ID");
    if (requestIds.has(id)) throw new Error(`Duplicate example request ID ${id}.`);
    requestIds.add(id);
    const requesterUsername = normalizeUsername(text(request.requesterUsername, "Example requester"));
    const budgetId = text(request.budgetId, "Example request budget");
    if (!mappingKeys.has(JSON.stringify([requesterUsername, budgetId]))) {
      throw new Error(`Example request ${id} needs a matching requester index entry.`);
    }
    if (request.grantKind !== "permanent" && request.grantKind !== "billingPeriod") {
      throw new Error(`Example request ${id} grantKind must be permanent or billingPeriod.`);
    }
    const justification = text(request.justification, "Example justification");
    if (justification.length < 20 || justification.length > 2000) {
      throw new Error(`Example request ${id} justification must be 20-2000 characters.`);
    }
    const requestedAt = text(request.requestedAt, "Example request timestamp");
    if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z$/.test(requestedAt) ||
        !Number.isFinite(Date.parse(requestedAt))) {
      throw new Error(`Example request ${id} needs a UTC requestedAt timestamp.`);
    }
    exampleRequests.push({ id, requesterUsername, budgetId, justification, grantKind: request.grantKind, requestedAt });
  }
  return {
    version: 1,
    tenantId: guid(data.tenantId, "tenantId"),
    clientId: guid(data.clientId, "clientId"),
    apiBaseUrl: data.apiBaseUrl === undefined ? undefined : apiUrl(data.apiBaseUrl),
    apiScope: data.apiScope === undefined ? undefined : text(data.apiScope, "apiScope"),
    serviceId: text(data.serviceId, "serviceId"),
    budgets,
    requesterIndex,
    exampleRequests,
  };
}

function apiUrl(value: unknown): string {
  const url = new URL(text(value, "apiBaseUrl"));
  if (url.protocol !== "https:" && !(url.protocol === "http:" && url.hostname === "localhost")) {
    throw new Error("API URL must use HTTPS (or localhost for development).");
  }
  if (url.username || url.password || url.search || url.hash || url.pathname !== "/") {
    throw new Error("API URL must be an origin without credentials, path or query.");
  }
  return url.origin;
}

export async function loadConfig(): Promise<DemoConfig> {
  const response = await fetch(`${import.meta.env.BASE_URL}demo-config.json`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Cannot load demo-config.json (HTTP ${response.status}).`);
  return parseConfig(await response.json());
}
