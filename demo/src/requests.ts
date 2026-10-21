import { normalizeUsername, type BudgetConfig, type DemoConfig, type GrantKind } from "./config";

export interface RequestIdentity {
  username: string;
  tenantId: string;
}
export type RequestRoute =
  | { state: "ready"; budget: BudgetConfig; approvers: string[] }
  | { state: "needs-routing"; reason: string; budget?: BudgetConfig };

export function resolveRequestRoute(config: DemoConfig, account: RequestIdentity): RequestRoute {
  if (account.tenantId.toLowerCase() !== config.tenantId.toLowerCase()) {
    return { state: "needs-routing", reason: "Sign in to the configured organization before requesting an increase." };
  }
  const username = normalizeUsername(account.username);
  const ids = [...new Set((config.requesterIndex ?? [])
    .filter((entry) => entry.username === username).map((entry) => entry.budgetId))];
  if (ids.length === 0) return { state: "needs-routing", reason: "No budget mapping was found for your signed-in account. An administrator must add it to the requester index." };
  if (ids.length !== 1) return { state: "needs-routing", reason: "More than one budget matches this account. An administrator must resolve the routing; no owner has been guessed." };
  const budget = config.budgets.find((item) => item.id === ids[0]);
  if (!budget) return { state: "needs-routing", reason: "The indexed budget is unavailable. An administrator must repair the mapping." };
  const owners = [...new Set(budget.owners.map(normalizeUsername))];
  if (owners.length === 0 || owners.some((owner) => !/^[^@\s]+@[^@\s]+$/.test(owner))) {
    return { state: "needs-routing", budget, reason: "This budget needs valid owner sign-in usernames before requests can be routed." };
  }
  return { state: "ready", budget, approvers: owners };
}

export function periodEndUtc(requestedAt: Date): string {
  if (!Number.isFinite(requestedAt.getTime())) throw new Error("A valid request timestamp is required.");
  return new Date(Date.UTC(requestedAt.getUTCFullYear(), requestedAt.getUTCMonth() + 1, 1)).toISOString();
}

export function requestExpiry(kind: GrantKind, requestedAt: Date): string | null {
  return kind === "permanent" ? null : periodEndUtc(requestedAt);
}

export function expiryLabel(expiresAt: string | null): string {
  return expiresAt === null ? "Permanent - no scheduled expiry" :
    `Until ${new Date(expiresAt).toLocaleString("en-GB", { timeZone: "UTC", timeZoneName: "short" })} (exclusive)`;
}

export interface RequestPreview {
  requester: string;
  justification: string;
  grantKind: GrantKind;
  requestedAt: string;
  expiresAt: string | null;
  route: RequestRoute;
}

export function createRequestPreview(
  config: DemoConfig,
  account: RequestIdentity,
  justification: string,
  grantKind: GrantKind,
  now = new Date(),
): RequestPreview {
  const reason = justification.trim();
  if (reason.length < 20 || reason.length > 2000) throw new Error("Business justification must be 20-2000 characters.");
  if (grantKind !== "permanent" && grantKind !== "billingPeriod") throw new Error("Choose a permanent or current-billing-period increase.");
  if (!Number.isFinite(now.getTime())) throw new Error("A valid request timestamp is required.");
  return {
    requester: account.username,
    justification: reason,
    grantKind,
    requestedAt: now.toISOString(),
    expiresAt: requestExpiry(grantKind, now),
    route: resolveRequestRoute(config, account),
  };
}
