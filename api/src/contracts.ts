export interface Budget {
  id: string;
  displayName: string;
  businessArea: string;
  costCenter: string;
  owners: string[];
  plannedUsd?: number;
  periodLabel?: string;
  tiers: Record<"standard" | "enhanced", { policyId: string; groupId: string }>;
}
export interface Caller { tenantId: string; objectId: string }
export interface Metadata { version: 1; tenantId: string; serviceId: string; budgets: Budget[] }
export type GrantKind = "permanent" | "billingPeriod";
export interface Submission { operationId: string; justification: string; grantKind: GrantKind }
export interface RequestRecord extends Submission {
  version: 1;
  id: string;
  tenantId: string;
  requesterId: string;
  budgetId: string;
  createdAt: string;
  expiresAt: string | null;
  status: "pending" | "approved" | "rejected";
  decisionBy?: string;
  decisionAt?: string;
  rejectionReason?: string;
  etag: string;
}
export interface Route {
  state: "ready" | "needs-routing";
  budgetName?: string;
  message: string;
}
export interface Context {
  serviceId: string;
  ownedBudgets: Budget[];
  route: Route;
}
export interface Page<T> { items: T[]; cursor?: string }
export interface Decision { status: "approved" | "rejected"; reason?: string; etag: string }
