export interface Budget {
  id: string;
  displayName: string;
  businessArea: string;
  costCenter: string;
  owners: string[];
  plannedUsd?: number;
  approvedCreditCeiling?: number;
  periodLabel?: string;
  tiers: Record<"standard" | "enhanced", { policyId: string; groupId: string }>;
}
export interface Caller { tenantId: string; objectId: string; displayName?: string; contactHint?: string }
export interface PolicyTierDefaults {
  perUserCredits: number | null;
  billingMethods: "notSpecified" | "prepaidCredits" | "payAsYouGo" | "prepaidCredits,payAsYouGo";
}
export interface PolicyDefaults { standard: PolicyTierDefaults; enhanced: PolicyTierDefaults }
export interface Metadata { version: 1; tenantId: string; serviceId: string; budgets: Budget[]; onboardingDefaults?: PolicyDefaults }
export type GrantKind = "permanent" | "billingPeriod";
export interface Submission { operationId: string; justification: string; grantKind: GrantKind }
export interface Fulfilment {
  state: "pending" | "applying" | "needs-attention" | "applied";
  standard: Budget["tiers"]["standard"];
  enhanced: Budget["tiers"]["enhanced"];
  lease?: { id: string; actor: string; until: string };
  appliedAt?: string;
  appliedBy?: string;
}
export interface FulfilmentUpdate {
  action: "claim" | "heartbeat" | "release" | "complete";
  etag: string;
  leaseId?: string;
  verified?: true;
}
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
  fulfilment?: Fulfilment;
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
