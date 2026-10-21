import { createHash, randomUUID } from "node:crypto";
import type { Budget, Caller, Context, Decision, Fulfilment, FulfilmentUpdate, Metadata, Page, RequestRecord, Submission } from "./contracts.js";
import { ApiError, periodEnd } from "./model.js";

export interface Store {
  metadata(): Promise<Metadata>;
  routing(caller: Caller): Promise<string | null>;
  get(tenant: string, id: string): Promise<RequestRecord | null>;
  create(record: RequestRecord): Promise<boolean>;
  replace(record: RequestRecord, etag: string): Promise<boolean>;
  list(tenant: string, requesterId?: string, budgets?: string[], cursor?: string): Promise<Page<RequestRecord>>;
}
export function requestId(caller: Caller, operationId: string): string {
  return createHash("sha256").update(`${caller.tenantId}:${caller.objectId}:${operationId}`).digest("hex");
}
function routingError(message: string): never { throw new ApiError(409, "needs-routing", message); }
export class RequestService {
  constructor(private readonly store: Store, private readonly clock: () => Date = () => new Date()) {}

  private async route(caller: Caller, metadata: Metadata): Promise<Budget> {
    const id = await this.store.routing(caller);
    if (!id) routingError("No requester mapping exists. Contact the administrator; the request was not submitted.");
    const budget = metadata.budgets.find((b) => b.id === id);
    if (!budget) routingError("The requester mapping references an unavailable budget. Contact the administrator.");
    if (!budget.owners.length) {
      routingError("There is no current budget owner. Contact the administrator.");
    }
    return budget;
  }
  async context(caller: Caller): Promise<Context> {
    const metadata = await this.store.metadata();
    let route: Context["route"];
    try {
      const budget = await this.route(caller, metadata);
      route = { state: "ready", budgetName: budget.displayName, message: "Routed to the current budget owners." };
    } catch (error) {
      if (!(error instanceof ApiError) || error.code !== "needs-routing") throw error;
      route = { state: "needs-routing", message: error.message };
    }
    return { serviceId: metadata.serviceId, ownedBudgets: metadata.budgets.filter((b) => b.owners.includes(caller.objectId)), route };
  }
  async submit(caller: Caller, input: Submission): Promise<RequestRecord> {
    const id = requestId(caller, input.operationId);
    const replay = (record: RequestRecord): RequestRecord => {
      if (record.justification !== input.justification || record.grantKind !== input.grantKind) {
        throw new ApiError(409, "idempotency-conflict", "This operation ID was already used for different content.");
      }
      return record;
    };
    const existing = await this.store.get(caller.tenantId, id);
    if (existing) return replay(existing);
    const budget = await this.route(caller, await this.store.metadata());
    const now = this.clock();
    const record: RequestRecord = {
      ...input, version: 1, id, tenantId: caller.tenantId, requesterId: caller.objectId, budgetId: budget.id,
      createdAt: now.toISOString(), expiresAt: input.grantKind === "billingPeriod" ? periodEnd(now) : null,
      status: "pending", etag: "",
    };
    if (!await this.store.create(record)) {
      const duplicate = await this.store.get(caller.tenantId, id);
      if (!duplicate) throw new Error("Concurrent submission could not be read. Retry with the same operation ID.");
      return replay(duplicate);
    }
    const saved = await this.store.get(caller.tenantId, id);
    if (!saved) throw new Error("Saved request could not be read. Retry with the same operation ID.");
    return saved;
  }
  history(caller: Caller, cursor?: string) { return this.store.list(caller.tenantId, caller.objectId, undefined, cursor); }
  async inbox(caller: Caller, cursor?: string) {
    const metadata = await this.store.metadata();
    const budgets = metadata.budgets.filter((b) => b.owners.includes(caller.objectId)).map((b) => b.id);
    if (!budgets.length) throw new ApiError(403, "not-owner", "You do not own any budgets.");
    return this.store.list(caller.tenantId, undefined, budgets, cursor);
  }
  async decide(caller: Caller, id: string, input: Decision): Promise<RequestRecord> {
    const record = await this.store.get(caller.tenantId, id);
    const metadata = await this.store.metadata();
    const budget = metadata.budgets.find((b) => b.id === record?.budgetId);
    if (!record || !budget?.owners.includes(caller.objectId)) {
      throw new ApiError(404, "not-found", "Request not found in your current budgets.");
    }
    const currentBudget = await this.route({ tenantId: caller.tenantId, objectId: record.requesterId }, metadata);
    if (currentBudget.id !== record.budgetId) routingError("Requester routing changed. Contact the administrator before deciding.");
    if (record.status !== "pending") throw new ApiError(409, "already-decided", "This request already has a decision.");
    const now = this.clock();
    if (record.expiresAt && Date.parse(record.expiresAt) <= now.getTime()) {
      throw new ApiError(409, "expired", "The temporary request has expired.");
    }
    const updated = {
      ...record, status: input.status, decisionAt: now.toISOString(), decisionBy: caller.objectId,
      ...(input.status === "rejected" ? { rejectionReason: input.reason } : {}),
      ...(input.status === "approved" ? { fulfilment: {
        state: "pending" as const, standard: { ...budget.tiers.standard }, enhanced: { ...budget.tiers.enhanced },
      } } : {}),
    };
    if (!await this.store.replace(updated, input.etag)) throw new ApiError(409, "decision-conflict", "The request changed. Refresh before deciding.");
    const saved = await this.store.get(caller.tenantId, id);
    if (!saved) throw new Error("Decision was stored but could not be read. Refresh the request history.");
    return saved;
  }
  async fulfil(caller: Caller, id: string, input: FulfilmentUpdate): Promise<RequestRecord> {
    const record = await this.store.get(caller.tenantId, id);
    const metadata = await this.store.metadata();
    const budget = metadata.budgets.find((value) => value.id === record?.budgetId);
    if (!record || !budget?.owners.includes(caller.objectId)) {
      throw new ApiError(404, "not-found", "Request not found in your current budgets.");
    }
    if (record.status !== "approved") throw new ApiError(409, "not-approved", "Approve this request before applying the upgrade.");
    const currentBudget = await this.route({ tenantId: caller.tenantId, objectId: record.requesterId }, metadata);
    if (currentBudget.id !== record.budgetId) routingError("Requester routing changed. Reconcile before applying the upgrade.");
    const previous = record.fulfilment;
    if (previous?.state === "applied") return record;
    if (record.etag !== input.etag) throw new ApiError(409, "fulfilment-conflict", "The request changed. Refresh before retrying.");
    const fulfilment: Fulfilment = previous ?? {
      state: "pending", standard: { ...budget.tiers.standard }, enhanced: { ...budget.tiers.enhanced },
    };
    if (fulfilment.standard.groupId === fulfilment.enhanced.groupId ||
        (["standard", "enhanced"] as const).some((tier) =>
          fulfilment[tier].groupId !== budget.tiers[tier].groupId || fulfilment[tier].policyId !== budget.tiers[tier].policyId)) {
      throw new ApiError(409, "fulfilment-mapping-changed", "Policy/group mappings changed. Administrator reconciliation is required.");
    }
    const now = this.clock();
    if ((input.action === "claim" || input.action === "heartbeat") && record.expiresAt && Date.parse(record.expiresAt) <= now.getTime()) {
      throw new ApiError(409, "expired", "The request period expired. No further membership changes are authorized; review any partial upgrade manually.");
    }
    if (input.action === "claim") {
      if (fulfilment.lease && Date.parse(fulfilment.lease.until) > now.getTime()) {
        throw new ApiError(409, "fulfilment-busy", `An upgrade attempt is active until ${fulfilment.lease.until}. Refresh before retrying.`);
      }
      fulfilment.state = "applying";
      fulfilment.lease = { id: randomUUID(), actor: caller.objectId, until: new Date(now.getTime() + 300_000).toISOString() };
    } else {
      if (!fulfilment.lease || fulfilment.lease.id !== input.leaseId || fulfilment.lease.actor !== caller.objectId ||
          (input.action !== "release" && Date.parse(fulfilment.lease.until) <= now.getTime())) {
        throw new ApiError(409, "fulfilment-lease", "This upgrade attempt no longer holds the lease. Refresh before retrying.");
      }
      if (input.action === "heartbeat") fulfilment.lease.until = new Date(now.getTime() + 300_000).toISOString();
      else {
        if (input.action === "complete" && input.verified !== true) {
          throw new ApiError(409, "verification-required", "Verify the final group memberships before completing.");
        }
        fulfilment.state = input.action === "complete" ? "applied" : "needs-attention";
        delete fulfilment.lease;
        if (input.action === "complete") {
          fulfilment.appliedAt = now.toISOString();
          fulfilment.appliedBy = caller.objectId;
        }
      }
    }
    if (!await this.store.replace({ ...record, fulfilment }, input.etag)) {
      throw new ApiError(409, "fulfilment-conflict", "The request changed. Refresh before retrying.");
    }
    const saved = await this.store.get(caller.tenantId, id);
    if (!saved) throw new Error("Upgrade progress could not be read. Refresh before retrying.");
    return saved;
  }
}
