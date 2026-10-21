import { randomUUID } from "node:crypto";
import type { z } from "zod";
import type { Budget, Caller, Metadata, PolicyDefaults } from "./contracts.js";
import { ApiError, metadataSchema, seedSchema } from "./model.js";
import { requestId } from "./service.js";
import { checkAllocation, splitCredits, usdToCredits } from "./budget-math.js";

import { resourceKeys, type BuSubmission, type BuDecision, type WorkUpdate, type OnboardingRecord,
  type OnboardingStore, type allocationSchema } from "./onboarding.js";
function conflict(message: string): never { throw new ApiError(409, "onboarding-conflict", message); }
function sameDefaults(left: PolicyDefaults, right: PolicyDefaults): boolean {
  return (["standard", "enhanced"] as const).every((tier) =>
    left[tier].perUserCredits === right[tier].perUserCredits && left[tier].billingMethods === right[tier].billingMethods);
}
export class OnboardingService {
  constructor(private readonly store: OnboardingStore, private readonly admins: string[],
    private readonly clock: () => Date = () => new Date()) {}
  isAdmin(caller: Caller) { return this.admins.includes(caller.objectId); }
  requireAdmin(caller: Caller) {
    if (!this.isAdmin(caller)) throw new ApiError(403, "not-onboarding-admin", "Only a configured onboarding administrator can perform this operation.");
  }
  async setup(caller: Caller) {
    this.requireAdmin(caller);
    return (await this.store.metadata()).metadata;
  }
  async get(caller: Caller, id: string): Promise<OnboardingRecord> {
    const record = await this.store.get(id);
    if (!record || record.tenantId !== caller.tenantId) throw new ApiError(404, "not-found", "Onboarding operation not found.");
    if (record.kind === "allocation") {
      const budget = await this.budget(caller, record.budgetId);
      checkAllocation(record.standard.credits, record.enhanced.credits, budget.approvedCreditCeiling!);
      if (record.resources.standardPolicy !== budget.tiers.standard.policyId ||
          record.resources.enhancedPolicy !== budget.tiers.enhanced.policyId ||
          record.resources.standardGroup !== budget.tiers.standard.groupId ||
          record.resources.enhancedGroup !== budget.tiers.enhanced.groupId) conflict("Budget resource mappings changed. Administrator reconciliation is required.");
    }
    else if (record.requesterId !== caller.objectId) this.requireAdmin(caller);
    return record;
  }
  async list(caller: Caller, inbox: boolean, cursor?: string) {
    if (inbox) this.requireAdmin(caller);
    return this.store.list(inbox ? undefined : caller.objectId, cursor);
  }
  private async saved(record: OnboardingRecord, etag?: string) {
    const success = etag ? await this.store.replace(record, etag) : await this.store.create(record);
    if (!success) conflict("The operation changed. Refresh before continuing.");
    const saved = await this.store.get(record.id);
    if (!saved) throw new Error("Saved onboarding operation could not be read.");
    return saved;
  }
  async submit(caller: Caller, input: BuSubmission): Promise<OnboardingRecord> {
    const id = requestId(caller, input.operationId);
    const existing = await this.store.get(id);
    if (existing) {
      if (existing.requesterId !== caller.objectId || existing.kind !== "onboarding" ||
          existing.name !== input.name || existing.costCenter !== input.costCenter ||
          existing.monthlyBudgetUsd !== input.monthlyBudgetUsd || existing.justification !== input.justification) {
        conflict("The operation ID was already used for another submission.");
      }
      return existing;
    }
    const split = splitCredits(usdToCredits(input.monthlyBudgetUsd));
    const record: OnboardingRecord = {
      ...input, version: 1, kind: "onboarding", id, tenantId: caller.tenantId, requesterId: caller.objectId,
      ...(caller.displayName ? { displayName: caller.displayName } : {}),
      ...(caller.contactHint ? { contactHint: caller.contactHint } : {}),
      budgetId: `bu-${id.slice(0, 24)}`, createdAt: this.clock().toISOString(), status: "pending",
      approvedCredits: 0, standard: { credits: split.standard, perUserCredits: 0, billingMethods: "notSpecified" },
      enhanced: { credits: split.enhanced, perUserCredits: 0, billingMethods: "notSpecified" },
      resources: {}, intents: [], etag: "",
    };
    if (!await this.store.create(record)) return this.submit(caller, input);
    const result = await this.store.get(id);
    if (!result) throw new Error("Submitted onboarding request could not be read. Retry with the same operation ID.");
    return result;
  }
  async decide(caller: Caller, id: string, input: BuDecision) {
    this.requireAdmin(caller);
    const record = await this.get(caller, id);
    if (record.kind !== "onboarding" || record.status !== "pending") conflict("Only a pending onboarding request can be decided.");
    const updated: OnboardingRecord = { ...record, decisionBy: caller.objectId, decisionAt: this.clock().toISOString() };
    if (input.status === "rejected") {
      updated.status = "rejected"; updated.rejectionReason = input.reason;
    } else {
      const defaults = (await this.store.metadata()).metadata?.onboardingDefaults;
      if (!defaults) throw new ApiError(409, "onboarding-defaults-required",
        "Fixed onboarding defaults are not configured. Import the private JSON under Admin > Setup requests before approving.");
      if (!sameDefaults(input, defaults)) {
        conflict("The policy defaults changed or differ from the private configuration. Refresh setup requests before approving.");
      }
      updated.status = "provisioning"; updated.approvedCredits = usdToCredits(record.monthlyBudgetUsd);
      updated.standard = { ...defaults.standard, credits: record.standard.credits };
      updated.enhanced = { ...defaults.enhanced, credits: record.enhanced.credits };
      updated.reviewNote = input.reviewNote;
      updated.financeConfirmed = true;
    }
    return this.saved(updated, input.etag);
  }
  private async budget(caller: Caller, id: string): Promise<Budget> {
    const budget = (await this.store.metadata()).metadata?.budgets.find((item) => item.id === id);
    if (!budget || (!budget.owners.includes(caller.objectId) && !this.isAdmin(caller))) {
      throw new ApiError(404, "not-found", "Budget not found in your current workspace.");
    }
    if (!budget.approvedCreditCeiling) throw new ApiError(409, "no-approved-ceiling", "This budget has no onboarding-approved credit ceiling.");
    return budget;
  }
  async allocation(caller: Caller, budgetId: string, input?: z.infer<typeof allocationSchema>) {
    const budget = await this.budget(caller, budgetId);
    const id = `allocation-${budgetId}`;
    const existing = await this.store.get(id);
    if (!input) return existing;
    try { checkAllocation(input.standard, input.enhanced, budget.approvedCreditCeiling!); }
    catch (error) { throw new ApiError(400, "invalid-allocation", error instanceof Error ? error.message : "Invalid allocation."); }
    if (existing?.operationId === input.operationId) {
      if (existing.standard.credits !== input.standard || existing.enhanced.credits !== input.enhanced) conflict("Allocation operation ID has different content.");
      return existing;
    }
    if (existing && existing.status !== "completed") conflict("Resume the previous allocation before starting another.");
    const record: OnboardingRecord = {
      version: 1, kind: "allocation", id, tenantId: caller.tenantId, requesterId: caller.objectId,
      operationId: input.operationId, name: budget.displayName, costCenter: budget.costCenter,
      monthlyBudgetUsd: String(budget.plannedUsd), justification: "Owner allocation within the approved monthly credit ceiling.",
      budgetId, createdAt: this.clock().toISOString(), status: "provisioning", approvedCredits: budget.approvedCreditCeiling!,
      standard: { credits: input.standard, perUserCredits: 0, billingMethods: "notSpecified" },
      enhanced: { credits: input.enhanced, perUserCredits: 0, billingMethods: "notSpecified" },
      resources: { standardGroup: budget.tiers.standard.groupId, enhancedGroup: budget.tiers.enhanced.groupId,
        standardPolicy: budget.tiers.standard.policyId, enhancedPolicy: budget.tiers.enhanced.policyId },
      intents: [], etag: "",
    };
    return this.saved(record, existing?.etag);
  }
  async update(caller: Caller, id: string, input: WorkUpdate) {
    const record = await this.get(caller, id);
    if (record.kind === "onboarding") this.requireAdmin(caller);
    if (record.status !== "provisioning") conflict("The operation is not awaiting provisioning.");
    if (input.etag !== record.etag) conflict("The operation changed. Refresh before continuing.");
    const now = this.clock();
    if (input.action === "claim") {
      if (record.lease && Date.parse(record.lease.until) > now.getTime()) conflict("Another provisioning session is active. Wait for its lease to expire.");
      return this.saved({ ...record, lastError: undefined, lease: { id: randomUUID(), actor: caller.objectId,
        until: new Date(now.getTime() + 300_000).toISOString() } }, input.etag);
    }
    if (!record.lease || record.lease.id !== input.leaseId || record.lease.actor !== caller.objectId ||
        Date.parse(record.lease.until) <= now.getTime()) conflict("The provisioning lease expired or belongs to another session. Refresh and resume.");
    const updated: OnboardingRecord = { ...record, lease: { ...record.lease, until: new Date(now.getTime() + 300_000).toISOString() } };
    if (input.action === "release") { delete updated.lease; if (input.failure) updated.lastError = input.failure; }
    else if (input.action === "intent") {
      if (!input.resource || record.resources[input.resource]) conflict("A missing resource step is required.");
      updated.intents = [...new Set([...record.intents, input.resource])];
    } else if (input.action === "clear-intent") {
      if (!input.resource || record.resources[input.resource]) conflict("Only an unrecorded resource intent can be cleared.");
      updated.intents = record.intents.filter((key) => key !== input.resource);
    } else if (input.action === "resource") {
      if (!input.resource || !input.resourceId || !record.intents.includes(input.resource)) conflict("Record creation intent before recording a resource.");
      const existing = record.resources[input.resource];
      if (existing && existing !== input.resourceId) conflict("A different resource ID is already recorded.");
      if (Object.entries(record.resources).some(([key, value]) => key !== input.resource && value === input.resourceId)) conflict("Resource IDs must be distinct.");
      updated.resources = { ...record.resources, [input.resource]: input.resourceId };
    } else if (input.action === "complete") {
      if (!input.verified || !resourceKeys.every((key) => record.resources[key])) conflict("Verify all four resources before completing provisioning.");
      checkAllocation(record.standard.credits, record.enhanced.credits, record.approvedCredits);
      if (record.kind === "onboarding") {
        const snapshot = await this.store.metadata();
        const metadata: Metadata = snapshot.metadata ?? { version: 1, tenantId: caller.tenantId, serviceId: "cowork", budgets: [] };
        const budget: Budget = {
          id: record.budgetId, displayName: record.name, businessArea: record.name, costCenter: record.costCenter,
          owners: [record.requesterId], plannedUsd: Number(record.monthlyBudgetUsd), approvedCreditCeiling: record.approvedCredits,
          periodLabel: "Monthly planning budget (USD 0.01 per Copilot Credit)",
          tiers: {
            standard: { groupId: record.resources.standardGroup!, policyId: record.resources.standardPolicy! },
            enhanced: { groupId: record.resources.enhancedGroup!, policyId: record.resources.enhancedPolicy! },
          },
        };
        const existing = metadata.budgets.find((item) => item.id === budget.id);
        if (existing && JSON.stringify(existing) !== JSON.stringify(metadataSchema.parse({ ...metadata, budgets: [budget] }).budgets[0])) {
          conflict("The BU metadata already exists with different content. Administrator reconciliation is required.");
        }
        if (!existing) await this.store.saveMetadata(metadataSchema.parse({ ...metadata, budgets: [...metadata.budgets, budget] }), snapshot.etag);
      }
      updated.status = "completed"; updated.completedAt = now.toISOString(); delete updated.lease; delete updated.lastError;
    }
    return this.saved(updated, input.etag);
  }
  async bootstrap(caller: Caller, input: z.infer<typeof seedSchema>) {
    this.requireAdmin(caller);
    if (input.metadata.tenantId !== caller.tenantId) throw new ApiError(400, "wrong-tenant", "Seed tenant does not match the authenticated tenant.");
    const snapshot = await this.store.metadata();
    const metadata = snapshot.metadata ?? { ...input.metadata, budgets: [] };
    if (metadata.serviceId !== input.metadata.serviceId) conflict("Seed service does not match existing metadata.");
    const defaults = input.metadata.onboardingDefaults;
    if (defaults && metadata.onboardingDefaults && !sameDefaults(defaults, metadata.onboardingDefaults)) {
      conflict("Onboarding defaults already exist with different settings. Nothing was reset.");
    }
    const budgets = [...metadata.budgets];
    for (const budget of input.metadata.budgets) {
      const existing = budgets.find((item) => item.id === budget.id);
      if (existing && JSON.stringify(existing) !== JSON.stringify(budget)) conflict(`Budget ${budget.id} already exists with different settings. Nothing was reset.`);
      if (!existing) budgets.push(budget);
    }
    const merged = metadataSchema.parse({ ...metadata, budgets, ...(defaults ? { onboardingDefaults: defaults } : {}) });
    await this.store.saveMetadata(merged, snapshot.etag);
    for (const mapping of input.requesterIndex) await this.store.seedRouting(mapping.objectId, mapping.budgetId);
    return { budgets: input.metadata.budgets.length, mappings: input.requesterIndex.length };
  }
}
