import { beforeEach, describe, expect, it } from "vitest";
import type { Caller, Metadata } from "../src/contracts.js";
import { ApiError, metadataSchema } from "../src/model.js";
import { checkAllocation, splitCredits, usdToCredits } from "../src/budget-math.js";
import { buDecisionSchema, buSubmissionSchema, resourceKeys, type BuDecision, type BuSubmission, type OnboardingRecord, type OnboardingStore } from "../src/onboarding.js";
import { OnboardingService } from "../src/onboarding-service.js";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const owner: Caller = { tenantId: id(1), objectId: id(2), displayName: "Example owner", contactHint: "owner@example.test" };
const admin: Caller = { tenantId: id(1), objectId: id(3) };
const other: Caller = { tenantId: id(1), objectId: id(4) };
const submission: BuSubmission = { operationId: id(5), name: "Finance", costCenter: "CC200",
  monthlyBudgetUsd: "20000", justification: "We need a governed business unit for this team." };
const legal = metadataSchema.parse({
  version: 1, tenantId: id(1), serviceId: "cowork", budgets: [{
    id: "legal", displayName: "Legal", businessArea: "Legal", costCenter: "CC100", owners: [admin.objectId],
    tiers: { standard: { policyId: id(11), groupId: id(12) }, enhanced: { policyId: id(13), groupId: id(14) } },
  }],
});
class MemoryStore implements OnboardingStore {
  records = new Map<string, OnboardingRecord>();
  routes = new Map<string, string>();
  value: Metadata | null = null;
  revision = 0;
  metadataRevision = 0;
  async get(id: string) { return structuredClone(this.records.get(id) ?? null); }
  async create(record: OnboardingRecord) {
    if (this.records.has(record.id)) return false;
    this.records.set(record.id, { ...structuredClone(record), etag: String(++this.revision) }); return true;
  }
  async replace(record: OnboardingRecord, etag: string) {
    if (this.records.get(record.id)?.etag !== etag) return false;
    this.records.set(record.id, { ...structuredClone(record), etag: String(++this.revision) }); return true;
  }
  async list(requesterId?: string) { return { items: [...this.records.values()]
    .filter((record) => record.kind === "onboarding" && (!requesterId || record.requesterId === requesterId)).map((record) => structuredClone(record)) }; }
  async metadata() { return { metadata: structuredClone(this.value), etag: this.value ? String(this.metadataRevision) : null }; }
  async saveMetadata(value: Metadata, etag: string | null) {
    if (etag !== (this.value ? String(this.metadataRevision) : null)) throw new ApiError(409, "metadata-conflict", "Metadata changed.");
    this.value = structuredClone(value); this.metadataRevision++;
  }
  async seedRouting(objectId: string, budgetId: string) {
    if (this.routes.has(objectId) && this.routes.get(objectId) !== budgetId) throw new ApiError(409, "routing-conflict", "Existing routing is different.");
    this.routes.set(objectId, budgetId);
  }
}
let store: MemoryStore;
let service: OnboardingService;
let now: Date;
beforeEach(() => {
  store = new MemoryStore(); now = new Date("2026-10-21T09:00:00Z");
  service = new OnboardingService(store, [admin.objectId], () => now);
});
const approval = (etag: string): BuDecision => ({ status: "approved", etag, financeConfirmed: true,
  reviewNote: "Cost centre and cross charging validated.",
  standard: { perUserCredits: 5000, billingMethods: "prepaidCredits" },
  enhanced: { perUserCredits: 50000, billingMethods: "prepaidCredits" } });
async function approved() {
  const pending = await service.submit(owner, submission);
  return service.decide(admin, pending.id, approval(pending.etag));
}
async function completed() {
  let record = await approved();
  record = await service.update(admin, record.id, { action: "claim", etag: record.etag });
  const leaseId = record.lease!.id;
  for (const [index, key] of resourceKeys.entries()) {
    record = await service.update(admin, record.id, { action: "intent", resource: key, etag: record.etag, leaseId });
    record = await service.update(admin, record.id, { action: "resource", resource: key, resourceId: id(20 + index), etag: record.etag, leaseId });
  }
  return service.update(admin, record.id, { action: "complete", verified: true, etag: record.etag, leaseId });
}
describe("USD planning conversion and approved ceiling", () => {
  it("converts the exact required $20K example into two one-million-credit policies", () => {
    expect(usdToCredits("20000")).toBe(2_000_000);
    expect(splitCredits(usdToCredits("20000.00"))).toEqual({ standard: 1_000_000, enhanced: 1_000_000 });
    expect(splitCredits(usdToCredits("0.03"))).toEqual({ standard: 1, enhanced: 2 });
    expect(splitCredits(usdToCredits("0.01"))).toEqual({ standard: 0, enhanced: 1 });
    expect(usdToCredits("1.23")).toBe(123);
    expect(() => checkAllocation(1_500_000, 500_000, 2_000_000)).not.toThrow();
    expect(() => checkAllocation(1_500_000, 500_001, 2_000_000)).toThrow();
  });
  it.each(["0", "-2", "1.001", "2e4", "NaN", "Infinity", "999999999999999999"])("rejects an invalid planning amount %s", (value) => {
    expect(() => usdToCredits(value)).toThrow();
    expect(buSubmissionSchema.safeParse({ ...submission, monthlyBudgetUsd: value }).success).toBe(false);
  });
  it("rejects fractional, negative and unsafe credit allocations", () => {
    for (const value of [-1, 0.1, Infinity, Number.MAX_SAFE_INTEGER + 1]) expect(() => checkAllocation(value, 1, 2_000_000)).toThrow();
  });
});
describe("onboarding authorization and persistence", () => {
  it("accepts ordinary-user submissions before metadata exists without granting ownership", async () => {
    const record = await service.submit(owner, submission);
    expect(record).toMatchObject({ requesterId: owner.objectId, displayName: owner.displayName,
      contactHint: owner.contactHint, status: "pending", approvedCredits: 0 });
    expect(store.value).toBeNull();
    expect(buSubmissionSchema.safeParse({ ...submission, requesterId: admin.objectId }).success).toBe(false);
    expect(buSubmissionSchema.safeParse({ ...submission, approvedCredits: 5_000_000 }).success).toBe(false);
  });
  it("binds idempotency to the requester and complete payload, including concurrent submissions", async () => {
    const [first, second] = await Promise.all([service.submit(owner, submission), service.submit(owner, submission)]);
    expect(first.id).toBe(second.id); expect(store.records.size).toBe(1);
    await expect(service.submit(owner, { ...submission, costCenter: "other" })).rejects.toMatchObject({ status: 409 });
    expect((await service.submit(other, submission)).id).not.toBe(first.id);
  });
  it("restricts inbox, setup, import and decisions to configured administrators, not budget owners", async () => {
    store.value = legal;
    const record = await service.submit(owner, submission);
    expect(service.isAdmin(owner)).toBe(false);
    await expect(service.list(owner, true)).rejects.toMatchObject({ status: 403 });
    await expect(service.setup(owner)).rejects.toMatchObject({ status: 403 });
    await expect(service.bootstrap(owner, { metadata: legal, requesterIndex: [] })).rejects.toMatchObject({ status: 403 });
    await expect(service.decide(owner, record.id, approval(record.etag))).rejects.toMatchObject({ status: 403 });
    await expect(service.get(other, record.id)).rejects.toMatchObject({ status: 403 });
    await expect(service.get({ ...owner, tenantId: id(99) }, record.id)).rejects.toMatchObject({ status: 404 });
    expect((await service.list(other, false)).items).toEqual([]);
    expect((await service.list(admin, true)).items).toHaveLength(1);
  });
  it("captures the approved ceiling and finance review without publishing a BU early", async () => {
    const record = await approved();
    expect(record).toMatchObject({ status: "provisioning", approvedCredits: 2_000_000,
      financeConfirmed: true, decisionBy: admin.objectId, standard: { credits: 1_000_000 }, enhanced: { credits: 1_000_000 } });
    expect(store.value).toBeNull();
    await expect(service.decide(admin, record.id, approval(record.etag))).rejects.toMatchObject({ status: 409 });
    expect(buDecisionSchema.safeParse({ ...approval(record.etag), financeConfirmed: false }).success).toBe(false);
    expect(buDecisionSchema.safeParse(approval("*")).success).toBe(false);
  });
  it("protects competing decisions with ETags and records rejections", async () => {
    const record = await service.submit(owner, submission);
    await expect(service.decide(admin, record.id, approval("stale"))).rejects.toMatchObject({ status: 409 });
    const rejected = await service.decide(admin, record.id, { status: "rejected", etag: record.etag, reason: "Cost centre was not validated." });
    expect(rejected.status).toBe("rejected"); expect(store.value).toBeNull();
  });
  it("uses exclusive expiring leases and requires recorded intent before accepting resource IDs", async () => {
    let record = await approved();
    await expect(service.update(owner, record.id, { action: "claim", etag: record.etag })).rejects.toMatchObject({ status: 403 });
    record = await service.update(admin, record.id, { action: "claim", etag: record.etag });
    await expect(service.update(admin, record.id, { action: "claim", etag: record.etag })).rejects.toMatchObject({ status: 409 });
    await expect(service.update(admin, record.id, { action: "resource", etag: record.etag, leaseId: record.lease!.id,
      resource: "standardGroup", resourceId: id(20) })).rejects.toMatchObject({ status: 409 });
    await expect(service.update(admin, record.id, { action: "complete", verified: true, etag: record.etag, leaseId: record.lease!.id })).rejects.toMatchObject({ status: 409 });
    now = new Date(now.getTime() + 300_001);
    await expect(service.update(admin, record.id, { action: "heartbeat", etag: record.etag, leaseId: record.lease!.id })).rejects.toMatchObject({ status: 409 });
    const resumed = await service.update(admin, record.id, { action: "claim", etag: record.etag });
    expect(resumed.lease!.id).not.toBe(record.lease!.id);
  });
  it("publishes the completed BU without modifying Legal or existing requester routing", async () => {
    store.value = structuredClone(legal); store.routes.set(owner.objectId, "legal");
    const record = await completed();
    expect(record.status).toBe("completed"); expect(record.lease).toBeUndefined();
    expect(store.value!.budgets[0]).toEqual(legal.budgets[0]);
    expect(store.value!.budgets[1]).toMatchObject({ id: record.budgetId, owners: [owner.objectId],
      approvedCreditCeiling: 2_000_000, plannedUsd: 20000, tiers: { standard: { groupId: id(20), policyId: id(22) } } });
    expect(store.routes.get(owner.objectId)).toBe("legal");
  });
  it("authorizes allocations by current ownership and validates the ceiling server-side", async () => {
    const record = await completed();
    await expect(service.allocation(other, record.budgetId, { operationId: id(6), standard: 1, enhanced: 1 })).rejects.toMatchObject({ status: 404 });
    await expect(service.allocation(owner, record.budgetId, { operationId: id(6), standard: 1_500_000, enhanced: 500_001 })).rejects.toMatchObject({ status: 400 });
    const work = await service.allocation(owner, record.budgetId, { operationId: id(6), standard: 1_500_000, enhanced: 500_000 });
    expect(work!.approvedCredits).toBe(2_000_000);
    await expect(service.allocation(owner, record.budgetId, { operationId: id(7), standard: 1_000_000, enhanced: 1_000_000 })).rejects.toMatchObject({ status: 409 });
    store.value!.budgets[0].owners = [other.objectId];
    await expect(service.update(owner, work!.id, { action: "claim", etag: work!.etag })).rejects.toMatchObject({ status: 404 });
  });
  it("allows additive, repeatable seed imports but refuses routing conflicts and different tenant seeds", async () => {
    const seed = { metadata: legal, requesterIndex: [{ objectId: owner.objectId, budgetId: "legal" }] };
    await service.bootstrap(admin, seed); await service.bootstrap(admin, seed);
    expect(store.value!.budgets).toHaveLength(1); expect(store.routes.get(owner.objectId)).toBe("legal");
    await expect(service.bootstrap(admin, { ...seed, metadata: { ...legal, tenantId: id(99) } })).rejects.toMatchObject({ status: 400 });
    await expect(service.bootstrap(admin, { ...seed, metadata: { ...legal, budgets: [{ ...legal.budgets[0], costCenter: "different" }] } })).rejects.toMatchObject({ status: 409 });
    store.routes.set(owner.objectId, "already-mapped");
    await expect(service.bootstrap(admin, seed)).rejects.toMatchObject({ status: 409 });
    expect(store.routes.get(owner.objectId)).toBe("already-mapped");
  });
});
