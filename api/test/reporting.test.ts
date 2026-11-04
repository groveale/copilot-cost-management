import { beforeEach, describe, expect, it, vi } from "vitest";
import type { Caller, Metadata } from "../src/contracts.js";
import { ReportingService, policyOwners, reportingConfiguration, reportingSecuritySchema, type ReportingConfiguration,
  type ReportingSnapshot, type ReportingStore, type SecurityState } from "../src/reporting.js";
import { syncSettings } from "../src/sync.js";
import type { SyncJob } from "../src/sync-contracts.js";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const admin: Caller = { tenantId: id(1), objectId: id(2) };
const owner: Caller = { tenantId: id(1), objectId: id(3) };
const metadata: Metadata = { version: 1, tenantId: id(1), serviceId: "cowork", budgets: [{
  id: "legal", displayName: "Legal", businessArea: "Legal", costCenter: "A", owners: [owner.objectId],
  tiers: { standard: { policyId: id(11), groupId: id(12) }, enhanced: { policyId: id(13), groupId: id(14) } },
}] };
class MemoryStore implements ReportingStore {
  snapshot: ReportingSnapshot = { metadata: structuredClone(metadata), etag: "metadata-v1" };
  value: SecurityState | null = null;
  assert = vi.fn();
  async metadata() { return structuredClone(this.snapshot); }
  async state() { return structuredClone(this.value); }
  async save(value: SecurityState) { this.value = structuredClone(value); }
  async locked<T>(operation: (assert: () => void, leaseId?: string) => Promise<T>) { return operation(this.assert, "metadata-lease"); }
}
const config: ReportingConfiguration = { enabled: true, accessMode: "sharepoint-acl", snapshotDestination: null,
  tenantId: id(1), clientId: id(4), certificateId: "https://example.vault.azure.net/certificates/reporting/" + "a".repeat(32),
  admins: [admin.objectId], destination: { siteId: `example.sharepoint.com,${id(5)},${id(6)}`,
    siteUrl: "https://example.sharepoint.com/sites/reporting", policyListId: id(7), balanceListId: id(8), jobListId: id(9),
    adminGroupId: 20, maxScopes: 5000, scopeReserve: 250 } };
let store: MemoryStore;
let service: ReportingService;
let reconcile: ReturnType<typeof vi.fn>;
let job: SyncJob | null;
beforeEach(() => {
  store = new MemoryStore(); job = null;
  reconcile = vi.fn().mockResolvedValue({ policies: 2, rows: 4 });
  service = new ReportingService(store, config, true, () => ({ reconcile }), async () => job);
});
describe("server-authoritative reporting", () => {
  it("does not simulate readiness or disclose destination before durable verification", async () => {
    expect(await service.settings(owner)).toMatchObject({ enabled: true, securityReady: false, siteId: "", policyIds: [id(11), id(13)], isAdmin: false });
    expect(await service.settings(admin)).toMatchObject({ securityReady: false, siteId: config.destination!.siteId, policyIds: [], isAdmin: true });
    await expect(service.settings({ ...owner, objectId: id(99) })).rejects.toMatchObject({ status: 403 });
    await expect(service.settings({ ...owner, tenantId: id(99) })).rejects.toMatchObject({ status: 403 });
  });
  it("uses only current metadata and persists checked metadata version and ownership hash", async () => {
    const result = await service.reconcile(admin, {});
    expect(result).toMatchObject({ ready: true, policies: 2, rows: 4 });
    expect(reconcile).toHaveBeenCalledWith(policyOwners(metadata));
    expect(store.value).toMatchObject({ ready: true, metadataEtag: "metadata-v1", ownershipHash: expect.stringMatching(/^[a-f0-9]{64}$/) });
    const settings = await service.settings(owner);
    expect(settings).toMatchObject({ securityReady: true, policyListId: id(7), balanceListId: id(8) });
    expect(settings).not.toHaveProperty("jobListId"); expect(settings).not.toHaveProperty("admins");
    store.snapshot.etag = "metadata-v2";
    expect((await service.settings(owner)).securityReady).toBe(false);
  });
  it("rejects owners, missing OBO configuration and browser-supplied principals/URLs", async () => {
    await expect(service.reconcile(owner, {})).rejects.toMatchObject({ status: 403 });
    const unconfigured = new ReportingService(store, { ...config, destination: null }, true, () => ({ reconcile }), async () => job);
    await expect(unconfigured.reconcile(admin, {})).rejects.toMatchObject({ code: "reporting-not-configured" });
    for (const value of [{ siteUrl: "https://bad.example" }, { owners: [id(10)] }, { jobId: id(10) }, { leaseId: id(10) }]) {
      expect(reportingSecuritySchema.safeParse(value).success).toBe(false);
    }
    expect(reconcile).not.toHaveBeenCalled();
  });
  it("verifies current tenant job, policy and lease both before and after ACL work", async () => {
    const input = { policyId: id(11), jobId: id(20), leaseId: id(21) };
    await expect(service.reconcile(admin, input)).rejects.toMatchObject({ code: "reporting-sync-lease" });
    job = { id: id(20), tenantId: id(1), executorId: admin.objectId, state: "running",
      lease: { id: id(21), expiresAt: new Date(Date.now() + 90000).toISOString() }, policies: [{ id: id(11) }] } as SyncJob;
    await service.reconcile(admin, input);
    reconcile.mockImplementationOnce(async () => { job!.lease!.expiresAt = "2020-01-01T00:00:00Z"; return { policies: 2, rows: 4 }; });
    await expect(service.reconcile(admin, input)).rejects.toMatchObject({ code: "reporting-sync-lease" });
    expect(store.value!.ready).toBe(false);
  });
  it("fails closed on partial revocation, lock loss, or metadata changing out of band", async () => {
    await service.reconcile(admin, {});
    reconcile.mockRejectedValueOnce(new Error("revocation failed"));
    await expect(service.reconcile(admin, {})).rejects.toThrow("revocation failed");
    expect(store.value!.ready).toBe(false);
    reconcile.mockImplementationOnce(async () => { store.snapshot.etag = "operator-changed"; return { policies: 2, rows: 4 }; });
    await expect(service.reconcile(admin, {})).rejects.toMatchObject({ code: "reporting-metadata-conflict" });
    store.assert.mockImplementation(() => { throw new Error("lease lost"); });
    await expect(service.reconcile(admin, {})).rejects.toThrow("lease lost");
    expect(store.value!.ready).toBe(false);
  });
  it("invalidates configuration changes rather than trusting an environment flag", async () => {
    await service.reconcile(admin, {});
    const changed = new ReportingService(store, { ...config, admins: [admin.objectId, id(99)] }, true, () => ({ reconcile }), async () => job);
    expect((await changed.settings(owner)).securityReady).toBe(false);
  });
});
describe("ownership mutations", () => {
  const changed = () => ({ ...metadata, budgets: [{ ...metadata.budgets[0], owners: [id(40)] }] });
  it("revokes all existing owners before changing authority and never marks it verified", async () => {
    await service.reconcile(admin, {});
    const write = vi.fn(async (leaseId?: string) => {
      expect(leaseId).toBe("metadata-lease");
      expect(store.value!.ready).toBe(false);
      expect(reconcile).toHaveBeenLastCalledWith(new Map(), true);
      store.snapshot = { metadata: changed(), etag: "metadata-v2" };
    });
    await service.mutateMetadata(changed(), "metadata-v1", write);
    expect(write).toHaveBeenCalledOnce();
    expect(store.value!.ready).toBe(false);
    await expect(service.settings(owner)).rejects.toMatchObject({ status: 403 });
  });
  it("does not commit metadata when any old permission cannot be revoked/read back", async () => {
    const write = vi.fn();
    reconcile.mockRejectedValueOnce(new Error("some old grants remain"));
    await expect(service.mutateMetadata(changed(), "metadata-v1", write)).rejects.toThrow("some old grants remain");
    expect(write).not.toHaveBeenCalled();
    expect(store.snapshot.etag).toBe("metadata-v1"); expect(store.value!.ready).toBe(false);
  });
  it("blocks stale writes and cannot bypass revocation by removing configuration", async () => {
    const write = vi.fn();
    await expect(service.mutateMetadata(changed(), "stale", write)).rejects.toMatchObject({ code: "metadata-conflict" });
    await service.reconcile(admin, {});
    const disabled = new ReportingService(store, { ...config, enabled: false, destination: null }, false, () => ({ reconcile }), async () => job);
    await expect(disabled.mutateMetadata(changed(), "metadata-v1", write)).rejects.toMatchObject({ code: "reporting-not-configured" });
    expect(write).not.toHaveBeenCalled();
  });
  it("preserves ordinary non-reporting bootstrap but invalidates ready state on unrelated metadata changes", async () => {
    const write = vi.fn();
    const disabled = new ReportingService(store, { ...config, enabled: false, destination: null }, false, () => ({ reconcile }), async () => job);
    await disabled.mutateMetadata(changed(), null, write);
    expect(write).toHaveBeenCalledOnce();
    await service.reconcile(admin, {});
    await service.mutateMetadata(metadata, "metadata-v1", write);
    expect(reconcile).toHaveBeenCalledOnce(); expect(store.value!.ready).toBe(false);
  });
});
describe("explicit prototype app filtering", () => {
  const prototype = () => new ReportingService(store, {
    ...config, accessMode: "prototype", destination: null, certificateId: "",
    snapshotDestination: { siteId: config.destination!.siteId, policyListId: id(7), balanceListId: id(8) },
  }, true, () => { throw new Error("Prototype must never construct the ACL/OBO client."); }, async () => job);
  it("discloses only API-derived owner policies without claiming ACL verification or using OBO", async () => {
    const reporting = prototype();
    expect(await reporting.settings(owner)).toEqual({ enabled: true, accessMode: "prototype",
      siteId: config.destination!.siteId, policyListId: id(7), balanceListId: id(8),
      isAdmin: false, policyIds: [id(11), id(13)], securityReady: false });
    expect(await reporting.settings(admin)).toMatchObject({ isAdmin: true, policyIds: [], securityReady: false });
    await expect(reporting.settings({ ...owner, objectId: id(99) })).rejects.toMatchObject({ status: 403 });
    await expect(reporting.settings({ ...owner, tenantId: id(99) })).rejects.toMatchObject({ status: 403 });
    await expect(reporting.reconcile(admin, {})).rejects.toMatchObject({ code: "reporting-prototype-mode" });
    expect(store.value).toBeNull();
    expect(reconcile).not.toHaveBeenCalled();
  });
  it("allows ownership edits with normal concurrency guards, invalidates ACL readiness and reflects current owners", async () => {
    await service.reconcile(admin, {});
    reconcile.mockClear();
    const reporting = prototype();
    const next = { ...metadata, budgets: [{ ...metadata.budgets[0], owners: [id(40)] }] };
    const write = vi.fn(async (leaseId?: string) => {
      expect(leaseId).toBe("metadata-lease");
      expect(store.value?.ready).toBe(false);
      store.snapshot = { metadata: next, etag: "metadata-v2" };
    });
    await expect(reporting.mutateMetadata(next, "stale", write)).rejects.toMatchObject({ code: "metadata-conflict" });
    expect(write).not.toHaveBeenCalled();
    await reporting.mutateMetadata(next, "metadata-v1", write);
    await expect(reporting.settings(owner)).rejects.toMatchObject({ status: 403 });
    expect(await reporting.settings({ ...owner, objectId: id(40) })).toMatchObject({ policyIds: [id(11), id(13)], securityReady: false });
    expect(reconcile).not.toHaveBeenCalled();
    expect((await service.settings({ ...owner, objectId: id(40) })).securityReady).toBe(false);
  });
  it("requires an explicit valid mode, a complete sync destination and no ACL opt-in", () => {
    const env = { TENANT_ID: id(1), API_AUDIENCE: id(4), ONBOARDING_ADMIN_OBJECT_IDS: admin.objectId,
      SYNC_SITE_ID: config.destination!.siteId, SYNC_POLICY_LIST_ID: id(7), SYNC_BALANCE_LIST_ID: id(8), SYNC_JOB_LIST_ID: id(9) };
    const normal = reportingConfiguration(syncSettings(env), env);
    expect(normal).toMatchObject({ enabled: false, accessMode: "sharepoint-acl", destination: null });
    const opted = { ...env, REPORTING_ACCESS_MODE: "prototype" };
    expect(reportingConfiguration(syncSettings(opted), opted)).toMatchObject({ enabled: true, accessMode: "prototype",
      certificateId: "", destination: null, snapshotDestination: { policyListId: id(7) } });
    expect(() => reportingConfiguration(syncSettings(env), { ...env, REPORTING_ACCESS_MODE: "typo" })).toThrow("Invalid reporting access mode");
    expect(() => reportingConfiguration(syncSettings(opted), { ...opted, REPORTING_SECURITY_ENABLED: "true" })).toThrow("cannot enable");
    const empty = { TENANT_ID: id(1), API_AUDIENCE: id(4), REPORTING_ACCESS_MODE: "prototype" };
    expect(reportingConfiguration(syncSettings(empty), empty).snapshotDestination).toBeNull();
    expect(reportingConfiguration(syncSettings(opted), { ...opted, ONBOARDING_ADMIN_OBJECT_IDS: "" }).snapshotDestination).toBeNull();
  });
  it("does not bypass storage errors or lock loss when editing prototype ownership", async () => {
    const write = vi.fn();
    const save = vi.spyOn(store, "save").mockRejectedValueOnce(new Error("Cannot invalidate ACL marker"));
    await expect(prototype().mutateMetadata(metadata, "metadata-v1", write)).rejects.toThrow("Cannot invalidate ACL marker");
    expect(write).not.toHaveBeenCalled();
    save.mockRestore();
    store.assert.mockImplementation(() => { throw new Error("metadata lease lost"); });
    await expect(prototype().mutateMetadata(metadata, "metadata-v1", write)).rejects.toThrow("metadata lease lost");
    expect(write).not.toHaveBeenCalled();
  });
});
it("requires explicit complete configuration and enforces the 5,000-scope ceiling", () => {
  const env = { TENANT_ID: id(1), API_AUDIENCE: id(4), ONBOARDING_ADMIN_OBJECT_IDS: admin.objectId,
    SYNC_SITE_ID: config.destination!.siteId, SYNC_POLICY_LIST_ID: id(7), SYNC_BALANCE_LIST_ID: id(8), SYNC_JOB_LIST_ID: id(9),
    REPORTING_SITE_URL: config.destination!.siteUrl, REPORTING_ADMIN_GROUP_ID: "20", REPORTING_CERTIFICATE_ID: config.certificateId };
  expect(reportingConfiguration(syncSettings(env), env).destination).toBeNull();
  const opted = { ...env, REPORTING_SECURITY_ENABLED: "true" };
  expect(reportingConfiguration(syncSettings(opted), opted).destination?.maxScopes).toBe(5000);
  expect(() => reportingConfiguration(syncSettings(opted), { ...opted, REPORTING_MAX_SCOPES: "5001" })).toThrow("Invalid REPORTING_MAX_SCOPES");
  expect(() => reportingConfiguration(syncSettings(opted), { ...opted, REPORTING_SITE_URL: "https://bad.example/sites/reporting" })).toThrow();
});
