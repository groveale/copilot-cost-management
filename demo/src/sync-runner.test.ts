import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import type { SyncJob, SyncSettings, SyncUpdate } from "../../api/src/sync-contracts";
import type { ReportingSettings } from "../../api/src/reporting-contracts";
import { RequestApi } from "./api";
import { DemoAuth } from "./auth";
import { GraphError } from "./graph";
import { runSync } from "./sync-runner";

const mocks = vi.hoisted(() => ({
  token: vi.fn(), inventory: vi.fn(), probe: vi.fn(), collect: vi.fn(), validate: vi.fn(), links: vi.fn(),
  mirror: vi.fn(), recover: vi.fn(), fence: vi.fn(), publish: vi.fn(), cleanup: vi.fn(), absent: vi.fn(), security: vi.fn<RequestApi["reportingReconcile"]>(),
  settings: vi.fn<RequestApi["reportingSettings"]>(),
}));
vi.mock("./auth", () => ({
  DemoAuth: class { syncToken = mocks.token; },
  isGraphInteractionRequired: () => false,
}));
vi.mock("./sync-collector", () => ({
  policyInventory: mocks.inventory, probeCollectionAccess: mocks.probe, collectPolicy: mocks.collect,
}));
vi.mock("./sync-sharepoint", () => ({
  SharePointSync: class {
    validate = mocks.validate; links = mocks.links; mirror = mocks.mirror; recover = mocks.recover;
    fence = mocks.fence; publish = mocks.publish; cleanupPolicy = mocks.cleanup; markAbsent = mocks.absent;
  },
}));
const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const policies = [1, 2].map((n) => ({ id: id(n), displayName: `Policy ${n}`, type: "selected", state: "pending" as const }));
const original: SyncJob = { version: 1, revision: 1, id: id(3), tenantId: id(4), initiatorId: id(5), actors: [id(5)], concurrency: 3,
  state: "paused", phase: "Paused", createdAt: "2026-10-28T13:00:00Z", updatedAt: "2026-10-28T13:00:00Z",
  inventoryComplete: false, policies: [], mirrorPending: true };
const settings: SyncSettings = { enabled: true, siteId: "example.sharepoint.com,a,b", policyListId: id(6), balanceListId: id(7),
  jobListId: id(8), defaultConcurrency: 3, maxConcurrency: 8, retentionDays: 30 };
const account: AccountInfo = { homeAccountId: "home", localAccountId: id(5), tenantId: id(4), environment: "login.microsoftonline.com", username: "admin@example.test" };
class Api extends RequestApi {
  job = structuredClone(original);
  updates: SyncUpdate[] = [];
  constructor() { super("", async () => "unused"); }
  override async onboardingCapabilities() { return { isAdmin: true, bootstrapEnabled: false, objectId: id(5) }; }
  override async onboardingSetup() { return null; }
  override reportingReconcile = mocks.security;
  override reportingSettings = mocks.settings;
  override async syncUpdate(_id: string, input: SyncUpdate) {
    this.updates.push(input);
    if (input.action === "claim") { this.job.state = "running"; this.job.lease = { id: id(9), expiresAt: "2099-01-01T00:00:00Z" }; }
    if (input.action === "inventory") {
      this.job.inventoryComplete = true; this.job.policies = input.policies.map((row) => ({ ...row, state: "pending" }));
    }
    if (input.action === "policy") this.job.policies = this.job.policies.map((row) => row.id === input.policy.id ? input.policy : row);
    if (input.action === "pause") { this.job.state = input.accessRequired ? "access-required" : "paused"; this.job.lease = undefined; }
    if (input.action === "finish") {
      this.job.state = this.job.policies.some((row) => row.state === "failed") ? "partially-failed" : "completed";
      this.job.lease = undefined;
    }
    if (input.action === "mirrored") this.job.mirrorPending = false;
    else { this.job.revision++; this.job.mirrorPending = true; }
    return structuredClone(this.job);
  }
}
const auth = () => new DemoAuth({ version: 1, tenantId: id(4), clientId: id(10), serviceId: "cowork", budgets: [] });
beforeEach(() => {
  vi.resetAllMocks(); vi.useFakeTimers();
  mocks.token.mockResolvedValue("synthetic");
  mocks.inventory.mockResolvedValue(policies);
  mocks.probe.mockResolvedValue(["Live probe"]);
  mocks.validate.mockResolvedValue(["Lists verified"]);
  mocks.security.mockResolvedValue({ ready: true, checkedAt: original.createdAt, policies: 2, rows: 4 });
  mocks.settings.mockResolvedValue({ enabled: true, accessMode: "sharepoint-acl",
    siteId: settings.siteId, policyListId: settings.policyListId, balanceListId: settings.balanceListId,
    isAdmin: true, policyIds: [], securityReady: true });
  mocks.links.mockResolvedValue([]);
  mocks.mirror.mockResolvedValue(undefined);
  mocks.collect.mockResolvedValue({ users: 1, groups: [], exclusions: 0, unsupported: false });
  mocks.publish.mockImplementation(async (_fence, _snapshot, _policy, _job, _metadata, _signal, checkLease) => { await checkLease(); });
});
afterEach(() => vi.useRealTimers());
const run = (api: Api, controller = new AbortController()) =>
  runSync(api, auth(), account, settings, api.job, controller.signal, vi.fn(), vi.fn());
describe("browser sync orchestration", () => {
  it("collects and publishes in explicit prototype mode without any ACL request while retaining lease heartbeats", async () => {
    mocks.settings.mockResolvedValue({ enabled: true, accessMode: "prototype",
      siteId: settings.siteId, policyListId: settings.policyListId, balanceListId: settings.balanceListId,
      isAdmin: true, policyIds: [], securityReady: false });
    mocks.security.mockRejectedValue(new Error("ACL service is deliberately unavailable"));
    const api = new Api();
    expect((await run(api)).state).toBe("completed");
    expect(mocks.publish).toHaveBeenCalledTimes(2);
    expect(mocks.security).not.toHaveBeenCalled();
    expect(api.updates.filter((input) => input.action === "heartbeat").length).toBeGreaterThanOrEqual(6);
  });
  it("stops publication when access mode changes during a prototype run", async () => {
    const base: ReportingSettings = { enabled: true, accessMode: "prototype",
      siteId: settings.siteId, policyListId: settings.policyListId, balanceListId: settings.balanceListId,
      isAdmin: true, policyIds: [], securityReady: false };
    mocks.settings.mockResolvedValueOnce(base).mockResolvedValue({ ...base, accessMode: "sharepoint-acl" });
    const api = new Api();
    await expect(run(api)).rejects.toThrow("access mode changed");
    expect(api.job.state).toBe("paused");
    expect(api.updates.some((input) => input.action === "policy" && input.policy.state === "completed")).toBe(false);
  });
  it("does not bypass checks for disabled, invalid or mismatched reporting configuration", async () => {
    const base = { enabled: true, accessMode: "prototype",
      siteId: settings.siteId, policyListId: settings.policyListId, balanceListId: settings.balanceListId,
      isAdmin: true, policyIds: [], securityReady: false };
    for (const change of [{ enabled: false }, { isAdmin: false }, { siteId: "wrong-site" }, { accessMode: "typo" }]) {
      const transport = vi.fn<typeof fetch>(async () => new Response(JSON.stringify({ ...base, ...change })));
      const api = new Api();
      mocks.settings.mockImplementation((signal) => new RequestApi("", async () => "unused", transport).reportingSettings(signal));
      await expect(run(api)).rejects.toThrow();
      expect(api.updates).toEqual([]);
    }
    expect(mocks.security).not.toHaveBeenCalled();
    expect(mocks.collect).not.toHaveBeenCalled();
  });
  it("freshly validates access, checkpoints every policy and publishes independent outcomes", async () => {
    const api = new Api();
    const result = await run(api);
    expect(result.state).toBe("completed");
    expect(mocks.token).toHaveBeenCalledWith(account, true);
    expect(mocks.publish).toHaveBeenCalledTimes(2);
    expect(mocks.security).toHaveBeenCalledTimes(3);
    expect(mocks.security).toHaveBeenNthCalledWith(2, expect.any(AbortSignal), { policyId: policies[0].id, jobId: original.id, leaseId: id(9) });
    expect(mocks.absent).toHaveBeenCalledTimes(1);
    expect(api.updates.filter((input) => input.action === "policy" && input.policy.state === "completed")).toHaveLength(2);
    expect(vi.getTimerCount()).toBe(0);
  });
  it("records policy-local failures without overwriting their snapshots or blocking other policies", async () => {
    mocks.collect.mockRejectedValueOnce(new Error("Malformed service row"));
    const api = new Api(); const result = await run(api);
    expect(result.state).toBe("partially-failed");
    expect(result.policies.map((policy) => policy.state)).toEqual(["failed", "completed"]);
    expect(mocks.publish).toHaveBeenCalledTimes(1);
  });
  it("pauses explicitly on PIM/access loss and does not publish incomplete work", async () => {
    mocks.collect.mockRejectedValueOnce(new GraphError("PIM or permission required", 403, "Denied", null));
    const api = new Api();
    await expect(run(api)).rejects.toThrow("permission required");
    expect(api.job.state).toBe("access-required");
    expect(mocks.publish).not.toHaveBeenCalled();
    expect(vi.getTimerCount()).toBe(0);
  });
  it("recovers already published policy generations after a lost checkpoint acknowledgement", async () => {
    mocks.recover.mockResolvedValueOnce({ ...policies[0], state: "completed", users: 1 });
    const api = new Api();
    await run(api);
    expect(mocks.collect).toHaveBeenCalledTimes(1);
    expect(mocks.publish).toHaveBeenCalledTimes(1);
  });
  it("never marks unrelated policies stale during a single-policy sync", async () => {
    const api = new Api(); api.job.policyId = policies[0].id;
    mocks.inventory.mockResolvedValue([policies[0]]);
    await run(api);
    expect(mocks.absent).not.toHaveBeenCalled();
  });
  it("does not claim a task or imply success when live preflight fails", async () => {
    mocks.probe.mockRejectedValue(new GraphError("No role", 403, "Denied", null));
    const api = new Api();
    await expect(run(api)).rejects.toThrow("No role");
    expect(api.updates).toEqual([]);
    expect(vi.getTimerCount()).toBe(0);
  });
  it("does not collect or claim work when SharePoint row isolation cannot be verified", async () => {
    mocks.security.mockRejectedValue(new Error("SharePoint permission verification failed"));
    const api = new Api();
    await expect(run(api)).rejects.toThrow("permission verification failed");
    expect(api.updates).toEqual([]);
    expect(mocks.collect).not.toHaveBeenCalled();
    expect(mocks.inventory).not.toHaveBeenCalled();
  });
  it("pauses without a completed policy checkpoint when staged-row permissions fail", async () => {
    mocks.security.mockResolvedValueOnce({ ready: true, checkedAt: original.createdAt, policies: 2, rows: 0 })
      .mockRejectedValueOnce(new Error("Owner ACL could not be verified"));
    const api = new Api();
    await expect(run(api)).rejects.toThrow("Owner ACL could not be verified");
    expect(api.job.state).toBe("paused");
    expect(api.updates.some((input) => input.action === "policy" && input.policy.state === "completed")).toBe(false);
    expect(vi.getTimerCount()).toBe(0);
  });
});
