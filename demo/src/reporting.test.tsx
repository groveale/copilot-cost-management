import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import type { ReportingSettings } from "../../api/src/reporting-contracts";
import { RequestApi } from "./api";
import { GraphClient } from "./graph";
import { SharePointReporting } from "./reporting";
import { loadSnapshotReports, SnapshotPolicyPanel } from "./SnapshotReports";
import { ReportingAccessNotice } from "./ReportingAccessNotice";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const settings: ReportingSettings = { enabled: true, accessMode: "sharepoint-acl", siteId: `example.sharepoint.com,${id(2)},${id(3)}`,
  policyListId: id(4), balanceListId: id(5), isAdmin: true, policyIds: [], securityReady: true };
const tenant = id(1), policyId = id(6), snapshotId = id(7), userId = id(8), jobId = id(9), groupId = id(10);
const time = "2026-11-04T12:00:00Z";
const policyKey = `${tenant}:${policyId}`;
function fixture() {
  const master: Record<string, unknown> = {
    SchemaVersion: 1, TenantId: tenant, PolicyId: policyId, PolicyKey: policyKey, DisplayName: "Legal Standard",
    ActiveSnapshotId: snapshotId, JobId: jobId, Presence: "present", StartedAt: time, ObservedAt: time,
    Configuration: JSON.stringify({ policy: { id: policyId, displayName: "Legal Standard" }, services: [{ serviceId: "cowork", state: "enabled" }],
      groups: [{ id: groupId, displayName: "Legal" }], mappings: [{ budgetId: "legal", owners: [id(11)] }] }),
    Coverage: JSON.stringify({ users: 1, groups: 1, exclusions: 0, observations: 1, unsupported: false, reason: null }),
    Totals: JSON.stringify([{ serviceId: "cowork", unit: "credit", consumed: 7, users: 1 }]),
  };
  const observation: Record<string, unknown> = {
    SchemaVersion: 1, TenantId: tenant, PolicyId: policyId, PolicyKey: policyKey, SnapshotId: snapshotId,
    SnapshotKey: `${policyKey}:${snapshotId}`, JobId: jobId, UserId: userId, ServiceId: "cowork", DisplayName: "Private user",
    UserPrincipalName: "private@example.test", Groups: JSON.stringify([groupId]), ReportedPolicyId: policyId, PolicyMatch: "matched",
    ConsumedQuantity: 7, Unit: "credit", RemainingKind: "missing", RemainingValue: null, ConsumptionAllowed: "not-reported",
    ObservedAt: time, SourceTime: null, AccessReason: null,
  };
  const rows = new Map<string, Record<string, unknown>[]>([[settings.policyListId, [master]], [settings.balanceListId, [observation]]]);
  const transport = vi.fn<typeof fetch>(async (input) => {
    const url = new URL(String(input));
    if (!url.pathname.startsWith("/v1.0/sites/") || !url.pathname.endsWith("/items")) throw new Error("Sync reporting must only call SharePoint list APIs.");
    const list = /\/lists\/([^/]+)/.exec(url.pathname)![1];
    const [, field, value] = /fields\/(\w+) eq '(.*)'/.exec(url.searchParams.get("$filter")!)!;
    return new Response(JSON.stringify({ value: (rows.get(list) ?? []).filter((row) => row[field] === value)
      .map((fields, index) => ({ id: String(index + 1), fields })) }));
  });
  const graph = new GraphClient(async () => "synthetic-token", transport);
  return { master, observation, rows, transport, graph, reader: new SharePointReporting(graph, settings, tenant) };
}
const signal = () => new AbortController().signal;
describe("delegated active-generation reporting", () => {
  it("uses verified Title names rather than a missing or stale custom DisplayName field", async () => {
    const { reader, master, observation } = fixture();
    master.Title = "Current policy"; observation.Title = "Current user";
    expect((await reader.read(policyId, signal()))).toMatchObject({
      policy: { displayName: "Current policy" }, observations: [{ displayName: "Current user" }],
    });
    delete master.DisplayName; delete observation.DisplayName;
    expect((await reader.read(policyId, signal()))).toMatchObject({
      policy: { displayName: "Current policy" }, observations: [{ displayName: "Current user" }],
    });
    observation.Title = userId;
    expect((await reader.read(policyId, signal())).observations[0].displayName).toBeNull();
  });
  it("allows prototype owner reads without ACL readiness but retains policy filtering and authorization checks", async () => {
    const { graph, transport } = fixture();
    const prototype: ReportingSettings = { ...settings, accessMode: "prototype", securityReady: false, isAdmin: false, policyIds: [policyId] };
    const reader = new SharePointReporting(graph, prototype, tenant);
    expect((await reader.inventory(signal())).map((row) => row.id)).toEqual([policyId]);
    expect((await reader.read(policyId, signal())).observations).toHaveLength(1);
    expect(transport.mock.calls.every(([url]) =>
      /fields\/(?:PolicyKey|SnapshotKey) eq '/.test(new URL(String(url)).searchParams.get("$filter")!))).toBe(true);
    transport.mockClear();
    await expect(reader.read(id(99), signal())).rejects.toThrow("not authorized");
    expect(transport).not.toHaveBeenCalled();
    expect(renderToStaticMarkup(<ReportingAccessNotice accessMode="prototype" />)).toContain("read other policies and user balances directly");
    expect(renderToStaticMarkup(<ReportingAccessNotice accessMode="sharepoint-acl" />)).toBe("");
  });
  it("rejects foreign policy rows even when a prototype list query returns them", async () => {
    const { master } = fixture();
    const transport = vi.fn<typeof fetch>(async () => new Response(JSON.stringify({
      value: [{ fields: { ...master, PolicyId: id(99), PolicyKey: `${tenant}:${id(99)}` } }],
    })));
    const reader = new SharePointReporting(new GraphClient(async () => "token", transport),
      { ...settings, accessMode: "prototype", securityReady: false, isAdmin: false, policyIds: [policyId] }, tenant);
    await expect(reader.inventory(signal())).rejects.toThrow("outside your authorized workspace");
  });
  it("defaults older settings to ACL enforcement and rejects unknown modes before any SharePoint request", async () => {
    const { graph, transport } = fixture();
    for (const mode of [undefined, "typo"]) {
      const api = new RequestApi("", async () => "api", vi.fn<typeof fetch>(async () => new Response(JSON.stringify({
        ...settings, isAdmin: false, securityReady: false, policyIds: [policyId], accessMode: mode,
      }))));
      await expect(loadSnapshotReports(api, graph, tenant, signal())).rejects.toThrow(
        mode === undefined ? "not been verified" : "Invalid reporting access mode");
    }
    expect(transport).not.toHaveBeenCalled();
  });
  it("reports prototype mode to the UI even with an empty inventory and still rejects personal owner reports", async () => {
    const { graph, rows, transport } = fixture();
    rows.set(settings.policyListId, []);
    const prototype: ReportingSettings = { ...settings, accessMode: "prototype", isAdmin: false, policyIds: [policyId], securityReady: false };
    const api = new RequestApi("", async () => "api", vi.fn<typeof fetch>(async () => new Response(JSON.stringify(prototype))));
    const onSettings = vi.fn();
    expect(await loadSnapshotReports(api, graph, tenant, signal(), { inventoryOnly: true, onSettings })).toEqual([]);
    expect(onSettings).toHaveBeenCalledWith(prototype);
    transport.mockClear();
    await expect(loadSnapshotReports(api, graph, tenant, signal(), { userId })).rejects.toThrow("administrators only");
    expect(transport).not.toHaveBeenCalled();
  });
  it("reads only the active generation and renders normalized observations without claiming a live/full response", async () => {
    const { reader, rows, observation, transport } = fixture();
    rows.get(settings.balanceListId)!.push({ ...observation, SnapshotId: id(30), SnapshotKey: `${policyKey}:${id(30)}`, ConsumedQuantity: 900 });
    const reading = await reader.read(policyId, signal());
    expect(reading.observations).toHaveLength(1);
    expect(reading.observations[0]).toMatchObject({ consumed: 7, allowed: null, remainingKind: "missing" });
    expect(transport.mock.calls).toHaveLength(3);
    const html = renderToStaticMarkup(<SnapshotPolicyPanel report={reading} />);
    for (const value of ["SHAREPOINT SNAPSHOT", "not a real-time reading", "normalized snapshots, not full Graph responses",
      "nested users", "7 credit", "unavailable here"]) expect(html).toContain(value);
    expect(html).not.toContain("Full Graph balance response");
    const personal = renderToStaticMarkup(<SnapshotPolicyPanel report={reading} personal />);
    expect(personal).toContain("Policy-wide consumption totals are not shown");
    expect(personal).not.toContain("matched consumption</dt>");
  });
  it("checks all pages rather than reporting a partial first page", async () => {
    const { master, observation } = fixture();
    master.Coverage = JSON.stringify({ users: 2, groups: 1, exclusions: 0, observations: 2, unsupported: false });
    master.Totals = JSON.stringify([{ serviceId: "cowork", unit: "credit", consumed: 14, users: 2 }]);
    const transport = vi.fn<typeof fetch>(async (input) => {
      const url = new URL(String(input));
      const balance = url.pathname.includes(settings.balanceListId);
      return new Response(JSON.stringify(!balance ? { value: [{ fields: master }] } :
        url.searchParams.has("page") ? { value: [{ fields: { ...observation, UserId: id(20) } }] } :
          { value: [{ fields: observation }], "@odata.nextLink": `https://graph.microsoft.com${url.pathname}?page=2` }));
    });
    const reading = await new SharePointReporting(new GraphClient(async () => "token", transport), settings, tenant).read(policyId, signal());
    expect(reading.observations).toHaveLength(2);
    expect(transport).toHaveBeenCalledTimes(4);
  });
  it("retries a generation swap during cleanup instead of returning an empty snapshot", async () => {
    const { master, observation } = fixture();
    let reads = 0;
    const next = id(40);
    const graph = new GraphClient(async () => "token", vi.fn<typeof fetch>(async (input) => {
      const url = new URL(String(input));
      if (url.pathname.includes(settings.policyListId)) {
        reads++;
        return new Response(JSON.stringify({ value: [{ fields: { ...master, ActiveSnapshotId: reads === 1 ? snapshotId : next } }] }));
      }
      return new Response(JSON.stringify({ value: url.searchParams.get("$filter")!.includes(next) ?
        [{ fields: { ...observation, SnapshotId: next, SnapshotKey: `${policyKey}:${next}` } }] : [] }));
    }));
    const reading = await new SharePointReporting(graph, settings, tenant).read(policyId, signal());
    expect(reading.policy.activeSnapshotId).toBe(next);
    expect(reading.observations).toHaveLength(1);
  });
  it("rejects inaccessible rows, malformed totals and duplicates instead of substituting zero", async () => {
    const missing = fixture(); missing.rows.set(settings.balanceListId, []);
    await expect(missing.reader.read(policyId, signal())).rejects.toThrow("incomplete or inaccessible");
    const invalid = fixture(); invalid.master.Totals = JSON.stringify([{ serviceId: "cowork", unit: "credit", consumed: 900, users: 1 }]);
    await expect(invalid.reader.read(policyId, signal())).rejects.toThrow("totals do not match");
    const duplicate = fixture(); duplicate.rows.get(settings.balanceListId)!.push(duplicate.observation);
    duplicate.master.Coverage = JSON.stringify({ users: 1, groups: 1, exclusions: 0, observations: 2, unsupported: false });
    await expect(duplicate.reader.read(policyId, signal())).rejects.toThrow("Duplicate user/service");
  });
  it("accepts fractional summation-order precision without accepting different totals", async () => {
    const { master, observation, rows, reader } = fixture();
    observation.ConsumedQuantity = 0.1;
    rows.get(settings.balanceListId)!.push({ ...observation, UserId: id(20), ConsumedQuantity: 0.2 });
    master.Coverage = JSON.stringify({ users: 2, groups: 1, exclusions: 0, observations: 2, unsupported: false });
    master.Totals = JSON.stringify([{ serviceId: "cowork", unit: "credit", consumed: 0.3, users: 2 }]);
    expect((await reader.read(policyId, signal())).observations).toHaveLength(2);
    master.Totals = JSON.stringify([{ serviceId: "cowork", unit: "credit", consumed: 0.300001, users: 2 }]);
    await expect(reader.read(policyId, signal())).rejects.toThrow("totals do not match");
  });
  it("rejects another tenant and prevents unauthorized policy requests before sending Graph tokens", async () => {
    const wrong = fixture(); wrong.master.TenantId = id(50);
    await expect(wrong.reader.read(policyId, signal())).rejects.toThrow("tenant/key");
    const { graph, transport } = fixture();
    const owner = new SharePointReporting(graph, { ...settings, isAdmin: false, policyIds: [id(20)] }, tenant);
    await expect(owner.read(policyId, signal())).rejects.toThrow("not authorized");
    expect(transport).not.toHaveBeenCalled();
    expect(() => new SharePointReporting(graph, { ...settings, isAdmin: false, securityReady: false }, tenant)).toThrow("not been verified");
  });
  it("distinguishes no publication, unsupported inventory and stale retained snapshots", async () => {
    const unpublished = fixture(); unpublished.master.ActiveSnapshotId = null;
    const html = renderToStaticMarkup(<SnapshotPolicyPanel report={await unpublished.reader.read(policyId, signal())} />);
    expect(html).toContain("No successful snapshot");
    const stale = fixture(); stale.master.Presence = "stale";
    expect(renderToStaticMarkup(<SnapshotPolicyPanel report={await stale.reader.read(policyId, signal())} />)).toContain("Stale policy");
    const unsupported = fixture();
    unsupported.master.Coverage = JSON.stringify({ users: 0, groups: 0, exclusions: 0, observations: 0, unsupported: true, reason: "Tenant-wide policy" });
    unsupported.rows.set(settings.balanceListId, []);
    const reading = await unsupported.reader.read(policyId, signal());
    expect(reading.policy.totals).toBeNull();
    expect(renderToStaticMarkup(<SnapshotPolicyPanel report={reading} />)).toContain("Inventory only");
  });
  it("loads every reporting shape without any live Cost Management or directory API call", async () => {
    const { graph, transport } = fixture();
    const api = new RequestApi("", async () => "api", vi.fn<typeof fetch>(async () => new Response(JSON.stringify(settings))));
    expect(await loadSnapshotReports(api, graph, tenant, signal(), { inventoryOnly: true })).toHaveLength(1);
    expect(await loadSnapshotReports(api, graph, tenant, signal(), { policyIds: [policyId] })).toHaveLength(1);
    const personal = await loadSnapshotReports(api, graph, tenant, signal(), { userId: id(99) });
    expect(personal[0].observations).toEqual([]);
    const selectedPersonal = await loadSnapshotReports(api, graph, tenant, signal(), { policyIds: [policyId], userId: id(99) });
    expect(selectedPersonal[0].observations).toEqual([]);
    expect(transport.mock.calls.every(([url]) => !/costManagement|\/users\/|\/groups\//.test(String(url)))).toBe(true);
    const nonAdminApi = new RequestApi("", async () => "api", vi.fn<typeof fetch>(async () =>
      new Response(JSON.stringify({ ...settings, isAdmin: false, policyIds: [policyId] }))));
    transport.mockClear();
    await expect(loadSnapshotReports(nonAdminApi, graph, tenant, signal(), { userId })).rejects.toThrow("administrators only");
    expect(transport).not.toHaveBeenCalled();
  });
  it("honors abort without a reporting request", async () => {
    const { reader, transport } = fixture();
    const controller = new AbortController(); controller.abort();
    await expect(reader.read(policyId, controller.signal)).rejects.toBeDefined();
    expect(transport).not.toHaveBeenCalled();
  });
});
