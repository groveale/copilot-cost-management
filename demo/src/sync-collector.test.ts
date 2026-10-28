import { describe, expect, it, vi } from "vitest";
import { GraphClient, RequestScheduler } from "./graph";
import { aggregateObservations, collectPolicy, observationsFor, policyInventory, probeCollectionAccess, selectedPolicyType } from "./sync-collector";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const descriptor = { id: id(1), displayName: "Policy", type: selectedPolicyType, state: "pending" as const };
const user = { id: id(4), displayName: "User", username: "user@example.test", groups: [id(2)] };
const balance = (serviceId = "cowork", spendingPolicyId = id(1), consumedQuantity = 4, unit = "credit") => ({
  serviceId, spendingPolicyId, consumedQuantity, unit,
  remainingQuantity: { "@odata.type": "#microsoft.graph.unlimitedCostManagementQuantity", unit },
});
const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status });
function graph(options: { denied?: boolean; unsupported?: boolean; invalid?: boolean } = {}) {
  const fetcher = vi.fn<typeof fetch>(async (input) => {
    const url = new URL(String(input)), path = url.pathname;
    if (path.endsWith("spendingPolicies")) return json({ value: [{ id: id(1), displayName: "Policy", "@odata.type": selectedPolicyType }] });
    if (path.endsWith(`/spendingPolicies/${id(1)}`)) return json({ id: id(1), displayName: "Policy",
      "@odata.type": options.unsupported ? "#microsoft.graph.allUsersSpendingPolicy" : selectedPolicyType });
    if (path.endsWith("serviceSettings")) return json({ value: [{ serviceId: "cowork", limit: 10 }] });
    if (path.endsWith("assignedGroups")) return json({ value: [{ id: id(2) }, { id: id(3) }] });
    if (path.endsWith("transitiveMembers")) {
      if (options.denied && path.includes(id(3))) return json({ error: { code: "Denied", message: "Membership unreadable" } }, 403);
      if (url.search) return json({ value: [{ id: id(5), "@odata.type": "#microsoft.graph.user" }] });
      return json({ value: [{ id: id(4), displayName: "User", "@odata.type": "#microsoft.graph.user" }, { id: id(20), "@odata.type": "#microsoft.graph.group" }],
        ...(path.includes(id(2)) ? { "@odata.nextLink": `${url.origin}${path}?page=2` } : {}) });
    }
    if (path.includes(`/userBalances/${id(5)}/`)) return json({ value: [] });
    if (path.endsWith("serviceBalances")) return json({ value: [balance(), balance("workiq", id(9), 100), balance("other", id(1), 2, "token"),
      ...(options.invalid ? [balance()] : [])] });
    throw new Error(`Unexpected ${url.href}`);
  });
  return { fetcher, client: new GraphClient(async () => "synthetic", fetcher, new RequestScheduler(3)) };
}
const signal = () => new AbortController().signal;
describe("policy collection and attribution", () => {
  it("collects nested/overlapping users once, retaining provenance and all services without double counting", async () => {
    const { client, fetcher } = graph();
    const result = await collectPolicy(client, descriptor, 3, signal());
    expect(result.users).toBe(2);
    expect(result.observations).toHaveLength(4);
    expect(result.observations.find((row) => row.userId === id(4))?.groups).toEqual([id(2), id(3)]);
    expect(result.totals).toEqual([{ serviceId: "cowork", unit: "credit", consumed: 4, users: 1 },
      { serviceId: "other", unit: "token", consumed: 2, users: 1 }]);
    expect(result.exclusions).toBe(2);
    expect(fetcher.mock.calls.filter(([url]) => String(url).includes(`/userBalances/${id(4)}/`))).toHaveLength(1);
    expect(result.observations.find((row) => row.userId === id(5))?.match).toBe("no-balance");
  });
  it("never converts unreadable membership or ambiguous balances into a successful snapshot", async () => {
    await expect(collectPolicy(graph({ denied: true }).client, descriptor, 3, signal())).rejects.toThrow("unreadable");
    await expect(collectPolicy(graph({ invalid: true }).client, descriptor, 3, signal())).rejects.toThrow("ambiguous");
  });
  it("inventories unsupported policies without a directory scan or zero totals", async () => {
    const { client, fetcher } = graph({ unsupported: true });
    const result = await collectPolicy(client, descriptor, 3, signal());
    expect(result.unsupported).toBe(true);
    expect(result.observations).toEqual([]);
    expect(fetcher).toHaveBeenCalledTimes(2);
    expect(result.services).toEqual([{ serviceId: "cowork", limit: 10 }]);
    expect(fetcher.mock.calls.some(([url]) => String(url).includes("/groups/") || String(url).includes("/userBalances/"))).toBe(false);
  });
  it("does not imply authorization for probe routes without real resources", async () => {
    const client = new GraphClient(async () => "t", vi.fn(async () => json({ value: [] })));
    expect(await probeCollectionAccess(client, [], signal())).toContain("Assigned-group, membership and balance access: untested (no supported policy)");
    expect(await policyInventory(client, signal())).toEqual([]);
  });
  it("validates quantities, duplicate services, timestamps, and policy matches", () => {
    expect(observationsFor(user, id(1), [balance()])[0].remainingKind).toBe("unlimited");
    expect(observationsFor(user, id(1), [{ ...balance(), remainingQuantity: undefined }])[0].remainingKind).toBe("missing");
    expect(() => observationsFor(user, id(1), [balance(), balance()])).toThrow("ambiguous");
    expect(() => observationsFor(user, id(1), [balance("x", id(1), -1)])).toThrow("quantity");
    expect(() => observationsFor(user, id(1), [{ ...balance(), dataStatus: { lastModifiedDateTime: "not-a-date" } }])).toThrow("timestamp");
    expect(() => observationsFor(user, id(1), [{ ...balance(), remainingQuantity: { unit: "other" } }])).toThrow("unit");
    const rows = observationsFor(user, id(1), [balance()]);
    expect(() => aggregateObservations([...rows, ...rows])).toThrow("Duplicate");
  });
});
describe("measured shared Graph concurrency", () => {
  it("honors the configured ceiling across reads and writes while actually running in parallel", async () => {
    let active = 0, maximum = 0;
    const fetcher = vi.fn<typeof fetch>(async () => {
      active++; maximum = Math.max(maximum, active);
      await new Promise((resolve) => setTimeout(resolve, 5));
      active--; return json({ ok: true });
    });
    const client = new GraphClient(async () => "t", fetcher, new RequestScheduler(3));
    await Promise.all(Array.from({ length: 20 }, (_, n) =>
      n % 2 ? client.get("/v1.0/example", signal()) : client.post("/v1.0/example", {}, signal())));
    expect(maximum).toBe(3); expect(active).toBe(0);
    expect(() => new RequestScheduler(9)).toThrow("1-8");
  });
});
