import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import type { Metadata } from "../../api/src/contracts";
import { loadPolicyInventory, PolicyInventoryTable } from "./AdminPolicies";
import { GraphClient } from "./graph";

const id = (n: number) => `10000000-0000-0000-0000-${String(n).padStart(12, "0")}`;
const metadata: Metadata = { version: 1, tenantId: id(1), serviceId: "cowork", budgets: [{
  id: "legal", displayName: "Legal", businessArea: "Legal operations", costCenter: "LEGAL-001", owners: [id(7), id(8)],
  tiers: { standard: { policyId: id(3), groupId: id(4) }, enhanced: { policyId: id(5), groupId: id(6) } },
}] };
const policy = (n: number, displayName: string) => ({ id: id(n), displayName });
const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status });

function setup(options: { uninitialized?: boolean; deniedOwner?: boolean; deniedPage?: boolean; malformed?: boolean; repeated?: boolean } = {}) {
  const fetcher = vi.fn<typeof fetch>(async (input) => {
    const url = new URL(String(input));
    if (url.pathname.startsWith("/v1.0/users/")) {
      if (options.deniedOwner && url.pathname.includes(id(8))) return json({ error: { code: "AccessDenied", message: "Owner name lookup denied" } }, 403);
      const ownerId = url.pathname.split("/").at(-1);
      return json({ id: ownerId, displayName: `Mapped owner ${ownerId === id(7) ? "one" : "two"}`, userPrincipalName: `owner${ownerId === id(7) ? "1" : "2"}@example.test` });
    }
    if (url.searchParams.get("page") === "2") {
      if (options.deniedPage) return json({ error: { code: "AccessDenied", message: "Second policy page denied" } }, 403);
      return json({ value: [policy(5, "Legal Enhanced"), policy(9, "Unregistered policy"),
        ...(options.repeated ? [policy(3, "Legal Standard")] : [])] });
    }
    return json({ value: options.malformed ? [{ id: "not-a-guid", displayName: "Invalid" }] : [policy(3, "Legal Standard")],
      "@odata.nextLink": "https://graph.microsoft.com/beta/copilot/costManagement/spendingPolicies?page=2" });
  });
  return { fetcher, load: (signal = new AbortController().signal) =>
    loadPolicyInventory(new GraphClient(async () => "synthetic-token", fetcher), options.uninitialized ? null : metadata, signal) };
}

describe("administrator Graph policy inventory", () => {
  it("lists every Graph page, joins owners by policy ID, and retains unmapped policies", async () => {
    const { load, fetcher } = setup();
    const inventory = await load();
    expect(inventory.policies).toHaveLength(3);
    expect(inventory.policies.find((row) => row.id === id(3))?.mappings).toMatchObject([{ tier: "standard", budget: { id: "legal" } }]);
    expect(inventory.policies.find((row) => row.id === id(5))?.mappings).toMatchObject([{ tier: "enhanced", budget: { owners: [id(7), id(8)] } }]);
    expect(inventory.policies.find((row) => row.id === id(9))?.mappings).toEqual([]);
    expect(fetcher.mock.calls.filter(([url]) => String(url).includes("/users/"))).toHaveLength(2);
    expect(fetcher.mock.calls.every(([, init]) => init?.method === undefined)).toBe(true);
    const html = renderToStaticMarkup(<PolicyInventoryTable inventory={inventory} />);
    for (const value of ["Legal Standard", "Legal Enhanced", "Unregistered policy", "Unmapped", "Mapped owner one",
      "Mapped owner two", "owner1@example.test", "LEGAL-001", "Legal operations", "3 Graph policies. 1 unmapped."]) {
      expect(html).toContain(value);
    }
  });
  it("does not invent owner mappings before metadata is initialized", async () => {
    const { load, fetcher } = setup({ uninitialized: true });
    const inventory = await load();
    expect(inventory.policies.every((row) => !row.mappings.length)).toBe(true);
    expect(fetcher.mock.calls.some(([url]) => String(url).includes("/users/"))).toBe(false);
    const html = renderToStaticMarkup(<PolicyInventoryTable inventory={inventory} />);
    expect(html).toContain("Initial private metadata has not been imported");
    expect(html).toContain("3 Graph policies. 3 unmapped.");
  });
  it("keeps explicit owner IDs and lookup errors without hiding the policy list", async () => {
    const { load } = setup({ deniedOwner: true });
    const inventory = await load();
    expect(inventory.owners[id(8)]).toMatchObject({ state: "error", error: expect.stringContaining("Owner name lookup denied") });
    const html = renderToStaticMarkup(<PolicyInventoryTable inventory={inventory} />);
    expect(html).toContain(id(8));
    expect(html).toContain("Owner name unavailable");
    expect(html).toContain("Owner name lookup denied");
    expect(html).toContain("Unregistered policy");
  });
  it("never reports a partial first page as a successful complete inventory", async () => {
    await expect(setup({ deniedPage: true }).load()).rejects.toThrow("Second policy page denied");
  });
  it("rejects malformed policy IDs explicitly", async () => {
    await expect(setup({ malformed: true }).load()).rejects.toThrow("Spending policy ID must be a non-zero GUID");
  });
  it("deduplicates policies repeated across pages", async () => {
    expect((await setup({ repeated: true }).load()).policies).toHaveLength(3);
  });
  it("normalizes the policy ID before matching metadata", async () => {
    const upper = "AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE";
    const input: Metadata = { ...metadata, budgets: [{ ...metadata.budgets[0],
      tiers: { ...metadata.budgets[0].tiers, standard: { ...metadata.budgets[0].tiers.standard, policyId: upper.toLowerCase() } } }] };
    const graph = new GraphClient(async () => "synthetic", vi.fn<typeof fetch>(async (url) =>
      String(url).includes("/users/") ? json({ id: String(url).split("/users/")[1].split("?")[0], displayName: "Owner" }) :
        json({ value: [{ id: upper, displayName: "Case-insensitive policy" }] })));
    const result = await loadPolicyInventory(graph, input, new AbortController().signal);
    expect(result.policies[0].mappings).toHaveLength(1);
  });
  it("keeps empty Graph inventories distinct from a failure", async () => {
    const graph = new GraphClient(async () => "synthetic", vi.fn<typeof fetch>(async () => json({ value: [] })));
    const inventory = await loadPolicyInventory(graph, metadata, new AbortController().signal);
    expect(renderToStaticMarkup(<PolicyInventoryTable inventory={inventory} />)).toContain("No spending policies were returned by Graph");
  });
  it("honors cancellation without making a Graph request", async () => {
    const { load, fetcher } = setup();
    const controller = new AbortController();
    controller.abort();
    await expect(load(controller.signal)).rejects.toBeDefined();
    expect(fetcher).not.toHaveBeenCalled();
  });
});
