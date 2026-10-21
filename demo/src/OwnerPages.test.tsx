import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, describe, expect, it, vi } from "vitest";
import { RequestApi, type Context } from "./api";
import { DemoAuth } from "./auth";
import { parseConfig } from "./config";
import { BusinessUnitList, OwnerNavigation, OwnerWorkspace, ownerLocation, type OwnerLocation } from "./HostedPortal";

vi.mock("@azure/msal-browser", async (importOriginal) => ({
  ...await importOriginal<typeof import("@azure/msal-browser")>(),
  PublicClientApplication: class {},
}));
afterEach(() => vi.unstubAllGlobals());
const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const config = parseConfig({
  version: 1, tenantId: id(1), clientId: id(2), serviceId: "cowork",
  budgets: [
    { id: "legal", displayName: "Legal", businessArea: "Corporate services", costCenter: "LEGAL-01", owners: [id(3)],
      plannedUsd: 20000, tiers: { standard: { policyId: id(4), groupId: id(5) }, enhanced: { policyId: id(6), groupId: id(7) } } },
    { id: "finance", displayName: "Finance", businessArea: "Finance operations", costCenter: "FIN-02", owners: [id(3)],
      tiers: { standard: { policyId: id(8), groupId: id(9) }, enhanced: { policyId: id(10), groupId: id(11) } } },
  ],
});
const context: Context = { serviceId: "cowork", ownedBudgets: config.budgets,
  route: { state: "ready", budgetName: "Legal", message: "Routed to current owners." } };
function workspace(location: OwnerLocation, data: Context | null = context) {
  vi.stubGlobal("window", { location: { origin: "http://localhost:5173" } });
  return renderToStaticMarkup(<OwnerWorkspace location={location} context={data} onNavigate={() => undefined}
    api={new RequestApi("https://example.test", async () => "synthetic-token")} auth={new DemoAuth(config)}
    account={{ tenantId: id(1), localAccountId: id(3), homeAccountId: "home",
      environment: "login.microsoftonline.com", username: "owner@example.test" }}
    capabilities={null} onPublished={() => undefined} />);
}

describe("owner URL and page navigation", () => {
  it("defaults legacy and unknown pages to the business unit list", () => {
    for (const search of ["", "?view=owner", "?view=owner&ownerPage=unknown"]) {
      expect(ownerLocation(search)).toEqual({ page: "budgets" });
    }
    expect(ownerLocation("?view=owner&ownerPage=budgets&budget=legal")).toEqual({ page: "budgets", budgetId: "legal" });
    expect(ownerLocation("?view=owner&budget=legal")).toEqual({ page: "budgets", budgetId: "legal" });
  });
  it.each(["requests", "setup"] as const)("ignores stale budget parameters on the %s page", (page) => {
    expect(ownerLocation(`?view=owner&ownerPage=${page}&budget=legal`)).toEqual({ page });
  });
  it("marks only the selected owner page without changing the top-level identity/view", () => {
    const html = renderToStaticMarkup(<OwnerNavigation location={{ page: "requests" }} onNavigate={() => undefined} />);
    expect(html).toContain('aria-label="Policy owner pages"');
    expect(html).toContain('aria-pressed="true">Credit requests');
    expect(html).toContain('aria-pressed="false">Business units');
    expect(html).toContain('aria-pressed="false">Set up a BU');
    expect(html).not.toContain("Admin");
  });
});

describe("focused owner screens", () => {
  it("lands on BU metadata without mounting Graph details, requests or setup forms", () => {
    const html = workspace({ page: "budgets" });
    for (const text of ["Your business units", "Legal", "Finance", "LEGAL-01", "FIN-02", "$20,000.00", "Not configured",
      "View policies for Legal", "View policies for Finance", "Standard + Enhanced"]) expect(html).toContain(text);
    for (const text of ["Owner approval inbox", "Current readings available", 'name="costCenter"', "Refresh policies",
      "Approval moves the requester", id(4), id(8)]) expect(html).not.toContain(text);
  });
  it("shows only the opened authorized BU and a back action, without a budget dropdown", () => {
    const html = workspace({ page: "budgets", budgetId: "finance" });
    expect(html).toContain("Back to business units");
    expect(html).toContain("<h2>Finance</h2>");
    expect(html).toContain("Cost centre: FIN-02");
    expect(html).toContain("Standard and Enhanced policy settings");
    expect(html).not.toContain("LEGAL-01");
    expect(html).not.toContain("<select");
    expect(html).not.toContain("Owner approval inbox");
    expect(html).not.toContain('name="costCenter"');
  });
  it("does not substitute another BU for an unauthorized/deleted deep link", () => {
    const html = workspace({ page: "budgets", budgetId: "not-owned" });
    expect(html).toContain('role="alert"');
    expect(html).toContain("unavailable or you are no longer an owner");
    expect(html).toContain("Back to business units");
    expect(html).not.toContain("Loading Graph");
    expect(html).not.toContain("LEGAL-01");
  });
  it("shows requests and their approval warning without policy details or setup forms", () => {
    const html = workspace({ page: "requests" });
    expect(html).toContain("Owner approval inbox");
    expect(html).toContain("across all your business units");
    expect(html).toContain("manual return to Standard");
    expect(html).not.toContain("Refresh policies");
    expect(html).not.toContain("View policies");
    expect(html).not.toContain('name="costCenter"');
  });
  it("keeps setup available before metadata exists, on its own page only", () => {
    const html = workspace({ page: "setup" }, null);
    expect(html).toContain("Request a new business unit");
    expect(html).toContain("Your BU requests");
    expect(html).toContain('name="costCenter"');
    expect(html).not.toContain("Owner approval inbox");
    expect(html).not.toContain("View policies");
    expect(html).not.toContain("Refresh policies");
    expect(workspace({ page: "budgets" }, null)).not.toContain('name="costCenter"');
  });
  it("provides empty-state setup actions instead of requesting an unauthorized inbox", () => {
    const empty = { ...context, ownedBudgets: [] };
    expect(workspace({ page: "budgets" }, empty)).toContain("No business units yet");
    const requests = workspace({ page: "requests" }, empty);
    expect(requests).toContain("Set up your first BU");
    expect(requests).not.toContain("Owner approval inbox");
    expect(requests).not.toContain("Approval moves the requester");
  });
  it("does not turn a zero planning budget into an absent budget or actual spend", () => {
    const html = renderToStaticMarkup(<BusinessUnitList budgets={[{ ...config.budgets[0], plannedUsd: 0 }]}
      onNavigate={() => undefined} />);
    expect(html).toContain("$0.00");
    expect(html).not.toContain("Not configured");
    expect(html).not.toContain("spent");
  });
});
