import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, describe, expect, it, vi } from "vitest";
import { RequestApi, type Context } from "./api";
import { DemoAuth } from "./auth";
import { parseConfig } from "./config";
import { BusinessUnitList, OwnerNavigation, OwnerWorkspace, ownerLocation, type OwnerLocation } from "./HostedPortal";
import { BudgetAllocation } from "./Onboarding";
import { AddPolicyMember } from "./PolicyMembers";

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
  it.each(["standard", "enhanced"] as const)("renders explicit %s editors without running Graph or API actions on mount", (tier) => {
    vi.stubGlobal("window", { location: { origin: "http://localhost:5173" } });
    const api = new RequestApi("https://example.test", async () => "synthetic-token");
    const auth = new DemoAuth(config);
    const token = vi.spyOn(auth, "token");
    const allocation = vi.spyOn(api, "allocation");
    const member = vi.spyOn(api, "member");
    const onBusyChange = vi.fn();
    const account = { tenantId: id(1), localAccountId: id(3), homeAccountId: "home",
      environment: "login.microsoftonline.com", username: "owner@example.test" };
    const budget = { ...config.budgets[0], approvedCreditCeiling: 2_000_000 };
    const limits = renderToStaticMarkup(<BudgetAllocation budget={budget} tier={tier} api={api} auth={auth} account={account}
      onApplied={() => undefined} onBusyChange={onBusyChange} />);
    expect(limits).toContain('aria-expanded="false"');
    expect(limits).toContain('hidden=""');
    expect(limits).toContain("Monthly allocation (credits)");
    expect(limits).toContain("Per-user credit limit");
    expect(limits).toContain("Save policy limits");
    expect(limits).not.toContain('type="range"');
    expect(limits).toContain("BU monthly ceiling: 2,000,000 Copilot Credits");
    expect(limits).not.toContain("Demo planning budget:");
    const members = renderToStaticMarkup(<AddPolicyMember budget={budget} tier={tier} serviceId="cowork" api={api}
      auth={auth} account={account} disabled={false} onAdded={() => undefined} onBusyChange={onBusyChange} />);
    expect(members).toContain('aria-expanded="false"');
    expect(members).toContain('hidden=""');
    expect(members).toContain("Sign-in email (UPN)");
    expect(members).toContain("Find user");
    expect(members).toContain("already in the other tier cannot be added");
    expect(token).not.toHaveBeenCalled(); expect(allocation).not.toHaveBeenCalled();
    expect(member).not.toHaveBeenCalled(); expect(onBusyChange).not.toHaveBeenCalled();
  });
  it.each(["standard", "enhanced"] as const)("enables the %s editor with the shared GBP 5K demo ceiling when no approval is recorded", (tier) => {
    vi.stubGlobal("window", { location: { origin: "http://localhost:5173" } });
    const api = new RequestApi("https://example.test", async () => "synthetic-token");
    const allocation = vi.spyOn(api, "allocation");
    const account = { tenantId: id(1), localAccountId: id(3), homeAccountId: "home",
      environment: "login.microsoftonline.com", username: "owner@example.test" };
    const html = renderToStaticMarkup(<BudgetAllocation budget={config.budgets[0]} tier={tier} api={api}
      auth={new DemoAuth(config)} account={account} onApplied={() => undefined} onBusyChange={() => undefined} />);
    const editButton = html.match(/<button[^>]*aria-label="Edit (?:Standard|Enhanced) policy limits"[^>]*>/)?.[0];
    expect(editButton).toBeDefined();
    expect(editButton).not.toContain("disabled");
    expect(html).toContain("\u00a35,000/month = 500,000 credits shared by Standard and Enhanced");
    expect(html).toContain("assuming \u00a30.01 per credit");
    expect(html).toContain("Not financial approval or a currency conversion");
    expect(html).toContain("BU monthly ceiling: 500,000 Copilot Credits");
    expect(html).not.toContain("approved BU credit ceiling is required");
    expect(allocation).not.toHaveBeenCalled();
  });
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
    expect(html).toContain("Approval moves the requester");
    expect(html).not.toContain("manual return to Standard");
    expect(html).not.toContain("Refresh policies");
    expect(html).not.toContain("View policies");
    expect(html).not.toContain('name="costCenter"');
  });
  it("keeps setup available before metadata exists, on its own page only", () => {
    const html = workspace({ page: "setup" }, null);
    expect(html).toContain("Request a new business unit");
    expect(html).toContain("Your BU requests");
    expect(html).toContain('name="costCenter"');
    expect(html).toContain("bu-request-form");
    expect(html).toContain("bu-form-grid");
    expect(html).toContain("Proposed monthly credit allocation");
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
