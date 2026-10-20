import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { DemoAuth } from "./auth";
import { BudgetPortal } from "./BudgetPortal";
import { budgetsForOwner, parseConfig } from "./config";

const msal = vi.hoisted(() => ({
  getActiveAccount: vi.fn(),
  getAllAccounts: vi.fn(),
}));

vi.mock("@azure/msal-browser", () => ({
  PublicClientApplication: class {
    getActiveAccount = msal.getActiveAccount;
    getAllAccounts = msal.getAllAccounts;
  },
}));

const id = (n: number) => `10000000-0000-0000-0000-${String(n).padStart(12, "0")}`;
const config = parseConfig({
  version: 1, tenantId: id(1), clientId: id(2), serviceId: "cowork",
  budgets: [
    {
      id: "legal", displayName: "Legal", businessArea: "Legal", costCenter: "legal-cost-centre",
      owners: ["admin@m365cpi77517573.onmicrosoft.com", "coowner@example.test"],
      tiers: {
        standard: { policyId: id(3), groupId: id(4) },
        enhanced: { policyId: id(5), groupId: id(6) },
      },
    },
    {
      id: "finance", displayName: "Finance", businessArea: "Finance", costCenter: "finance-cost-centre",
      owners: ["finance@example.test"],
      tiers: {
        standard: { policyId: id(7), groupId: id(8) },
        enhanced: { policyId: id(9), groupId: id(10) },
      },
    },
  ],
});
const account: AccountInfo = {
  homeAccountId: "test-home",
  localAccountId: "test-local",
  tenantId: config.tenantId,
  environment: "login.microsoftonline.com",
  username: "admin@M365CPI77517573.onmicrosoft.com",
};

beforeEach(() => {
  vi.resetAllMocks();
  vi.stubGlobal("window", { location: { origin: "http://localhost:5173" } });
  msal.getActiveAccount.mockReturnValue(null);
  msal.getAllAccounts.mockReturnValue([]);
});
afterEach(() => vi.unstubAllGlobals());

function render(current: AccountInfo | null) {
  msal.getActiveAccount.mockReturnValue(current);
  return renderToStaticMarkup(<BudgetPortal config={config} auth={new DemoAuth(config)} />);
}

describe("JSON owner lookup", () => {
  it("does not return any budgets before sign-in", () => {
    expect(budgetsForOwner(config, null)).toEqual([]);
  });
  it("matches the real Legal owner case-insensitively", () => {
    expect(budgetsForOwner(config, account).map((budget) => budget.id)).toEqual(["legal"]);
  });
  it("supports a co-owner and whitespace without using substring or domain matches", () => {
    expect(budgetsForOwner(config, { ...account, username: "  COOWNER@example.test  " }).map((budget) => budget.id)).toEqual(["legal"]);
    for (const username of ["", " ", "admin", "@m365cpi77517573.onmicrosoft.com", "admin@m365cpi77517573.onmicrosoft.com.attacker.test"]) {
      expect(budgetsForOwner(config, { ...account, username })).toEqual([]);
    }
  });
  it("rejects the same username from a different tenant", () => {
    expect(budgetsForOwner(config, { ...account, tenantId: id(11) })).toEqual([]);
  });
  it("returns all and only the assigned budgets", () => {
    const multiple = {
      ...config,
      budgets: config.budgets.map((budget) => ({
        ...budget, owners: [...budget.owners, "multiple@example.test"],
      })),
    };
    expect(budgetsForOwner(multiple, { ...account, username: "multiple@example.test" }).map((budget) => budget.id)).toEqual(["legal", "finance"]);
    expect(budgetsForOwner(config, { ...account, username: "unknown@example.test" })).toEqual([]);
    expect(budgetsForOwner({ ...config, budgets: [] }, account)).toEqual([]);
    expect(budgetsForOwner({ ...config, budgets: config.budgets.map((budget) => ({ ...budget, owners: [] })) }, account)).toEqual([]);
  });
});

describe("owner workspace rendering", () => {
  it("renders only the login surface before authentication", () => {
    const html = render(null);
    expect(html).toContain("Sign in with Microsoft");
    expect(html).not.toContain("Legal");
    expect(html).not.toContain("Finance");
    expect(html).not.toContain("admin@m365cpi77517573.onmicrosoft.com");
    expect(html).not.toContain(config.tenantId);
    expect(html).not.toContain(config.clientId);
    expect(html).not.toContain(config.budgets[0].tiers.standard.policyId);
    expect(html).not.toContain("Refresh Graph details");
    expect(html).not.toContain("Connection and static configuration");
    expect(html).not.toContain("Requested delegated permissions");
  });
  it("renders a restored owner's assigned budget without other budgets' metadata", () => {
    const html = render(account);
    expect(html).toContain("Legal");
    expect(html).toContain("legal-cost-centre");
    expect(html).toContain("Refresh Graph details");
    expect(html).not.toContain("Finance");
    expect(html).not.toContain("finance-cost-centre");
    expect(html).not.toContain("finance@example.test");
  });
  it("shows an explicit no-assignment state with no fallback to the first configured budget", () => {
    const html = render({ ...account, username: "unknown@example.test" });
    expect(html).toContain("No budgets assigned");
    expect(html).not.toContain("Legal");
    expect(html).not.toContain("Finance");
    expect(html).not.toContain("Refresh Graph details");
    expect(html).not.toContain(config.budgets[0].tiers.standard.policyId);
  });
});
