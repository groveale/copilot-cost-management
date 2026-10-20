import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { DemoAuth, requesterScopes, scopes } from "./auth";
import { parseConfig } from "./config";

const msal = vi.hoisted(() => ({
  initialize: vi.fn(),
  getActiveAccount: vi.fn(),
  getAllAccounts: vi.fn(),
  loginPopup: vi.fn(),
  setActiveAccount: vi.fn(),
  acquireTokenSilent: vi.fn(),
  logoutPopup: vi.fn(),
}));

vi.mock("@azure/msal-browser", () => ({
  PublicClientApplication: class {
    initialize = msal.initialize;
    getActiveAccount = msal.getActiveAccount;
    getAllAccounts = msal.getAllAccounts;
    loginPopup = msal.loginPopup;
    setActiveAccount = msal.setActiveAccount;
    acquireTokenSilent = msal.acquireTokenSilent;
    logoutPopup = msal.logoutPopup;
  },
}));

const config = parseConfig({
  version: 1,
  tenantId: "10000000-0000-0000-0000-000000000001",
  clientId: "10000000-0000-0000-0000-000000000002",
  serviceId: "cowork",
  budgets: [],
});
const account: AccountInfo = {
  homeAccountId: "test-home-account",
  localAccountId: "test-local-account",
  environment: "login.microsoftonline.com",
  tenantId: config.tenantId,
  username: "operator@example.test",
};
const expected = [
  "CopilotCostManagement-Assignment.Read.All",
  "CopilotCostManagement-Assignment.ReadWrite.All",
  "CopilotCostManagement-Policy.Read.All",
  "CopilotCostManagement-Policy.ReadWrite.All",
  "CopilotCostManagement-UserData.Read.All",
  "Directory.AccessAsUser.All",
  "Directory.Read.All",
  "Directory.ReadWrite.All",
].map((scope) => `https://graph.microsoft.com/${scope}`);

beforeEach(() => {
  vi.resetAllMocks();
  vi.stubGlobal("window", { location: { origin: "http://localhost:5173" } });
  msal.getActiveAccount.mockReturnValue(account);
  msal.getAllAccounts.mockReturnValue([account]);
});
afterEach(() => vi.unstubAllGlobals());

describe("operator-selected delegated scopes", () => {
  it("contains exactly the eight requested Microsoft Graph scopes", () => {
    expect(scopes).toEqual(expected);
    expect(new Set(scopes).size).toBe(8);
  });
  it("requests the selected scopes during interactive sign-in", async () => {
    msal.loginPopup.mockResolvedValue({ account });
    const auth = new DemoAuth(config);
    expect(await auth.signIn()).toBe(account);
    expect(msal.loginPopup).toHaveBeenCalledWith({ scopes: expected, prompt: "select_account" });
    expect(msal.setActiveAccount).toHaveBeenCalledWith(account);
  });
  it("uses only identity scopes for the requester page and never acquires a Graph token", async () => {
    msal.loginPopup.mockResolvedValue({ account });
    expect(await new DemoAuth(config).signIn("requester")).toBe(account);
    expect(requesterScopes).toEqual(["openid", "profile"]);
    expect(msal.loginPopup).toHaveBeenCalledWith({ scopes: ["openid", "profile"], prompt: "select_account" });
    expect(msal.acquireTokenSilent).not.toHaveBeenCalled();
  });
  it("requests the same scopes during silent token acquisition", async () => {
    msal.acquireTokenSilent.mockResolvedValue({ accessToken: "test-token" });
    expect(await new DemoAuth(config).token()).toBe("test-token");
    expect(msal.acquireTokenSilent).toHaveBeenCalledWith({ account, scopes: expected });
  });
  it("does not issue a token for a previous workspace after the account changes", async () => {
    msal.getActiveAccount.mockReturnValue({ ...account, homeAccountId: "different-account" });
    await expect(new DemoAuth(config).token(account)).rejects.toThrow("account changed");
    expect(msal.acquireTokenSilent).not.toHaveBeenCalled();
  });
  it("does not issue a token for the wrong tenant or a signed-out workspace", async () => {
    await expect(new DemoAuth(config).token({ ...account, tenantId: config.clientId })).rejects.toThrow("account changed");
    await expect(new DemoAuth(config).token(null)).rejects.toThrow("Sign in");
    expect(msal.acquireTokenSilent).not.toHaveBeenCalled();
  });
});
