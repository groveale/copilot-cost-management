import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { InteractionRequiredAuthError, type AccountInfo } from "@azure/msal-browser";
import { balanceScopes, DemoAuth, isGraphInteractionRequired, requesterScopes, scopes } from "./auth";
import { parseConfig } from "./config";

const msal = vi.hoisted(() => ({
  initialize: vi.fn(),
  getActiveAccount: vi.fn(),
  getAllAccounts: vi.fn(),
  loginPopup: vi.fn(),
  setActiveAccount: vi.fn(),
  acquireTokenSilent: vi.fn(),
  acquireTokenPopup: vi.fn(),
  logoutPopup: vi.fn(),
}));

vi.mock("@azure/msal-browser", async (importOriginal) => ({
  ...await importOriginal<typeof import("@azure/msal-browser")>(),
  PublicClientApplication: class {
    initialize = msal.initialize;
    getActiveAccount = msal.getActiveAccount;
    getAllAccounts = msal.getAllAccounts;
    loginPopup = msal.loginPopup;
    setActiveAccount = msal.setActiveAccount;
    acquireTokenSilent = msal.acquireTokenSilent;
    acquireTokenPopup = msal.acquireTokenPopup;
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
    expect(msal.acquireTokenPopup).not.toHaveBeenCalled();
  });
  it("reuses cached or silently renewed Graph access without a popup", async () => {
    msal.acquireTokenSilent.mockResolvedValue({ accessToken: "cached-or-renewed-graph-token" });
    await new DemoAuth(config).consentGraph(account);
    expect(msal.acquireTokenSilent).toHaveBeenCalledWith({ account, scopes: expected });
    expect(msal.acquireTokenPopup).not.toHaveBeenCalled();
  });
  it("never opens a popup automatically when silent Graph authentication requires interaction", async () => {
    const error = new InteractionRequiredAuthError("consent_required", "Graph consent is required.");
    msal.acquireTokenSilent.mockRejectedValue(error);
    await expect(new DemoAuth(config).token(account)).rejects.toBe(error);
    expect(isGraphInteractionRequired(error)).toBe(true);
    expect(msal.acquireTokenPopup).not.toHaveBeenCalled();
  });
  it("prompts only when explicitly requested and silent authentication actually requires it", async () => {
    msal.acquireTokenSilent.mockRejectedValue(new InteractionRequiredAuthError("consent_required", "Graph consent is required."));
    msal.acquireTokenPopup.mockResolvedValue({ account });
    await new DemoAuth(config).consentGraph(account);
    expect(msal.acquireTokenPopup).toHaveBeenCalledExactlyOnceWith({ account, scopes: expected });
  });
  it("does not turn other authentication failures into an interactive fallback", async () => {
    const error = new Error("Token service unavailable");
    msal.acquireTokenSilent.mockRejectedValue(error);
    await expect(new DemoAuth(config).consentGraph(account)).rejects.toBe(error);
    expect(isGraphInteractionRequired(error)).toBe(false);
    expect(msal.acquireTokenPopup).not.toHaveBeenCalled();
  });
  it("rejects a different account returned by interactive Graph consent", async () => {
    msal.acquireTokenSilent.mockRejectedValue(new InteractionRequiredAuthError("interaction_required", "Graph sign-in needs interaction."));
    msal.acquireTokenPopup.mockResolvedValue({ account: { ...account, homeAccountId: "different" } });
    await expect(new DemoAuth(config).consentGraph(account)).rejects.toThrow("account changed");
  });
  it("keeps the Function API token audience separate from Graph", async () => {
    const apiScope = "api://10000000-0000-0000-0000-000000000003/access_as_user";
    msal.acquireTokenSilent.mockResolvedValue({ accessToken: "api-token" });
    expect(await new DemoAuth({ ...config, apiScope }).apiToken(account)).toBe("api-token");
    expect(msal.acquireTokenSilent).toHaveBeenCalledExactlyOnceWith({ account, scopes: [apiScope] });
    expect(msal.acquireTokenPopup).not.toHaveBeenCalled();
  });
  it("uses only the existing user-data and policy read scopes for the personal balance", async () => {
    msal.acquireTokenSilent.mockResolvedValue({ accessToken: "balance-read-token" });
    expect(balanceScopes).toEqual([
      "https://graph.microsoft.com/CopilotCostManagement-UserData.Read.All",
      "https://graph.microsoft.com/CopilotCostManagement-Policy.Read.All",
    ]);
    expect(balanceScopes.every((scope) => scopes.includes(scope))).toBe(true);
    expect(await new DemoAuth(config).balanceToken(account)).toBe("balance-read-token");
    expect(msal.acquireTokenSilent).toHaveBeenCalledExactlyOnceWith({ account, scopes: balanceScopes });
    expect(msal.acquireTokenPopup).not.toHaveBeenCalled();
  });
  it("does not open a balance popup automatically, and explicit consent requests no operator write permissions", async () => {
    const error = new InteractionRequiredAuthError("consent_required", "Balance consent is required.");
    msal.acquireTokenSilent.mockRejectedValue(error);
    msal.acquireTokenPopup.mockResolvedValue({ account });
    const auth = new DemoAuth(config);
    await expect(auth.balanceToken(account)).rejects.toBe(error);
    expect(msal.acquireTokenPopup).not.toHaveBeenCalled();
    await auth.consentBalance(account);
    expect(msal.acquireTokenPopup).toHaveBeenCalledExactlyOnceWith({ account, scopes: balanceScopes });
  });
  it("reuses silent balance access and surfaces non-interaction failures without prompting", async () => {
    const auth = new DemoAuth(config);
    msal.acquireTokenSilent.mockResolvedValue({ accessToken: "cached" });
    await auth.consentBalance(account);
    expect(msal.acquireTokenPopup).not.toHaveBeenCalled();
    msal.acquireTokenSilent.mockRejectedValue(new Error("Token service unavailable"));
    await expect(auth.consentBalance(account)).rejects.toThrow("Token service unavailable");
    expect(msal.acquireTokenPopup).not.toHaveBeenCalled();
  });
  it("rejects stale and changed accounts for personal balance authentication", async () => {
    const auth = new DemoAuth(config);
    msal.getActiveAccount.mockReturnValue({ ...account, homeAccountId: "other" });
    await expect(auth.balanceToken(account)).rejects.toThrow("account changed");
    expect(msal.acquireTokenSilent).not.toHaveBeenCalled();
    msal.getActiveAccount.mockReturnValue(account);
    msal.acquireTokenSilent.mockRejectedValue(new InteractionRequiredAuthError("consent_required", "Consent required."));
    msal.acquireTokenPopup.mockResolvedValue({ account: { ...account, localAccountId: "other" } });
    await expect(auth.consentBalance(account)).rejects.toThrow("account changed");
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
