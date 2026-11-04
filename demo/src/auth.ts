import { InteractionRequiredAuthError, PublicClientApplication, type AccountInfo, type AuthenticationResult } from "@azure/msal-browser";
import { object, type DemoConfig } from "./config";

export const scopes = [
  "CopilotCostManagement-Assignment.Read.All",
  "CopilotCostManagement-Assignment.ReadWrite.All",
  "CopilotCostManagement-Policy.Read.All",
  "CopilotCostManagement-Policy.ReadWrite.All",
  "CopilotCostManagement-UserData.Read.All",
  "Directory.AccessAsUser.All",
  "Directory.Read.All",
  "Directory.ReadWrite.All",
].map((scope) => `https://graph.microsoft.com/${scope}`);

export const requesterScopes = ["openid", "profile"];
export const reportingScopes = ["https://graph.microsoft.com/Sites.Selected"];
export interface TokenDiagnostic { scopes: string[]; acquiredAt: string; expiresAt: string | null; fromCache: boolean }
export const balanceScopes = [
  "https://graph.microsoft.com/CopilotCostManagement-UserData.Read.All",
  "https://graph.microsoft.com/CopilotCostManagement-Policy.Read.All",
];
export const syncScopes = [
  "CopilotCostManagement-Policy.Read.All", "CopilotCostManagement-Assignment.Read.All",
  "CopilotCostManagement-UserData.Read.All", "GroupMember.Read.All", "User.ReadBasic.All", "Sites.Selected",
].map((scope) => `https://graph.microsoft.com/${scope}`);
export const isGraphInteractionRequired = (error: unknown): error is InteractionRequiredAuthError =>
  error instanceof InteractionRequiredAuthError;
function syncClaims(claims?: string) {
  const request = claims ? object(JSON.parse(claims), "Claims request") : {};
  const access = request.access_token === undefined ? {} : object(request.access_token, "Access-token claims");
  return JSON.stringify({ ...request, access_token: { ...access, xms_cc: { values: ["cp1"] } } });
}

export class DemoAuth {
  private readonly app: PublicClientApplication;
  private syncChallenge?: { accountId: string; claims: string };
  private apiChallenge?: { accountId: string; claims: string };
  private reportingChallenge?: { accountId: string; claims: string };
  private acquisitions = new Map<string, TokenDiagnostic>();
  private diagnosticAccount?: string;

  constructor(private readonly config: DemoConfig) {
    this.app = new PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
        redirectUri: new URL("auth.html", new URL(import.meta.env.BASE_URL, window.location.origin)).href,
        postLogoutRedirectUri: new URL(import.meta.env.BASE_URL, window.location.origin).href,
      },
      cache: { cacheLocation: "sessionStorage" },
    });
  }

  async initialize(): Promise<void> {
    await this.app.initialize();
  }

  account(): AccountInfo | null {
    const active = this.app.getActiveAccount();
    if (active?.tenantId.toLowerCase() === this.config.tenantId) return active;
    const accounts = this.app.getAllAccounts().filter((account) =>
      account.tenantId.toLowerCase() === this.config.tenantId);
    return accounts.length === 1 ? accounts[0] : null;
  }

  async signIn(mode: "owner" | "requester" = "owner"): Promise<AccountInfo> {
    this.clearDiagnostics();
    const result = await this.app.loginPopup({
      scopes: mode === "requester" ? [...requesterScopes, ...(this.config.apiScope ? [this.config.apiScope] : [])] : scopes,
      prompt: "select_account",
    });
    if (result.account.tenantId.toLowerCase() !== this.config.tenantId) {
      throw new Error("Sign-in returned an account from a different tenant.");
    }
    this.app.setActiveAccount(result.account);
    this.record(result, result.scopes ?? requesterScopes, result.account);
    return result.account;
  }

  async token(account = this.account()): Promise<string> {
    return this.acquire(scopes, account);
  }

  async apiToken(account = this.account()): Promise<string> {
    if (!this.config.apiScope) throw new Error("The API permission is not configured.");
    const claims = account && this.apiChallenge?.accountId === account.homeAccountId ? this.apiChallenge.claims : undefined;
    return this.acquire([this.config.apiScope], account, claims ? { claims, forceRefresh: false } : {});
  }

  async refreshApi(account: AccountInfo, claims?: string): Promise<void> {
    if (!this.config.apiScope) throw new Error("The API permission is not configured.");
    const requested = claims ?? (this.apiChallenge?.accountId === account.homeAccountId ? this.apiChallenge.claims : undefined);
    const request = syncClaims(requested);
    await this.refreshResource([this.config.apiScope], account, request);
    this.apiChallenge = { accountId: account.homeAccountId, claims: request };
  }
  private async refreshResource(requestScopes: string[], account: AccountInfo, claims: string): Promise<void> {
    try { await this.acquire(requestScopes, account, { forceRefresh: true, claims }); }
    catch (error) {
      if (!isGraphInteractionRequired(error)) throw error;
      const result = await this.app.acquireTokenPopup({ account, scopes: requestScopes, claims });
      this.checkResult(result, account);
      this.record(result, requestScopes, account);
    }
  }

  async reportingToken(account: AccountInfo): Promise<string> {
    const claims = this.reportingChallenge?.accountId === account.homeAccountId ? this.reportingChallenge.claims : undefined;
    return this.acquire(reportingScopes, account, claims ? { claims } : {});
  }
  async refreshReporting(account: AccountInfo, claims?: string): Promise<void> {
    const requested = syncClaims(claims ?? (this.reportingChallenge?.accountId === account.homeAccountId ? this.reportingChallenge.claims : undefined));
    await this.refreshResource(reportingScopes, account, requested);
    this.reportingChallenge = { accountId: account.homeAccountId, claims: requested };
  }
  async consentReporting(account: AccountInfo): Promise<void> { return this.consent(reportingScopes, account); }

  diagnostics(account: AccountInfo): TokenDiagnostic[] {
    if (this.diagnosticAccount !== `${account.homeAccountId}:${account.localAccountId}`) return [];
    return [...this.acquisitions.values()].map((item) => ({ ...item, scopes: [...item.scopes] }));
  }
  private clearDiagnostics() {
    this.acquisitions.clear();
    this.diagnosticAccount = undefined;
    this.apiChallenge = undefined;
    this.reportingChallenge = undefined;
    this.syncChallenge = undefined;
  }
  private record(result: AuthenticationResult, requestScopes: string[], account: AccountInfo) {
    const id = `${account.homeAccountId}:${account.localAccountId}`;
    if (this.diagnosticAccount !== id) this.acquisitions.clear();
    this.diagnosticAccount = id;
    const acquiredScopes = [...(result.scopes ?? requestScopes)].sort();
    this.acquisitions.set(acquiredScopes.join(" "), { scopes: acquiredScopes, acquiredAt: new Date().toISOString(),
      expiresAt: result.expiresOn?.toISOString() ?? null, fromCache: result.fromCache === true });
  }
  private checkResult(result: AuthenticationResult, account: AccountInfo) {
    if (result.account.homeAccountId !== account.homeAccountId || result.account.localAccountId !== account.localAccountId ||
        result.account.tenantId.toLowerCase() !== this.config.tenantId) throw new Error("The account changed during authentication. Sign in again.");
  }

  async balanceToken(account = this.account()): Promise<string> {
    return this.acquire(balanceScopes, account);
  }

  async syncToken(account: AccountInfo, forceRefresh = false, claims?: string): Promise<string> {
    const accountId = `${account.homeAccountId}:${account.localAccountId}`;
    const requested = claims ?? (this.syncChallenge?.accountId === accountId ? this.syncChallenge.claims : undefined);
    const token = await this.acquire(syncScopes, account, { forceRefresh, claims: syncClaims(requested) });
    if (requested) this.syncChallenge = { accountId, claims: requested };
    return token;
  }

  async refreshSync(account: AccountInfo, claims?: string): Promise<void> {
    try { await this.syncToken(account, true, claims); return; }
    catch (error) { if (!isGraphInteractionRequired(error)) throw error; }
    const accountId = `${account.homeAccountId}:${account.localAccountId}`;
    const requested = claims ?? (this.syncChallenge?.accountId === accountId ? this.syncChallenge.claims : undefined);
    const result = await this.app.acquireTokenPopup({ account, scopes: syncScopes, claims: syncClaims(requested) });
    if (result.account.homeAccountId !== account.homeAccountId || result.account.localAccountId !== account.localAccountId ||
        result.account.tenantId.toLowerCase() !== this.config.tenantId) throw new Error("The account changed during sync authentication.");
    if (requested) this.syncChallenge = { accountId, claims: requested };
    this.record(result, syncScopes, account);
  }

  async consentGraph(account: AccountInfo): Promise<void> {
    return this.consent(scopes, account);
  }

  async consentBalance(account: AccountInfo): Promise<void> {
    return this.consent(balanceScopes, account);
  }

  private async consent(requestScopes: string[], account: AccountInfo): Promise<void> {
    try {
      await this.acquire(requestScopes, account);
      return;
    } catch (error) {
      if (!isGraphInteractionRequired(error)) throw error;
    }
    const result = await this.app.acquireTokenPopup({ account, scopes: requestScopes });
    if (result.account.homeAccountId !== account.homeAccountId || result.account.localAccountId !== account.localAccountId) {
      throw new Error("The account changed during Graph consent. Sign in again.");
    }
    this.checkResult(result, account);
    this.record(result, requestScopes, account);
  }

  private async acquire(requestScopes: string[], account: AccountInfo | null,
    options: { forceRefresh?: boolean; claims?: string } = {}): Promise<string> {
    const active = this.account();
    if (!account || !active) throw new Error("Sign in to the configured tenant before loading protected data.");
    if (account.tenantId.toLowerCase() !== this.config.tenantId ||
        account.homeAccountId !== active.homeAccountId || account.localAccountId !== active.localAccountId) {
      throw new Error("The signed-in account changed. Sign in again before loading protected data.");
    }
    const result = await this.app.acquireTokenSilent({ account, scopes: requestScopes, ...options });
    if (result.account) this.checkResult(result, account);
    this.record(result, requestScopes, account);
    return result.accessToken;
  }

  async signOut(): Promise<void> {
    this.clearDiagnostics();
    const account = this.account();
    if (account) await this.app.logoutPopup({ account, mainWindowRedirectUri: window.location.href });
  }
}
