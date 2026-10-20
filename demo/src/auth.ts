import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";
import type { DemoConfig } from "./config";

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

export class DemoAuth {
  private readonly app: PublicClientApplication;

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
    const result = await this.app.loginPopup({
      scopes: mode === "requester" ? requesterScopes : scopes,
      prompt: "select_account",
    });
    if (result.account.tenantId.toLowerCase() !== this.config.tenantId) {
      throw new Error("Sign-in returned an account from a different tenant.");
    }
    this.app.setActiveAccount(result.account);
    return result.account;
  }

  async token(account = this.account()): Promise<string> {
    const active = this.account();
    if (!account || !active) throw new Error("Sign in to the configured tenant before loading Graph details.");
    if (account.tenantId.toLowerCase() !== this.config.tenantId ||
        account.homeAccountId !== active.homeAccountId || account.localAccountId !== active.localAccountId) {
      throw new Error("The signed-in account changed. Sign in again before loading Graph details.");
    }
    const result = await this.app.acquireTokenSilent({ account, scopes });
    return result.accessToken;
  }

  async signOut(): Promise<void> {
    const account = this.account();
    if (account) await this.app.logoutPopup({ account, mainWindowRedirectUri: window.location.href });
  }
}
