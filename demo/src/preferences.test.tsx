import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { readPreferences, applyAppearance } from "./preferences";
import { identityDiagnostics, Settings } from "./Settings";
import type { DemoAuth } from "./auth";
import type { DemoConfig } from "./config";

afterEach(() => vi.unstubAllGlobals());
describe("browser settings and safe diagnostics", () => {
  it("defaults to Real time and System, with independent persisted choices", () => {
    expect(readPreferences({ getItem: () => null })).toEqual({ value: { source: "realtime", appearance: "system" }, error: "" });
    expect(readPreferences({ getItem: (key) => key === "copilotReportingSource" ? "sync" : "light" }).value)
      .toEqual({ source: "sync", appearance: "light" });
  });
  it("reports invalid or inaccessible preferences explicitly", () => {
    expect(readPreferences({ getItem: () => "unknown" }).error).toContain("invalid");
    expect(readPreferences({ getItem: () => { throw new Error("Storage blocked"); } }).error).toContain("Storage blocked");
  });
  it("honors the explicit theme override before saved/system choices", () => {
    const setAttribute = vi.fn();
    vi.stubGlobal("document", { documentElement: { setAttribute } });
    vi.stubGlobal("window", { location: { search: "?scoutTheme=light" }, matchMedia: () => ({ matches: true }) });
    applyAppearance("dark");
    expect(setAttribute).toHaveBeenLastCalledWith("data-theme", "light");
    window.location.search = "";
    applyAppearance("system");
    expect(setAttribute).toHaveBeenLastCalledWith("data-theme", "dark");
    applyAppearance("light");
    expect(setAttribute).toHaveBeenLastCalledWith("data-theme", "light");
  });
  it("allowlists decoded identity claims without exposing token strings or arbitrary claims", () => {
    const account: AccountInfo = { homeAccountId: "home", localAccountId: "object", username: "admin@example.test",
      tenantId: "tenant", environment: "login.microsoftonline.com", idToken: "secret-id-token",
      idTokenClaims: { tid: "tenant", oid: "object", amr: ["pwd", "mfa"], exp: 1900000000, secret: "unapproved-sensitive-claim", access_token: "unapproved-sensitive-claim" } };
    expect(identityDiagnostics(account)).toEqual({ tid: "tenant", oid: "object", exp: 1900000000, amr: ["pwd", "mfa"] });
    const diagnostics = vi.fn<DemoAuth["diagnostics"]>(() => []);
    const config: DemoConfig = { version: 1, clientId: "app", tenantId: "tenant", serviceId: "cowork", budgets: [] };
    const html = renderToStaticMarkup(<Settings account={account} auth={{ diagnostics }} config={config}
      preferences={{ value: { source: "sync", appearance: "system" }, error: "", save: vi.fn() }} locked onAuthenticate={vi.fn()} />);
    expect(html).toContain('aria-label="Settings"');
    expect(html).toContain("Sign-in details");
    expect(html).toContain("not proof of authorization");
    expect(html).toContain("disabled");
    expect(html).not.toContain("secret-id-token");
    expect(html).not.toContain("unapproved-sensitive-claim");
    expect(diagnostics).not.toHaveBeenCalled();
  });
});
