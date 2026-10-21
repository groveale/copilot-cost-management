import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, expect, it, vi } from "vitest";
import { RequestApi } from "./api";
import { HostedPortal } from "./HostedPortal";
import { DemoAuth } from "./auth";
import { parseConfig } from "./config";
import { BusinessUnitOnboarding } from "./Onboarding";

vi.mock("@azure/msal-browser", () => ({
  PublicClientApplication: class { getActiveAccount() { return null; } getAllAccounts() { return []; } },
}));
afterEach(() => vi.unstubAllGlobals());
it("shows only login before authentication, never metadata or request fields", () => {
  vi.stubGlobal("window", { location: { origin: "http://localhost:5173" } });
  const config = parseConfig({
    version: 1, tenantId: "10000000-0000-4000-8000-000000000001",
    clientId: "10000000-0000-4000-8000-000000000002", serviceId: "cowork", budgets: [],
  });
  const html = renderToStaticMarkup(<HostedPortal config={config} auth={new DemoAuth(config)} requestPage />);
  expect(html).toContain("Sign in with Microsoft");
  expect(html).not.toContain("textarea"); expect(html).not.toContain("Owner approval inbox");
  expect(html).not.toContain("Request a new BU");
});
it("renders the ordinary-user BU form with the exact 20K calculation and no administrator controls", () => {
  vi.stubGlobal("window", { location: { origin: "http://localhost:5173" } });
  const config = parseConfig({ version: 1, tenantId: "10000000-0000-4000-8000-000000000001",
    clientId: "10000000-0000-4000-8000-000000000002", serviceId: "cowork", budgets: [] });
  const account = { tenantId: config.tenantId, homeAccountId: "home", localAccountId: "local",
    environment: "login.microsoftonline.com", username: "owner@example.test", name: "Example owner" };
  const html = renderToStaticMarkup(<BusinessUnitOnboarding api={new RequestApi("https://example.test", async () => "token")}
    auth={new DemoAuth(config)} account={account} onPublished={() => undefined} />);
  expect(html).toContain("Request a new business unit");
  expect(html).toContain("2,000,000 credits/month: 1,000,000 Standard + 1,000,000 Enhanced");
  expect(html).toContain("owner@example.test");
  expect(html).not.toContain("Onboarding administration");
  expect(html).not.toContain('name="owners"');
});
it("uses only the API bearer token, includes an abort signal, and transmits the stable operation ID", async () => {
  const transport = vi.fn().mockResolvedValue(new Response(JSON.stringify({ id: "saved" })));
  const api = new RequestApi("https://example.test", async () => "api-token", transport);
  const signal = new AbortController().signal;
  const input = { operationId: "operation", justification: "Our project needs additional capacity.", grantKind: "permanent" as const };
  await api.submit(input, signal);
  expect(transport).toHaveBeenCalledWith("https://example.test/api/v1/requests", {
    method: "POST", signal, headers: { Authorization: "Bearer api-token", "Content-Type": "application/json" },
    body: JSON.stringify(input),
  });
});
it("surfaces failures and correlation IDs without successful fallback data", async () => {
  const api = new RequestApi("https://example.test", async () => "token",
    vi.fn().mockResolvedValue(new Response(JSON.stringify({ message: "Access denied", correlationId: "reference" }), { status: 403 })));
  await expect(api.history(new AbortController().signal)).rejects.toThrow("Access denied Reference: reference");
});
it("accepts explicit JSON null for uninitialized setup and absent allocation operations", async () => {
  const transport = vi.fn(async () => new Response("null", { headers: { "Content-Type": "application/json" } }));
  const api = new RequestApi("https://example.test", async () => "token", transport);
  const signal = new AbortController().signal;
  expect(await api.onboardingSetup(signal)).toBeNull();
  expect(await api.allocation(`bu-${"a".repeat(24)}`, signal)).toBeNull();
});
it("keeps pagination cursors inside the query value", async () => {
  const transport = vi.fn().mockResolvedValue(new Response('{"items":[]}'));
  const api = new RequestApi("https://example.test", async () => "token", transport);
  await api.inbox(new AbortController().signal, "next&requesterId=forged");
  expect(transport.mock.calls[0][0]).toBe("https://example.test/api/v1/inbox?cursor=next%26requesterId%3Dforged");
});
