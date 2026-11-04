import { HttpRequest, InvocationContext } from "@azure/functions";
import { afterEach, beforeAll, expect, it, vi } from "vitest";
import { handler } from "../src/functions.js";
import { OnboardingService } from "../src/onboarding-service.js";
import { ReportingService } from "../src/reporting.js";
import { ReportingInteractionRequired } from "../src/reporting-auth.js";
import { ApiError } from "../src/model.js";
const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
vi.mock("../src/auth.js", () => ({ createAuthenticator: () => async () => ({
  tenantId: "10000000-0000-4000-8000-000000000001", objectId: "10000000-0000-4000-8000-000000000004",
}) }));
beforeAll(() => {
  process.env.TENANT_ID = id(1); process.env.API_AUDIENCE = id(2); process.env.SPA_CLIENT_ID = id(3);
  process.env.BUSINESS_TABLE_URL = "https://example.table.core.windows.net";
  process.env.BUSINESS_BLOB_URL = "https://example.blob.core.windows.net";
});
afterEach(() => vi.restoreAllMocks());
async function request(path: string, input?: unknown) {
  return handler(new HttpRequest({ url: `https://example.test/api/v1/${path}`, method: input === undefined ? "GET" : "POST", params: { path },
    headers: input === undefined ? {} : { "content-type": "application/json" },
    ...(input === undefined ? {} : { body: { string: JSON.stringify(input) } }) }),
  new InvocationContext({ functionName: "test", logHandler: vi.fn() }));
}
it("exposes only the typed settings contract to authenticated authorized owners", async () => {
  const settings = { enabled: true, accessMode: "sharepoint-acl" as const, siteId: "site", policyListId: id(8), balanceListId: id(9), isAdmin: false, policyIds: [id(6)], securityReady: true };
  vi.spyOn(ReportingService.prototype, "settings").mockResolvedValue(settings);
  expect(await request("reporting/settings")).toMatchObject({ status: 200, jsonBody: settings, headers: { "Cache-Control": "no-store" } });
});
it("rejects non-admins and arbitrary ACL targets before reconciliation", async () => {
  const reconcile = vi.spyOn(ReportingService.prototype, "reconcile");
  vi.spyOn(OnboardingService.prototype, "requireAdmin").mockImplementation(() => { throw new ApiError(403, "admin", "Admin required."); });
  expect((await request("admin/reporting/security", {})).status).toBe(403);
  vi.spyOn(OnboardingService.prototype, "requireAdmin").mockImplementation(() => {});
  for (const input of [{ owners: [id(5)] }, { url: "https://bad.example" }, { jobId: id(5) }, { policyId: "bad" }]) {
    expect((await request("admin/reporting/security", input)).status).toBe(400);
  }
  expect(reconcile).not.toHaveBeenCalled();
});
it("returns a WWW-Authenticate challenge only for a validated downstream interaction requirement", async () => {
  vi.spyOn(OnboardingService.prototype, "requireAdmin").mockImplementation(() => {});
  vi.spyOn(ReportingService.prototype, "reconcile").mockRejectedValue(new ReportingInteractionRequired(id(1)));
  const result = await request("admin/reporting/security", {});
  expect(result).toMatchObject({ status: 401, headers: { "WWW-Authenticate": expect.stringContaining("insufficient_claims") },
    jsonBody: { code: "reporting-interaction-required" } });
});
