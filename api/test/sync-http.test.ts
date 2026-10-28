import { HttpRequest, InvocationContext } from "@azure/functions";
import { afterEach, beforeAll, expect, it, vi } from "vitest";
import { handler } from "../src/functions.js";
import { OnboardingService } from "../src/onboarding-service.js";
import { SyncService } from "../src/sync.js";
import { ApiError } from "../src/model.js";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
vi.mock("../src/auth.js", () => ({
  createAuthenticator: () => async () => ({ tenantId: "10000000-0000-4000-8000-000000000001", objectId: "10000000-0000-4000-8000-000000000004" }),
}));
beforeAll(() => {
  process.env.TENANT_ID = id(1); process.env.API_AUDIENCE = id(2); process.env.SPA_CLIENT_ID = id(3);
  process.env.BUSINESS_TABLE_URL = "https://example.table.core.windows.net";
  process.env.BUSINESS_BLOB_URL = "https://example.blob.core.windows.net";
});
afterEach(() => vi.restoreAllMocks());
async function request(path: string, body?: unknown) {
  return handler(new HttpRequest({ url: `https://example.test/api/v1/${path}`, method: body === undefined ? "GET" : "POST", params: { path },
    headers: body === undefined ? {} : { "content-type": "application/json" }, ...(body === undefined ? {} : { body: { string: JSON.stringify(body) } }) }),
  new InvocationContext({ functionName: "test", logHandler: vi.fn() }));
}
it("rejects all sync settings, history and mutation routes for non-admins", async () => {
  vi.spyOn(OnboardingService.prototype, "requireAdmin").mockImplementation(() => { throw new ApiError(403, "admin", "Admin required."); });
  for (const path of ["admin/sync/settings", "admin/sync/current", "admin/sync/jobs", `admin/sync/jobs/${id(5)}`]) {
    expect((await request(path)).status).toBe(403);
    expect((await request(path, {})).status).toBe(403);
  }
});
it("preserves JSON null current state and validates mutations before service execution", async () => {
  vi.spyOn(OnboardingService.prototype, "requireAdmin").mockImplementation(() => {});
  vi.spyOn(SyncService.prototype, "current").mockResolvedValue(null);
  const current = await request("admin/sync/current");
  expect(current.status).toBe(200); expect(current.body).toBe("null");
  const update = vi.spyOn(SyncService.prototype, "update");
  for (const input of [{ action: "heartbeat" }, { action: "claim", graphToken: "never-accept" }, { action: "mirrored", revision: 0 }]) {
    expect((await request(`admin/sync/jobs/${id(5)}`, input)).status).toBe(400);
  }
  expect(update).not.toHaveBeenCalled();
  expect((await request("admin/sync/jobs", { operationId: id(5), concurrency: 9 })).status).toBe(400);
  expect((await request("admin/sync/maintenance", {})).status).toBe(400);
});
