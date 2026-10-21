import { HttpRequest, HttpResponse, InvocationContext } from "@azure/functions";
import { afterEach, beforeAll, expect, it, vi } from "vitest";
import { ApiError } from "../src/model.js";
import { AzureStore } from "../src/storage.js";
import { OnboardingService } from "../src/onboarding-service.js";
import { handler } from "../src/functions.js";

vi.mock("../src/auth.js", () => ({
  createAuthenticator: () => async () => ({
    tenantId: "10000000-0000-4000-8000-000000000001",
    objectId: "10000000-0000-4000-8000-000000000004",
  }),
}));

beforeAll(() => {
  process.env.TENANT_ID = "10000000-0000-4000-8000-000000000001";
  process.env.API_AUDIENCE = "10000000-0000-4000-8000-000000000002";
  process.env.SPA_CLIENT_ID = "10000000-0000-4000-8000-000000000003";
  process.env.BUSINESS_TABLE_URL = "https://example.table.core.windows.net";
  process.env.BUSINESS_BLOB_URL = "https://example.blob.core.windows.net";
});
afterEach(() => vi.restoreAllMocks());

async function response(path: string) {
  const request = new HttpRequest({ url: `https://example.test/api/v1/${path}`, method: "GET", params: { path } });
  const result = await handler(request, new InvocationContext({ functionName: "test", logHandler: vi.fn() }));
  return new HttpResponse(result);
}

it("sends a real JSON null for setup before the initial private import", async () => {
  vi.spyOn(OnboardingService.prototype, "setup").mockResolvedValue(null);
  const result = await response("onboarding/setup");
  expect(result.status).toBe(200);
  expect(result.headers.get("content-type")).toBe("application/json");
  expect(await result.text()).toBe("null");
});
it("sends a real JSON null when an approved budget has no allocation operation", async () => {
  vi.spyOn(OnboardingService.prototype, "allocation").mockResolvedValue(null);
  const result = await response(`budgets/bu-${"a".repeat(24)}/allocation`);
  expect(result.status).toBe(200);
  expect(await result.json()).toBeNull();
});
it("preserves normal JSON response objects", async () => {
  const metadata = { version: 1 as const, tenantId: process.env.TENANT_ID!, serviceId: "cowork", budgets: [] };
  vi.spyOn(OnboardingService.prototype, "setup").mockResolvedValue(metadata);
  const result = await response("onboarding/setup");
  expect(result.status).toBe(200);
  expect(await result.json()).toEqual(metadata);
});
it("returns setup-required as a structured 409, not a generic outage", async () => {
  vi.spyOn(AzureStore.prototype, "metadata").mockRejectedValue(
    new ApiError(409, "setup-required", "Budget setup is not complete. Open Request a new BU."),
  );
  const result = await response("context");
  expect(result.status).toBe(409);
  expect(result.headers.get("cache-control")).toBe("no-store");
  expect(await result.json()).toMatchObject({
    code: "setup-required", message: expect.stringContaining("Request a new BU"), correlationId: expect.any(String),
  });
});
