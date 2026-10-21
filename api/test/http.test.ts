import { HttpRequest, InvocationContext } from "@azure/functions";
import { beforeAll, expect, it, vi } from "vitest";
import { handler } from "../src/functions.js";

beforeAll(() => {
  process.env.TENANT_ID = "10000000-0000-4000-8000-000000000001";
  process.env.API_AUDIENCE = "10000000-0000-4000-8000-000000000002";
  process.env.SPA_CLIENT_ID = "10000000-0000-4000-8000-000000000003";
  process.env.BUSINESS_TABLE_URL = "https://example.table.core.windows.net";
  process.env.BUSINESS_BLOB_URL = "https://example.blob.core.windows.net";
});
const cases: { url: string; headers: Record<string, string>; path?: string; method?: string }[] = [
  { url: "https://example.test/api/v1/context", headers: {} },
  { url: "https://example.test/api/v1/context?code=not-a-credential", headers: {} },
  { url: "https://example.test/api/v1/context", headers: { "x-functions-key": "not-a-credential" } },
  { url: "https://example.test/api/v1/context", headers: { "x-ms-client-principal": '{"oid":"forged"}' } },
  { url: "https://example.test/api/v1/onboarding/bootstrap", path: "onboarding/bootstrap", method: "POST", headers: { "x-functions-key": "not-a-credential" } },
  { url: "https://example.test/api/v1/onboarding/inbox", path: "onboarding/inbox", headers: { "x-ms-client-principal": '{"oid":"forged"}' } },
];
it.each(cases)("never treats Function keys or principal headers as user authentication", async ({ url, headers, path, method }) => {
  const context = new InvocationContext({ functionName: "test", logHandler: vi.fn() });
  const response = await handler(new HttpRequest({ url, method: method ?? "GET", headers, params: { path: path ?? "context" } }), context);
  expect(response.status).toBe(401);
  expect(response.headers).toMatchObject({ "Cache-Control": "no-store", "X-Correlation-ID": expect.any(String) });
  expect(response.jsonBody).toMatchObject({ code: "unauthorized" });
});
