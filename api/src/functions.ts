import { app, type HttpRequest, type HttpResponseInit, type InvocationContext } from "@azure/functions";
import { randomUUID } from "node:crypto";
import type { z } from "zod";
import { createAuthenticator } from "./auth.js";
import { ApiError, decisionSchema, submissionSchema, uuid } from "./model.js";
import { RequestService } from "./service.js";
import { AzureStore, credential, required } from "./storage.js";

let dependencies: { service: RequestService; authenticate: ReturnType<typeof createAuthenticator> } | undefined;
function services() {
  if (!dependencies) {
    const tenantId = uuid.parse(required("TENANT_ID"));
    dependencies = {
      authenticate: createAuthenticator({
        tenantId, audience: uuid.parse(required("API_AUDIENCE")), spaClientId: uuid.parse(required("SPA_CLIENT_ID")),
      }),
      service: new RequestService(new AzureStore(
        tenantId, required("BUSINESS_TABLE_URL"), required("BUSINESS_BLOB_URL"), credential(),
      )),
    };
  }
  return dependencies;
}
async function body<T>(request: HttpRequest, schema: z.ZodType<T>): Promise<T> {
  if (!request.headers.get("content-type")?.startsWith("application/json")) {
    throw new ApiError(415, "content-type", "Send an application/json request.");
  }
  const text = await request.text();
  if (text.length > 12000) throw new ApiError(413, "too-large", "Request body is too large.");
  let value: unknown;
  try { value = JSON.parse(text); }
  catch { throw new ApiError(400, "invalid-json", "The request body must be valid JSON."); }
  const result = schema.safeParse(value);
  if (!result.success) throw new ApiError(400, "invalid-input", "Invalid fields, lengths or request shape.");
  return result.data;
}
export async function handler(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  const correlationId = randomUUID();
  const headers = { "Cache-Control": "no-store", "X-Correlation-ID": correlationId, "X-Content-Type-Options": "nosniff" };
  try {
    const { authenticate, service } = services();
    const caller = await authenticate(request.headers.get("authorization"));
    const path = request.params.path ?? "";
    const cursor = new URL(request.url).searchParams.get("cursor") ?? undefined;
    if (cursor && cursor.length > 4096) throw new ApiError(400, "invalid-cursor", "Invalid pagination cursor.");
    let result: unknown;
    if (path === "context" && request.method === "GET") result = await service.context(caller);
    else if (path === "requests" && request.method === "POST") result = await service.submit(caller, await body(request, submissionSchema));
    else if (path === "requests" && request.method === "GET") result = await service.history(caller, cursor);
    else if (path === "inbox" && request.method === "GET") result = await service.inbox(caller, cursor);
    else if (/^requests\/[a-f0-9]{64}\/decision$/.test(path) && request.method === "POST") {
      result = await service.decide(caller, path.split("/")[1], await body(request, decisionSchema));
    } else throw new ApiError(404, "not-found", "API endpoint not found.");
    return { status: 200, headers, jsonBody: result };
  } catch (error) {
    const known = error instanceof ApiError;
    const status = known ? error.status : 503;
    const code = known ? error.code : "service-unavailable";
    if (!known) {
      const category = error instanceof Error ? error.name : "UnknownError";
      const storageCode = typeof error === "object" && error !== null && "code" in error &&
        typeof error.code === "string" && /^[A-Za-z0-9_]{1,80}$/.test(error.code) ? error.code : "unclassified";
      context.error(`API failure ${correlationId}; category=${category}; code=${storageCode}. Check identity, storage and configuration.`);
    }
    return {
      status, headers, jsonBody: {
        code, correlationId,
        message: known ? error.message : "The service could not complete this operation. Keep your draft and retry; contact the administrator with the correlation ID.",
      },
    };
  }
}
app.http("requestsApi", { route: "v1/{*path}", methods: ["GET", "POST"], authLevel: "anonymous", handler });
app.http("health", {
  route: "health", methods: ["GET"], authLevel: "anonymous",
  handler: () => ({ status: 200, jsonBody: { status: "running" }, headers: { "Cache-Control": "no-store" } }),
});
