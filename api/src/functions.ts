import { app, type HttpRequest, type HttpResponseInit, type InvocationContext } from "@azure/functions";
import { randomUUID } from "node:crypto";
import type { z } from "zod";
import { createAuthenticator } from "./auth.js";
import { ApiError, decisionSchema, fulfilmentUpdateSchema, seedSchema, submissionSchema, uuid } from "./model.js";
import { RequestService } from "./service.js";
import { AzureStore, credential, required, storageErrorCode } from "./storage.js";
import { allocationSchema, buDecisionSchema, buSubmissionSchema, memberUpdateSchema, workUpdateSchema } from "./onboarding.js";
import { OnboardingService } from "./onboarding-service.js";
import { AzureOnboardingStore } from "./onboarding-storage.js";

let dependencies: { service: RequestService; onboarding: OnboardingService; authenticate: ReturnType<typeof createAuthenticator> } | undefined;
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
      onboarding: new OnboardingService(new AzureOnboardingStore(
        tenantId, required("BUSINESS_TABLE_URL"), required("BUSINESS_BLOB_URL"), credential(),
      ), (process.env.ONBOARDING_ADMIN_OBJECT_IDS ?? "").split(",").filter((id) => id.trim()).map((id) => uuid.parse(id.trim()))),
    };
  }
  return dependencies;
}
async function body<T>(request: HttpRequest, schema: z.ZodType<T>, maxBytes = 12000): Promise<T> {
  if (!request.headers.get("content-type")?.startsWith("application/json")) {
    throw new ApiError(415, "content-type", "Send an application/json request.");
  }
  const text = await request.text();
  if (Buffer.byteLength(text, "utf8") > maxBytes) throw new ApiError(413, "too-large", "Request body is too large.");
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
    const { authenticate, service, onboarding } = services();
    const caller = await authenticate(request.headers.get("authorization"));
    const path = request.params.path ?? "";
    const cursor = new URL(request.url).searchParams.get("cursor") ?? undefined;
    if (cursor && cursor.length > 4096) throw new ApiError(400, "invalid-cursor", "Invalid pagination cursor.");
    let result: unknown;
    if (path === "onboarding/capabilities" && request.method === "GET") {
      result = { isAdmin: onboarding.isAdmin(caller), bootstrapEnabled: onboarding.isAdmin(caller) && process.env.BOOTSTRAP_ENABLED === "true",
        objectId: caller.objectId, displayName: caller.displayName, contactHint: caller.contactHint };
    } else if (path === "onboarding" && request.method === "POST") result = await onboarding.submit(caller, await body(request, buSubmissionSchema));
    else if (path === "onboarding" && request.method === "GET") result = await onboarding.list(caller, false, cursor);
    else if (path === "onboarding/inbox" && request.method === "GET") result = await onboarding.list(caller, true, cursor);
    else if (path === "onboarding/setup" && request.method === "GET") result = await onboarding.setup(caller);
    else if (path === "onboarding/bootstrap" && request.method === "POST") {
      onboarding.requireAdmin(caller);
      if (process.env.BOOTSTRAP_ENABLED !== "true") throw new ApiError(403, "bootstrap-disabled", "Bootstrap import is disabled by the operator.");
      result = await onboarding.bootstrap(caller, await body(request, seedSchema, 1_048_576));
    } else if (/^onboarding\/[a-f0-9]{64}\/decision$/.test(path) && request.method === "POST") {
      result = await onboarding.decide(caller, path.split("/")[1], await body(request, buDecisionSchema));
    } else if (/^onboarding\/(?:[a-f0-9]{64}|allocation-[a-z0-9-]{1,64})$/.test(path) && request.method === "GET") {
      result = await onboarding.get(caller, path.split("/")[1]);
    } else if (/^onboarding\/(?:[a-f0-9]{64}|allocation-[a-z0-9-]{1,64})\/work$/.test(path) && request.method === "POST") {
      result = await onboarding.update(caller, path.split("/")[1], await body(request, workUpdateSchema));
    } else if (/^budgets\/[a-z0-9-]{1,64}\/allocation$/.test(path) && ["GET", "POST"].includes(request.method)) {
      result = await onboarding.allocation(caller, path.split("/")[1],
        request.method === "POST" ? await body(request, allocationSchema) : undefined);
    } else if (/^budgets\/[a-z0-9-]{1,64}\/members$/.test(path) && request.method === "POST") {
      result = await onboarding.member(caller, path.split("/")[1], await body(request, memberUpdateSchema));
    } else if (path === "context" && request.method === "GET") result = await service.context(caller);
    else if (path === "requests" && request.method === "POST") result = await service.submit(caller, await body(request, submissionSchema));
    else if (path === "requests" && request.method === "GET") result = await service.history(caller, cursor);
    else if (path === "inbox" && request.method === "GET") result = await service.inbox(caller, cursor);
    else if (/^requests\/[a-f0-9]{64}\/decision$/.test(path) && request.method === "POST") {
      result = await service.decide(caller, path.split("/")[1], await body(request, decisionSchema));
    } else if (/^requests\/[a-f0-9]{64}\/fulfilment$/.test(path) && request.method === "POST") {
      result = await service.fulfil(caller, path.split("/")[1], await body(request, fulfilmentUpdateSchema));
    } else throw new ApiError(404, "not-found", "API endpoint not found.");
    // Functions treats jsonBody: null as absent, but nullable endpoints promise JSON null.
    if (result === null) return { status: 200, headers: { ...headers, "Content-Type": "application/json" }, body: "null" };
    return { status: 200, headers, jsonBody: result };
  } catch (error) {
    const known = error instanceof ApiError;
    const status = known ? error.status : 503;
    const code = known ? error.code : "service-unavailable";
    if (!known) {
      const category = error instanceof Error ? error.name : "UnknownError";
      const candidate = storageErrorCode(error);
      const storageCode = candidate && /^[A-Za-z0-9_]{1,80}$/.test(candidate) ? candidate : "unclassified";
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
