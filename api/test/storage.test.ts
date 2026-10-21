import { beforeEach, describe, expect, it, vi } from "vitest";
import { Readable } from "node:stream";
const sdk = vi.hoisted(() => ({
  getEntity: vi.fn(), createEntity: vi.fn(), updateEntity: vi.fn(), listEntities: vi.fn(),
  download: vi.fn(), uploadData: vi.fn(),
}));
vi.mock("@azure/data-tables", async (original) => ({
  ...await original<typeof import("@azure/data-tables")>(),
  TableClient: class {
    getEntity = sdk.getEntity; createEntity = sdk.createEntity;
    updateEntity = sdk.updateEntity; listEntities = sdk.listEntities;
  },
}));
vi.mock("@azure/storage-blob", () => ({
  BlobServiceClient: class {
    getContainerClient() { return { getBlockBlobClient: () => ({ download: sdk.download, uploadData: sdk.uploadData }) }; }
  },
}));
import { AzureStore, isStorageNotFound, storageErrorCode } from "../src/storage.js";
import { AzureOnboardingStore } from "../src/onboarding-storage.js";
import { ApiError } from "../src/model.js";
import type { RequestRecord } from "../src/contracts.js";

const id = "10000000-0000-4000-8000-000000000001";
const caller = { tenantId: id, objectId: id };
const request: RequestRecord = {
  version: 1, id: "a".repeat(64), operationId: id, tenantId: id, requesterId: id, budgetId: "example",
  justification: "We need capacity for a project.", grantKind: "permanent", createdAt: "2026-01-01T00:00:00Z",
  expiresAt: null, status: "pending", etag: "specific-etag",
};
const store = () => new AzureStore(id, "https://example.table.core.windows.net", "https://example.blob.core.windows.net",
  { getToken: async () => ({ token: "synthetic-not-a-credential", expiresOnTimestamp: 1 }) });
const onboarding = () => new AzureOnboardingStore(id, "https://example.table.core.windows.net", "https://example.blob.core.windows.net",
  { getToken: async () => ({ token: "synthetic-not-a-credential", expiresOnTimestamp: 1 }) });
beforeEach(() => vi.resetAllMocks());
describe("Azure Table persistence adapter", () => {
  describe("real Tables SDK error decoding", () => {
    it.each([
      { code: "ResourceNotFound", statusCode: 404, missingEntity: true },
      { code: "TableNotFound", statusCode: 404, missingEntity: false },
      { code: "AuthorizationPermissionMismatch", statusCode: 403, missingEntity: false },
    ])("handles the SDK OData shape without hiding $code", async ({ code, statusCode, missingEntity }) => {
      const { TableClient } = await vi.importActual<typeof import("@azure/data-tables")>("@azure/data-tables");
      const client = new TableClient("https://synthetic.table.core.windows.net", "Requests", {
        getToken: async () => ({ token: "synthetic-not-a-credential", expiresOnTimestamp: Date.now() + 60000 }),
      }, {
        retryOptions: { maxRetries: 0 },
        httpClient: {
          async sendRequest(request) {
            const values = new Headers({ "content-type": "application/json;odata=minimalmetadata" });
            return {
              request, status: statusCode,
              headers: {
                get: (name: string) => values.get(name) ?? undefined,
                has: (name: string) => values.has(name),
                set: (name: string, value: string) => values.set(name, value),
                delete: (name: string) => values.delete(name),
                toJSON: () => Object.fromEntries(values),
                [Symbol.iterator]: () => values.entries(),
              },
              bodyAsText: JSON.stringify({ "odata.error": { code, message: { lang: "en-US", value: "Synthetic SDK error." } } }),
            };
          },
        },
      });
      const error: unknown = await client.getEntity(id, request.id).catch((error: unknown) => error);
      expect(error).toMatchObject({ name: "RestError", statusCode, details: { odataError: { code } } });
      expect(storageErrorCode(error)).toBe(code);
      expect(isStorageNotFound(error, "ResourceNotFound")).toBe(missingEntity);
      sdk.getEntity.mockRejectedValue(error);
      if (missingEntity) {
        expect(await store().get(id, request.id)).toBeNull();
        expect(await store().routing(caller)).toBeNull();
        expect(await onboarding().get(request.id)).toBeNull();
      } else {
        await expect(store().get(id, request.id)).rejects.toBe(error);
        await expect(store().routing(caller)).rejects.toBe(error);
        await expect(onboarding().get(request.id)).rejects.toBe(error);
      }
    });
    it("supports parsed-response error codes and does not override a top-level failure", () => {
      const nested = { response: { parsedBody: { odataError: { code: "ResourceNotFound" } } }, statusCode: 404 };
      expect(isStorageNotFound(nested, "ResourceNotFound")).toBe(true);
      expect(isStorageNotFound({ ...nested, code: "AuthorizationFailure" }, "ResourceNotFound")).toBe(false);
      expect(isStorageNotFound({ ...nested, statusCode: 403 }, "ResourceNotFound")).toBe(false);
      expect(storageErrorCode({ details: { odataError: { code: 404 } } })).toBeUndefined();
      expect(storageErrorCode(null)).toBeUndefined();
    });
  });
  describe("workspace metadata initialization", () => {
    it("reports actionable setup-required without blocking the initial onboarding import", async () => {
      sdk.download.mockRejectedValue({ statusCode: 404, code: "BlobNotFound" });
      await expect(store().metadata()).rejects.toBeInstanceOf(ApiError);
      await expect(store().metadata()).rejects.toMatchObject({
        status: 409, code: "setup-required", message: expect.stringContaining('"Admin" > "Setup requests"'),
      });
      expect(await onboarding().metadata()).toEqual({ metadata: null, etag: null });
      expect(sdk.uploadData).not.toHaveBeenCalled();
    });
    it.each([
      { statusCode: 404, code: "ContainerNotFound" },
      { statusCode: 404, code: "ResourceNotFound" },
      { statusCode: 404 },
      { statusCode: 403, code: "AuthorizationPermissionMismatch" },
      { statusCode: 403, code: "BlobNotFound" },
      new Error("Storage connection failed."),
    ])("preserves non-initialization storage failures %#", async (error) => {
      sdk.download.mockRejectedValue(error);
      await expect(store().metadata()).rejects.toEqual(error);
      await expect(onboarding().metadata()).rejects.toEqual(error);
    });
    it("loads the existing workspace unchanged after the initial import", async () => {
      const metadata = { version: 1, tenantId: id, serviceId: "cowork", budgets: [] };
      sdk.download.mockResolvedValue({
        etag: "metadata-etag", readableStreamBody: Readable.from([Buffer.from(JSON.stringify(metadata))]),
      });
      expect(await store().metadata()).toEqual(metadata);
    });
    it("does not hide invalid metadata behind the first-run setup message", async () => {
      sdk.download.mockResolvedValue({
        etag: "metadata-etag", readableStreamBody: Readable.from([Buffer.from('{"version":99}')]),
      });
      await expect(store().metadata()).rejects.not.toBeInstanceOf(ApiError);
    });
  });
  it("persists searchable routing fields with a versioned payload, not an unconditional upsert", async () => {
    expect(await store().create(request)).toBe(true);
    expect(sdk.createEntity).toHaveBeenCalledWith({
      partitionKey: id, rowKey: request.id, requesterId: id, budgetId: "example",
      payload: JSON.stringify(Object.fromEntries(Object.entries(request).filter(([key]) => key !== "etag"))),
    });
  });
  it("round-trips durable fulfilment state while continuing to read legacy requests", async () => {
    const upgraded: RequestRecord = { ...request, status: "approved", fulfilment: {
      state: "applied", standard: { policyId: id, groupId: id }, enhanced: { policyId: id, groupId: id },
      appliedAt: "2026-01-01T00:00:00Z", appliedBy: id,
    } };
    for (const value of [request, upgraded]) {
      const { etag, ...payload } = value;
      sdk.getEntity.mockResolvedValue({ payload: JSON.stringify(payload), etag });
      expect(await store().get(id, request.id)).toEqual(value);
    }
    const { etag, ...payload } = upgraded;
    sdk.getEntity.mockResolvedValue({ payload: JSON.stringify({ ...payload, fulfilment: { ...payload.fulfilment, appliedAt: undefined } }), etag });
    await expect(store().get(id, request.id)).rejects.toThrow();
  });
  describe("private onboarding storage", () => {
      it("permits missing metadata only for BlobNotFound, not denied storage or a missing container", async () => {
        sdk.download.mockRejectedValue({ statusCode: 404, code: "BlobNotFound" });
        expect(await onboarding().metadata()).toEqual({ metadata: null, etag: null });
        for (const error of [{ statusCode: 404, code: "ContainerNotFound" }, { statusCode: 403, code: "AuthorizationFailure" }]) {
          sdk.download.mockRejectedValue(error);
          await expect(onboarding().metadata()).rejects.toEqual(error);
        }
      });
      it("validates metadata, preserves its ETag and conditionally creates/replaces the blob", async () => {
        const metadata = { version: 1 as const, tenantId: id, serviceId: "cowork", budgets: [] };
        sdk.download.mockResolvedValue({ etag: "metadata-etag", readableStreamBody: Readable.from([Buffer.from(JSON.stringify(metadata))]) });
        expect(await onboarding().metadata()).toEqual({ metadata, etag: "metadata-etag" });
        await onboarding().saveMetadata(metadata, "metadata-etag");
        expect(sdk.uploadData.mock.calls[0][1].conditions).toEqual({ ifMatch: "metadata-etag" });
        await onboarding().saveMetadata(metadata, null);
        expect(sdk.uploadData.mock.calls[1][1].conditions).toEqual({ ifNoneMatch: "*" });
        sdk.uploadData.mockRejectedValue({ statusCode: 412 });
        await expect(onboarding().saveMetadata(metadata, "stale")).rejects.toMatchObject({ status: 409 });
      });
      it("enforces the exact 1 MiB metadata byte limit", async () => {
        sdk.download.mockResolvedValue({ etag: "etag", readableStreamBody: Readable.from([Buffer.alloc(1_048_577)]) });
        await expect(onboarding().metadata()).rejects.toThrow("1 MiB");
        await expect(onboarding().saveMetadata({ version: 1, tenantId: id, serviceId: "x".repeat(1_048_576), budgets: [] }, null))
          .rejects.toMatchObject({ status: 413 });
        expect(sdk.uploadData).not.toHaveBeenCalled();
      });
      it("never turns a missing table or failed authorization into an empty operation", async () => {
        sdk.getEntity.mockRejectedValue({ statusCode: 404, code: "ResourceNotFound" });
        expect(await onboarding().get("operation")).toBeNull();
        for (const error of [{ statusCode: 404, code: "TableNotFound" }, { statusCode: 403, code: "AuthorizationPermissionMismatch" }]) {
          sdk.getEntity.mockRejectedValue(error);
          await expect(onboarding().get("operation")).rejects.toEqual(error);
        }
      });
      it("creates routing additively and refuses to overwrite a different existing mapping", async () => {
        sdk.createEntity.mockRejectedValue({ statusCode: 409 });
        sdk.getEntity.mockResolvedValue({ budgetId: "legal" });
        await onboarding().seedRouting(id, "legal");
        expect(sdk.updateEntity).not.toHaveBeenCalled();
        await expect(onboarding().seedRouting(id, "other")).rejects.toMatchObject({ status: 409 });
      });
    });
  it("maps duplicate conditional creates and stale conditional updates to conflicts", async () => {
    sdk.createEntity.mockRejectedValue({ statusCode: 409 });
    expect(await store().create(request)).toBe(false);
    sdk.updateEntity.mockRejectedValue({ statusCode: 412 });
    expect(await store().replace(request, "specific-etag")).toBe(false);
    expect(sdk.updateEntity.mock.calls[0][1]).toBe("Replace");
    expect(sdk.updateEntity.mock.calls[0][2]).toEqual({ etag: "specific-etag" });
    sdk.updateEntity.mockClear();
    expect(await store().replace(request, "*")).toBe(false);
    expect(sdk.updateEntity).not.toHaveBeenCalled();
  });
  it("does not mistake missing tables or identity failures for missing records", async () => {
    sdk.getEntity.mockRejectedValue({ statusCode: 404, code: "ResourceNotFound" });
    expect(await store().routing(caller)).toBeNull();
    for (const error of [{ statusCode: 404, code: "TableNotFound" }, { statusCode: 403, code: "AuthorizationPermissionMismatch" }]) {
      sdk.getEntity.mockRejectedValue(error);
      await expect(store().routing(caller)).rejects.toEqual(error);
      await expect(store().get(id, request.id)).rejects.toEqual(error);
    }
  });
  it("round-trips a stored entity and uses the service ETag", async () => {
    const { etag: _, ...payload } = request;
    sdk.getEntity.mockResolvedValue({ payload: JSON.stringify(payload), etag: "server-etag" });
    expect(await store().get(id, request.id)).toEqual({ ...request, etag: "server-etag" });
  });
  it("rejects corrupt or unsupported stored payloads", async () => {
    sdk.getEntity.mockResolvedValue({ payload: '{"version":99}', etag: "etag" });
    await expect(store().get(id, request.id)).rejects.toThrow();
  });
  it("constrains every continuation query to the current caller or currently owned budgets", async () => {
    const next = vi.fn().mockResolvedValue({ done: false, value: Object.assign([], { continuationToken: "next" }) });
    const byPage = vi.fn().mockReturnValue({ next });
    sdk.listEntities.mockReturnValue({ byPage });
    expect(await store().list(id, caller.objectId, undefined, "cursor")).toEqual({ items: [], cursor: "next" });
    expect(sdk.listEntities.mock.calls[0][0].queryOptions.filter).toBe(`PartitionKey eq '${id}' and requesterId eq '${id}'`);
    expect(byPage).toHaveBeenCalledWith({ maxPageSize: 50, continuationToken: "cursor" });
    await store().list(id, undefined, ["legal"]);
    expect(sdk.listEntities.mock.calls[1][0].queryOptions.filter).toBe(`PartitionKey eq '${id}' and (budgetId eq 'legal')`);
    expect(await store().list(id, undefined, [])).toEqual({ items: [] });
    expect(sdk.listEntities).toHaveBeenCalledTimes(2);
  });
});
