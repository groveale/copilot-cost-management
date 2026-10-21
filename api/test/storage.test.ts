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
import { AzureStore } from "../src/storage.js";
import { AzureOnboardingStore } from "../src/onboarding-storage.js";
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
  it("persists searchable routing fields with a versioned payload, not an unconditional upsert", async () => {
    expect(await store().create(request)).toBe(true);
    expect(sdk.createEntity).toHaveBeenCalledWith({
      partitionKey: id, rowKey: request.id, requesterId: id, budgetId: "example",
      payload: JSON.stringify(Object.fromEntries(Object.entries(request).filter(([key]) => key !== "etag"))),
    });
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
