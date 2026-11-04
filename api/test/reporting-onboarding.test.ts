import { beforeEach, expect, it, vi } from "vitest";
const sdk = vi.hoisted(() => ({ uploadData: vi.fn(), guard: vi.fn() }));
vi.mock("@azure/data-tables", () => ({ TableClient: class {} }));
vi.mock("@azure/storage-blob", () => ({
  BlobServiceClient: class {
    getContainerClient() { return { getBlockBlobClient: () => ({ uploadData: sdk.uploadData }) }; }
  },
}));
vi.mock("../src/reporting-runtime.js", () => ({ guardMetadataWrite: sdk.guard }));
import { AzureOnboardingStore } from "../src/onboarding-storage.js";
const tenantId = "10000000-0000-4000-8000-000000000001";
const metadata = { version: 1 as const, tenantId, serviceId: "cowork", budgets: [] };
const store = () => new AzureOnboardingStore(tenantId, "https://example.table.core.windows.net",
  "https://example.blob.core.windows.net", { getToken: async () => null }, "Bearer synthetic-api-assertion");
beforeEach(() => vi.resetAllMocks());
it("routes every bootstrap/onboarding metadata write through the security guard with the actual metadata lease", async () => {
  sdk.guard.mockImplementation(async (next, etag, write, authorization) => {
    expect(next).toEqual(metadata); expect(etag).toBe("metadata-v1");
    expect(authorization).toBe("Bearer synthetic-api-assertion");
    await write("verified-metadata-lease");
  });
  await store().saveMetadata(metadata, "metadata-v1");
  expect(sdk.uploadData).toHaveBeenCalledOnce();
  expect(sdk.uploadData).toHaveBeenCalledWith(Buffer.from(JSON.stringify(metadata)), {
    conditions: { ifMatch: "metadata-v1", leaseId: "verified-metadata-lease" },
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
});
it("never saves changed ownership metadata when the guard reports partial revocation", async () => {
  sdk.guard.mockRejectedValue(new Error("revocation incomplete"));
  await expect(store().saveMetadata(metadata, "metadata-v1")).rejects.toThrow("revocation incomplete");
  expect(sdk.uploadData).not.toHaveBeenCalled();
});
