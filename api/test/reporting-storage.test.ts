import { beforeEach, expect, it, vi } from "vitest";
const sdk = vi.hoisted(() => ({
  uploadData: vi.fn(), downloadToBuffer: vi.fn(), exists: vi.fn(), acquireLease: vi.fn(), releaseLease: vi.fn(), renewLease: vi.fn(),
}));
vi.mock("@azure/storage-blob", () => ({
  BlobServiceClient: class {
    getContainerClient() { return { getBlockBlobClient: (name: string) => ({
      uploadData: (...args: unknown[]) => sdk.uploadData(name, ...args),
      downloadToBuffer: sdk.downloadToBuffer, exists: sdk.exists,
      getBlobLeaseClient: () => ({ leaseId: `${name}-lease`, acquireLease: sdk.acquireLease, releaseLease: sdk.releaseLease, renewLease: sdk.renewLease }),
    }) }; }
  },
}));
import { AzureReportingStore } from "../src/reporting-storage.js";
import type { SecurityState } from "../src/reporting.js";
const state: SecurityState = { version: 1, ready: true, metadataEtag: "v1", ownershipHash: "a".repeat(64), configurationHash: "b".repeat(64), checkedAt: "2026-11-04T12:00:00Z" };
const store = () => new AzureReportingStore("10000000-0000-4000-8000-000000000001", "https://example.blob.core.windows.net",
  { getToken: async () => null });
beforeEach(() => {
  vi.resetAllMocks(); sdk.exists.mockResolvedValue(true);
  sdk.acquireLease.mockResolvedValue({}); sdk.releaseLease.mockResolvedValue({}); sdk.renewLease.mockResolvedValue({});
});
it("persists ready state under a real state-blob lease and leases metadata against concurrent writers", async () => {
  const value = store();
  await value.locked(async (assert, metadataLeaseId) => {
    assert(); expect(metadataLeaseId).toBe("budgets.json-lease");
    await value.save(state);
  });
  expect(sdk.acquireLease).toHaveBeenCalledTimes(2);
  expect(sdk.uploadData).toHaveBeenLastCalledWith("reporting-security.json", expect.any(Buffer), expect.objectContaining({
    conditions: { leaseId: "reporting-security.json-lease" },
  }));
  expect(sdk.releaseLease).toHaveBeenCalledTimes(2);
  await expect(value.save(state)).rejects.toThrow("active storage lease");
});
it("does not execute when another durable security operation holds the lease", async () => {
  sdk.acquireLease.mockRejectedValueOnce(new Error("lease held"));
  const operation = vi.fn();
  await expect(store().locked(operation)).rejects.toMatchObject({ code: "reporting-security-busy" });
  expect(operation).not.toHaveBeenCalled();
});
it("does not treat corruption, forbidden storage or missing containers as a ready/default state", async () => {
  sdk.downloadToBuffer.mockResolvedValueOnce(Buffer.from("malformed"));
  await expect(store().state()).rejects.toThrow();
  for (const error of [{ statusCode: 403, code: "AuthorizationFailure" }, { statusCode: 404, code: "ContainerNotFound" }]) {
    sdk.downloadToBuffer.mockRejectedValueOnce(error);
    await expect(store().state()).rejects.toBe(error);
  }
  sdk.downloadToBuffer.mockRejectedValueOnce({ statusCode: 404, code: "BlobNotFound" });
  expect(await store().state()).toBeNull();
});
it("releases both leases after a partially failed permission operation", async () => {
  await expect(store().locked(async () => { throw new Error("revocation failed"); })).rejects.toThrow("revocation failed");
  expect(sdk.releaseLease).toHaveBeenCalledTimes(2);
});
it("stops starting downstream requests before a stalled renewal can expire its lease", async () => {
  vi.useFakeTimers();
  try {
    await store().locked(async (assert) => {
      vi.setSystemTime(Date.now() + 30000);
      expect(assert).toThrow("renewal stalled");
    });
  } finally { vi.useRealTimers(); }
});
