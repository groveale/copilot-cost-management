import { BlobServiceClient, type BlockBlobClient } from "@azure/storage-blob";
import type { TokenCredential } from "@azure/identity";
import { z } from "zod";
import { ApiError } from "./model.js";
import { isStorageNotFound, readMetadata } from "./storage.js";
import type { ReportingStore, SecurityState } from "./reporting.js";
const stateSchema = z.object({
  version: z.literal(1), ready: z.boolean(), metadataEtag: z.string().nullable(),
  ownershipHash: z.string().regex(/^[a-f0-9]{64}$/), configurationHash: z.string().regex(/^[a-f0-9]{64}$/),
  checkedAt: z.iso.datetime(),
}).strict();
export class AzureReportingStore implements ReportingStore {
  private readonly metadataBlob: BlockBlobClient;
  private readonly stateBlob: BlockBlobClient;
  private stateLeaseId?: string;
  constructor(private readonly tenantId: string, blobUrl: string, credential: TokenCredential) {
    const container = new BlobServiceClient(blobUrl, credential).getContainerClient("metadata");
    this.metadataBlob = container.getBlockBlobClient("budgets.json");
    this.stateBlob = container.getBlockBlobClient("reporting-security.json");
  }
  async metadata() {
    try { return await readMetadata(this.metadataBlob, this.tenantId); }
    catch (error) { if (isStorageNotFound(error, "BlobNotFound")) return { metadata: null, etag: null }; throw error; }
  }
  async state(): Promise<SecurityState | null> {
    try {
      const content = await this.stateBlob.downloadToBuffer(0, 16384);
      return stateSchema.parse(JSON.parse(content.toString("utf8")));
    } catch (error) { if (isStorageNotFound(error, "BlobNotFound")) return null; throw error; }
  }
  async save(state: SecurityState) {
    if (!this.stateLeaseId) throw new Error("Reporting security writes require an active storage lease.");
    await this.stateBlob.uploadData(Buffer.from(JSON.stringify(stateSchema.parse(state))), {
      blobHTTPHeaders: { blobContentType: "application/json" },
      conditions: { leaseId: this.stateLeaseId },
    });
  }
  async locked<T>(operation: (assertLock: () => void, metadataLeaseId?: string) => Promise<T>): Promise<T> {
    try { await this.stateBlob.uploadData(Buffer.from(JSON.stringify({
      version: 1, ready: false, metadataEtag: null, ownershipHash: "0".repeat(64), configurationHash: "0".repeat(64), checkedAt: new Date().toISOString(),
    })), { conditions: { ifNoneMatch: "*" } }); }
    catch (error) {
      if (!(typeof error === "object" && error !== null && "statusCode" in error && [409, 412].includes(Number(error.statusCode)))) throw error;
    }
    const lease = this.stateBlob.getBlobLeaseClient();
    try { await lease.acquireLease(60); }
    catch { throw new ApiError(409, "reporting-security-busy", "Another ownership/security operation holds the reporting lock. Retry after it completes."); }
    let confirmedUntil = Date.now() + 60000;
    this.stateLeaseId = lease.leaseId;
    const metadataLease = await this.metadataBlob.exists() ? this.metadataBlob.getBlobLeaseClient() : undefined;
    try { await metadataLease?.acquireLease(60); }
    catch {
      await lease.releaseLease(); this.stateLeaseId = undefined;
      throw new ApiError(409, "reporting-security-busy", "Private metadata is locked by another operator. Retry after it completes.");
    }
    let lost = false;
    let renewal: Promise<unknown> | undefined;
    const timer = setInterval(() => {
      const started = Date.now();
      renewal ??= Promise.all([lease.renewLease(), metadataLease?.renewLease()])
        .then(() => { confirmedUntil = started + 60000; })
        .catch(() => { lost = true; }).finally(() => { renewal = undefined; });
    }, 15000);
    const assertLock = () => {
      // Do not start a 30-second downstream request while a stalled renewal
      // could let another operation acquire the lock before it completes.
      if (lost || Date.now() + 35000 >= confirmedUntil) {
        throw new ApiError(409, "reporting-security-lock-lost", "Reporting security lease was lost or renewal stalled. Retry verification; readiness was not confirmed.");
      }
    };
    try { return await operation(assertLock, metadataLease?.leaseId); }
    finally {
      clearInterval(timer); await renewal;
      await metadataLease?.releaseLease().catch(() => {});
      await lease.releaseLease().catch(() => {}); this.stateLeaseId = undefined;
    }
  }
}
