import { TableClient, odata, type TableEntityResult } from "@azure/data-tables";
import { BlobServiceClient } from "@azure/storage-blob";
import type { TokenCredential } from "@azure/identity";
import type { Metadata } from "./contracts.js";
import { ApiError } from "./model.js";
import { isStorageNotFound, readMetadata } from "./storage.js";
import { onboardingRecordSchema, type OnboardingRecord, type OnboardingStore } from "./onboarding.js";

function hasStatus(error: unknown, statusCode: number) {
  return typeof error === "object" && error !== null && "statusCode" in error && error.statusCode === statusCode;
}
function record(entity: TableEntityResult<Record<string, unknown>>): OnboardingRecord {
  if (typeof entity.payload !== "string" || !entity.etag) throw new Error("Invalid stored onboarding operation.");
  const value = onboardingRecordSchema.parse({ ...JSON.parse(entity.payload), etag: entity.etag });
  if (value.id !== entity.rowKey || value.tenantId !== entity.partitionKey) throw new Error("Stored onboarding payload does not match its entity keys.");
  return value;
}
function entity(record: OnboardingRecord) {
  const { etag: _, ...payload } = record;
  return { partitionKey: record.tenantId, rowKey: record.id, kind: record.kind, requesterId: record.requesterId,
    payload: JSON.stringify(payload) };
}
export class AzureOnboardingStore implements OnboardingStore {
  private readonly requests: TableClient;
  private readonly index: TableClient;
  private readonly blob;
  constructor(private readonly tenantId: string, tableUrl: string, blobUrl: string, token: TokenCredential) {
    this.requests = new TableClient(tableUrl, "OnboardingRequests", token);
    this.index = new TableClient(tableUrl, "RequesterBudgetIndex", token);
    this.blob = new BlobServiceClient(blobUrl, token).getContainerClient("metadata").getBlockBlobClient("budgets.json");
  }
  async get(id: string) {
    try { return record(await this.requests.getEntity(this.tenantId, id)); }
    catch (error) { if (isStorageNotFound(error, "ResourceNotFound")) return null; throw error; }
  }
  async create(value: OnboardingRecord) {
    try { await this.requests.createEntity(entity(value)); return true; }
    catch (error) { if (hasStatus(error, 409)) return false; throw error; }
  }
  async replace(value: OnboardingRecord, etag: string) {
    if (!etag || etag === "*") throw new Error("A specific onboarding entity ETag is required.");
    try { await this.requests.updateEntity(entity(value), "Replace", { etag }); return true; }
    catch (error) { if (hasStatus(error, 412)) return false; throw error; }
  }
  async list(requesterId?: string, cursor?: string) {
    let filter = odata`PartitionKey eq ${this.tenantId} and kind eq 'onboarding'`;
    if (requesterId) filter += odata` and requesterId eq ${requesterId}`;
    const iterator = this.requests.listEntities<Record<string, unknown>>({ queryOptions: { filter } })
      .byPage({ maxPageSize: 50, continuationToken: cursor });
    const next = await iterator.next();
    return next.done ? { items: [] } : { items: next.value.map(record).filter((value) =>
      value.tenantId === this.tenantId && value.kind === "onboarding" && (!requesterId || value.requesterId === requesterId)),
      cursor: next.value.continuationToken || undefined };
  }
  async metadata() {
    try { return await readMetadata(this.blob, this.tenantId); }
    catch (error) {
      if (isStorageNotFound(error, "BlobNotFound")) return { metadata: null, etag: null };
      throw error;
    }
  }
  async saveMetadata(metadata: Metadata, etag: string | null) {
    const data = Buffer.from(JSON.stringify(metadata));
    if (data.length > 1_048_576) throw new ApiError(413, "too-large", "Metadata exceeds the 1 MiB demo limit.");
    try {
      await this.blob.uploadData(data, { conditions: etag ? { ifMatch: etag } : { ifNoneMatch: "*" },
        blobHTTPHeaders: { blobContentType: "application/json" } });
    } catch (error) {
      if (hasStatus(error, 409) || hasStatus(error, 412)) throw new ApiError(409, "metadata-conflict", "Metadata changed. Refresh and retry; existing data was not reset.");
      throw error;
    }
  }
  async routing(objectId: string): Promise<string | null> {
    try {
      const entry = await this.index.getEntity<{ budgetId: string }>(this.tenantId, objectId);
      if (typeof entry.budgetId !== "string" || !entry.budgetId) throw new Error("Invalid requester index record.");
      return entry.budgetId;
    } catch (error) {
      if (isStorageNotFound(error, "ResourceNotFound")) return null;
      throw error;
    }
  }
  async seedRouting(objectId: string, budgetId: string): Promise<void> {
    try {
      await this.index.createEntity({ partitionKey: this.tenantId, rowKey: objectId, budgetId });
    } catch (error) {
      if (!hasStatus(error, 409)) throw error;
      const existing = await this.index.getEntity<{ budgetId: string }>(this.tenantId, objectId);
      if (existing.budgetId !== budgetId) throw new ApiError(409, "routing-conflict",
        "A requester already has a different mapping. Existing routing was not overwritten.");
    }
  }
}
