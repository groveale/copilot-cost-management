import { TableClient, odata, type TableEntityResult } from "@azure/data-tables";
import { BlobServiceClient } from "@azure/storage-blob";
import { AzureCliCredential, ManagedIdentityCredential, type TokenCredential } from "@azure/identity";
import { z } from "zod";
import type { Caller, Metadata, Page, RequestRecord } from "./contracts.js";
import { parseMetadata, validateStoredRequest } from "./model.js";
import type { Store } from "./service.js";

export function credential(): TokenCredential {
  if (process.env.DEMO_LOCAL_AUTH === "azure-cli" && !process.env.WEBSITE_INSTANCE_ID) return new AzureCliCredential();
  return new ManagedIdentityCredential();
}
export function required(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) throw new Error(`Missing required setting: ${name}.`);
  return value;
}
function status(error: unknown, code: number): boolean {
  return typeof error === "object" && error !== null && "statusCode" in error && error.statusCode === code;
}
function entityNotFound(error: unknown): boolean {
  return status(error, 404) && typeof error === "object" && error !== null &&
    "code" in error && error.code === "ResourceNotFound";
}
const storedSchema = z.object({
  version: z.literal(1), id: z.string().regex(/^[a-f0-9]{64}$/),
  operationId: z.uuid(), tenantId: z.uuid(), requesterId: z.uuid(), budgetId: z.string(),
  justification: z.string(), grantKind: z.enum(["permanent", "billingPeriod"]),
  createdAt: z.iso.datetime(), expiresAt: z.iso.datetime().nullable(),
  status: z.enum(["pending", "approved", "rejected"]),
  decisionBy: z.uuid().optional(), decisionAt: z.iso.datetime().optional(), rejectionReason: z.string().optional(),
}).strict();
function record(entity: TableEntityResult<Record<string, unknown>>): RequestRecord {
  if (typeof entity.payload !== "string" || !entity.etag) throw new Error("Invalid stored request entity.");
  return validateStoredRequest({ ...storedSchema.parse(JSON.parse(entity.payload)), etag: entity.etag });
}
function entity(record: RequestRecord) {
  const { etag: _, ...payload } = record;
  return {
    partitionKey: record.tenantId, rowKey: record.id,
    requesterId: record.requesterId, budgetId: record.budgetId, payload: JSON.stringify(payload),
  };
}
export class AzureStore implements Store {
  private readonly requests: TableClient;
  private readonly index: TableClient;
  private readonly blobs: BlobServiceClient;
  constructor(private readonly tenantId: string, tableUrl: string, blobUrl: string, token: TokenCredential) {
    this.requests = new TableClient(tableUrl, "Requests", token);
    this.index = new TableClient(tableUrl, "RequesterBudgetIndex", token);
    this.blobs = new BlobServiceClient(blobUrl, token);
  }
  async metadata(): Promise<Metadata> {
    const response = await this.blobs.getContainerClient("metadata").getBlobClient("budgets.json").download();
    if (!response.readableStreamBody) throw new Error("Metadata blob has no content.");
    const chunks: Buffer[] = [];
    let length = 0;
    for await (const chunk of response.readableStreamBody) {
      const buffer = Buffer.from(chunk);
      length += buffer.length;
      if (length > 1_048_576) throw new Error("Metadata exceeds the 1 MiB demo limit.");
      chunks.push(buffer);
    }
    return parseMetadata(JSON.parse(Buffer.concat(chunks).toString("utf8")), this.tenantId);
  }
  async routing(caller: Caller): Promise<string | null> {
    try {
      const entry = await this.index.getEntity<{ budgetId: string }>(caller.tenantId, caller.objectId);
      if (typeof entry.budgetId !== "string" || !entry.budgetId) throw new Error("Invalid requester index record.");
      return entry.budgetId;
    } catch (error) {
      if (entityNotFound(error)) return null;
      throw error;
    }
  }
  async get(tenant: string, id: string): Promise<RequestRecord | null> {
    try { return record(await this.requests.getEntity(tenant, id)); }
    catch (error) { if (entityNotFound(error)) return null; throw error; }
  }
  async create(request: RequestRecord): Promise<boolean> {
    try { await this.requests.createEntity(entity(request)); return true; }
    catch (error) { if (status(error, 409)) return false; throw error; }
  }
  async replace(request: RequestRecord, etag: string): Promise<boolean> {
    if (!etag || etag === "*") return false;
    try { await this.requests.updateEntity(entity(request), "Replace", { etag }); return true; }
    catch (error) { if (status(error, 412)) return false; throw error; }
  }
  async list(tenant: string, requesterId?: string, budgets?: string[], cursor?: string): Promise<Page<RequestRecord>> {
    let filter = odata`PartitionKey eq ${tenant}`;
    if (requesterId) filter += odata` and requesterId eq ${requesterId}`;
    if (budgets) {
      if (!budgets.length) return { items: [] };
      filter += ` and (${budgets.map((id) => odata`budgetId eq ${id}`).join(" or ")})`;
    }
    const iterator = this.requests.listEntities<Record<string, unknown>>({ queryOptions: { filter } })
      .byPage({ maxPageSize: 50, continuationToken: cursor });
    const next = await iterator.next();
    if (next.done) return { items: [] };
    return { items: next.value.map(record), cursor: next.value.continuationToken || undefined };
  }
}
