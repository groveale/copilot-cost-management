import { randomUUID } from "node:crypto";
import { TableClient, odata, type TableEntityResult } from "@azure/data-tables";
import { BlobServiceClient } from "@azure/storage-blob";
import type { TokenCredential } from "@azure/identity";
import type { Page } from "./contracts.js";
import type { StoredSync, SyncJob, SyncStore } from "./sync-contracts.js";
import { storedSyncSchema, syncTerminal } from "./sync.js";
import { isStorageNotFound } from "./storage.js";

export class AzureSyncStore implements SyncStore {
  private readonly table: TableClient;
  private readonly blobs;
  constructor(private readonly tenantId: string, tableUrl: string, blobUrl: string, credential: TokenCredential) {
    this.table = new TableClient(tableUrl, "SyncJobs", credential);
    this.blobs = new BlobServiceClient(blobUrl, credential).getContainerClient("sync-checkpoints");
  }
  private async read(entity: TableEntityResult<Record<string, unknown>>): Promise<StoredSync> {
    if (!entity.etag || typeof entity.blob !== "string" ||
        !entity.blob.startsWith(`${this.tenantId}/`) || entity.blob.includes("..")) throw new Error("Invalid sync state pointer.");
    const response = await this.blobs.getBlockBlobClient(entity.blob).download();
    if (!response.readableStreamBody || (response.contentLength ?? 0) > 1_048_576) throw new Error("Invalid/oversized sync checkpoint.");
    const chunks: Buffer[] = [];
    let length = 0;
    for await (const chunk of response.readableStreamBody) {
      const buffer = Buffer.from(chunk); length += buffer.length;
      if (length > 1_048_576) throw new Error("Sync state exceeds the checkpoint size limit.");
      chunks.push(buffer);
    }
    const job = storedSyncSchema.parse(JSON.parse(Buffer.concat(chunks).toString("utf8")));
    if (job.tenantId !== this.tenantId) throw new Error("Invalid stored sync tenant.");
    return { job, etag: entity.etag };
  }
  private async entry(row: string) {
    try { return await this.read(await this.table.getEntity(this.tenantId, row)); }
    catch (error) { if (isStorageNotFound(error, "ResourceNotFound")) return null; throw error; }
  }
  current() { return this.entry("current"); }
  async get(id: string) { return (await this.entry(id))?.job ?? null; }
  private async checkpoint(job: SyncJob) {
    storedSyncSchema.parse(job);
    const body = JSON.stringify(job);
    if (Buffer.byteLength(body) > 1_048_576) throw new Error("Sync inventory exceeds the 1 MiB checkpoint limit. Use a single-policy sync.");
    const blob = this.blobs.getBlockBlobClient(`${this.tenantId}/${job.id}/${randomUUID()}.json`);
    await blob.upload(body, Buffer.byteLength(body), { conditions: { ifNoneMatch: "*" } });
    return blob;
  }
  async save(job: SyncJob, etag?: string) {
    const blob = await this.checkpoint(job);
    const fields = { partitionKey: this.tenantId, blob: blob.name, updatedAt: job.updatedAt };
    try {
      await this.table.submitTransaction([
        etag ? ["update", { ...fields, rowKey: "current" }, "Replace", { etag }] :
          ["create", { ...fields, rowKey: "current" }],
        ["upsert", { ...fields, rowKey: job.id }, "Replace"],
      ]);
      return true;
    } catch (error) {
      const status = typeof error === "object" && error !== null && "statusCode" in error ? error.statusCode : undefined;
      if (status === 409 || status === 412) { await blob.delete(); return false; }
      throw error;
    }
  }
  async acknowledge(id: string, revision: number): Promise<SyncJob | null> {
    for (let attempt = 0; attempt < 5; attempt++) {
      const current = await this.current();
      const stored = current?.job.id === id ? current : await this.entry(id);
      if (!stored) return null;
      const job = stored.job;
      if (job.revision !== revision || !job.mirrorPending) return job;
      job.mirrorPending = false;
      if (current?.job.id === id) { if (await this.save(job, stored.etag)) return job; continue; }
      const blob = await this.checkpoint(job);
      try {
        await this.table.updateEntity({ partitionKey: this.tenantId, rowKey: id, blob: blob.name, updatedAt: job.updatedAt }, "Replace", { etag: stored.etag });
        return job;
      } catch (error) {
        if (typeof error === "object" && error !== null && "statusCode" in error && error.statusCode === 412) {
          await blob.delete(); continue;
        }
        throw error;
      }
    }
    throw new Error("Job mirror acknowledgement conflicted repeatedly. Refresh before reconciling.");
  }
  async prune(retentionDays: number) {
    const cutoff = new Date(Date.now() - retentionDays * 86400000);
    const referenced = new Set<string>();
    let removed = 0;
    for await (const entity of this.table.listEntities({ queryOptions: { filter: odata`PartitionKey eq ${this.tenantId}` } })) {
      if (typeof entity.blob !== "string" || !entity.rowKey || !entity.etag) throw new Error("Invalid checkpoint reference during maintenance.");
      if (entity.rowKey !== "current" && Date.parse(String(entity.updatedAt)) < cutoff.getTime()) {
        const { job } = await this.read(entity);
        const active = await this.current();
        if (syncTerminal(job.state) && active?.job.id !== job.id) {
          await this.table.deleteEntity(this.tenantId, entity.rowKey, { etag: entity.etag }); removed++; continue;
        }
      }
      referenced.add(entity.blob);
    }
    for await (const blob of this.blobs.listBlobsFlat({ prefix: `${this.tenantId}/` })) {
      if (!referenced.has(blob.name) && blob.properties.lastModified && blob.properties.lastModified < cutoff) {
        if (!blob.properties.etag) throw new Error("Checkpoint cleanup requires an explicit blob ETag.");
        await this.blobs.getBlockBlobClient(blob.name).delete({ conditions: { ifMatch: blob.properties.etag } });
      }
    }
    return removed;
  }
  async list(cursor?: string): Promise<Page<SyncJob>> {
    const pages = this.table.listEntities({ queryOptions: {
      filter: odata`PartitionKey eq ${this.tenantId} and RowKey ne ${"current"}`,
    } }).byPage({ maxPageSize: 20, continuationToken: cursor });
    const page = await pages.next();
    if (page.done) return { items: [] };
    const items: SyncJob[] = [];
    for (const entity of page.value) items.push((await this.read(entity)).job);
    return { items, cursor: page.value.continuationToken || undefined };
  }
}
