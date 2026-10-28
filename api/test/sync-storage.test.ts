import { Readable } from "node:stream";
import { beforeEach, describe, expect, it, vi } from "vitest";
import type { TokenCredential } from "@azure/identity";
import { AzureSyncStore } from "../src/sync-storage.js";
import type { SyncJob } from "../src/sync-contracts.js";

const fake = vi.hoisted(() => ({
  entities: new Map<string, Record<string, unknown>>(), blobs: new Map<string, string>(),
  revisions: 0, uploads: 0, failTransaction: false,
}));
vi.mock("@azure/data-tables", async (original) => {
  const actual = await original<typeof import("@azure/data-tables")>();
  return { ...actual, TableClient: class {
    async getEntity(_partition: string, key: string) {
      const row = fake.entities.get(key);
      if (!row) throw { statusCode: 404, code: "ResourceNotFound" };
      return structuredClone(row);
    }
    async submitTransaction(actions: [string, Record<string, unknown>, string?, { etag?: string }?][]) {
      if (fake.failTransaction) throw { statusCode: 412 };
      const [action, , , options] = actions[0];
      const current = fake.entities.get("current");
      if (action === "create" ? !!current : options?.etag !== current?.etag) throw { statusCode: 412 };
      for (const [, entity] of actions) {
        fake.entities.set(String(entity.rowKey), { ...entity, etag: String(++fake.revisions) });
      }
    }
    async updateEntity(entity: Record<string, unknown>, _mode: string, options: { etag: string }) {
      if (fake.entities.get(String(entity.rowKey))?.etag !== options.etag) throw { statusCode: 412 };
      fake.entities.set(String(entity.rowKey), { ...entity, etag: String(++fake.revisions) });
    }
    listEntities() {
      return {
        async *[Symbol.asyncIterator]() { for (const entity of [...fake.entities.values()]) yield structuredClone(entity); },
        byPage: () => ({ next: async () => ({ done: false, value: [...fake.entities.values()].filter((row) => row.rowKey !== "current") }) }),
      };
    }
    async deleteEntity(_partition: string, key: string, options: { etag: string }) {
      if (fake.entities.get(key)?.etag !== options.etag) throw new Error("Delete ETag mismatch.");
      fake.entities.delete(key);
    }
  } };
});
vi.mock("@azure/storage-blob", () => ({
  BlobServiceClient: class {
    getContainerClient() {
      return { async *listBlobsFlat() {
        for (const name of [...fake.blobs.keys()]) yield { name, properties: { lastModified: new Date("2020-01-01T00:00:00Z"), etag: "specific" } };
      }, getBlockBlobClient: (name: string) => ({
        name,
        upload: async (body: string) => { fake.blobs.set(name, body); fake.uploads++; },
        download: async () => ({ readableStreamBody: Readable.from([Buffer.from(fake.blobs.get(name)!)]), contentLength: Buffer.byteLength(fake.blobs.get(name)!) }),
        delete: async () => { fake.blobs.delete(name); },
      }) };
    }
  },
}));
const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const credential: TokenCredential = { getToken: async () => ({ token: "unused-test-token", expiresOnTimestamp: 0 }) };
const job: SyncJob = { version: 1, revision: 1, id: id(1), tenantId: id(2), initiatorId: id(3), actors: [id(3)],
  concurrency: 3, state: "queued", phase: "Awaiting validation", createdAt: "2026-10-28T13:00:00Z",
  updatedAt: "2026-10-28T13:00:00Z", inventoryComplete: false, policies: [], mirrorPending: true };
const store = () => new AzureSyncStore(job.tenantId, "https://example.table.core.windows.net", "https://example.blob.core.windows.net", credential);
beforeEach(() => {
  fake.entities.clear(); fake.blobs.clear(); fake.revisions = 0; fake.uploads = 0; fake.failTransaction = false;
});
describe("private sync checkpoint storage", () => {
  it("atomically advances the tenant gate and job history to an immutable validated checkpoint", async () => {
    const storage = store();
    expect(await storage.save(job)).toBe(true);
    expect(fake.entities.get("current")?.blob).toBe(fake.entities.get(job.id)?.blob);
    const first = await storage.current();
    expect(first?.job).toEqual(job);
    const next = { ...job, revision: 2, phase: "Updated" };
    expect(await storage.save(next, first!.etag)).toBe(true);
    expect(fake.blobs.size).toBe(2);
    expect((await storage.current())?.job.phase).toBe("Updated");
    fake.failTransaction = true;
    expect(await storage.save({ ...next, revision: 3 }, first!.etag)).toBe(false);
    expect(fake.blobs.size).toBe(2);
  });
  it("acknowledges an archived mirror without replacing the active execution gate", async () => {
    const storage = store();
    await storage.save(job);
    const first = await storage.current();
    const next = { ...job, id: id(4), state: "queued" as const };
    await storage.save(next, first!.etag);
    expect((await storage.acknowledge(job.id, 1))?.mirrorPending).toBe(false);
    expect((await storage.current())?.job.id).toBe(next.id);
  });
  it("does not acknowledge a newer revision using an old mirror response", async () => {
    const storage = store();
    await storage.save({ ...job, revision: 3 });
    expect((await storage.acknowledge(job.id, 2))?.mirrorPending).toBe(true);
  });
  it("rejects oversized checkpoints before uploading rather than overflowing a Table entity", async () => {
    const policy = { id: id(5), displayName: "x".repeat(255), type: "x".repeat(200), state: "pending" as const };
    await expect(store().save({ ...job, policies: Array.from({ length: 3000 }, () => policy) })).rejects.toThrow("1 MiB");
    expect(fake.uploads).toBe(0);
  });
  it("prunes expired terminal history and orphan revisions but preserves the current referenced checkpoint", async () => {
    const storage = store();
    await storage.save({ ...job, state: "completed", createdAt: "2020-01-01T00:00:00Z", updatedAt: "2020-01-01T00:00:00Z" });
    const first = await storage.current();
    await storage.save({ ...job, id: id(4) }, first!.etag);
    expect(await storage.prune(30)).toBe(1);
    expect(fake.entities.has(job.id)).toBe(false);
    expect(fake.entities.has(id(4))).toBe(true);
    const active = await storage.current();
    expect(active?.job.id).toBe(id(4));
    expect(fake.blobs.size).toBe(1);
  });
});
