import { readFile } from "node:fs/promises";
import { AzureCliCredential } from "@azure/identity";
import { BlobServiceClient } from "@azure/storage-blob";
import { TableClient } from "@azure/data-tables";
import { seedSchema } from "./model.js";

const [file, tableUrl, blobUrl, tenantId, mode] = process.argv.slice(2);
if (!file || !tableUrl || !blobUrl || !tenantId || !["validate", "apply"].includes(mode)) {
  throw new Error("Usage: node dist/seed.js <private-file> <table-url> <blob-url> <tenant-id> <validate|apply>");
}
for (const endpoint of [tableUrl, blobUrl]) {
  if (new URL(endpoint).protocol !== "https:") throw new Error("Azure HTTPS endpoints are required.");
}
const seed = seedSchema.parse(JSON.parse(await readFile(file, "utf8")));
if (seed.metadata.tenantId !== tenantId) throw new Error("Seed tenant does not match deployment tenant.");
console.log(`Validated ${seed.metadata.budgets.length} budgets and ${seed.requesterIndex.length} explicit requester mappings. No records are deleted.`);
if (mode === "apply") {
  const token = new AzureCliCredential({ tenantId });
  const index = new TableClient(tableUrl, "RequesterBudgetIndex", token);
  const blob = new BlobServiceClient(blobUrl, token).getContainerClient("metadata").getBlockBlobClient("budgets.json");
  // Metadata replacement is conditional: concurrent operator edits must not be lost.
  const current = await blob.exists() ? await blob.getProperties() : undefined;
  await blob.uploadData(Buffer.from(JSON.stringify(seed.metadata)), {
    conditions: current ? { ifMatch: current.etag } : { ifNoneMatch: "*" },
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
  for (const mapping of seed.requesterIndex) {
    await index.upsertEntity({
      partitionKey: tenantId, rowKey: mapping.objectId, budgetId: mapping.budgetId,
    }, "Replace");
  }
  console.log("Private metadata and the explicitly supplied mappings were saved. Existing requests were not modified.");
}
