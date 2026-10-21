import { randomUUID } from "node:crypto";
import { AzureCliCredential } from "@azure/identity";
import { TableClient } from "@azure/data-tables";
import { describe, expect, it } from "vitest";
import { AzureStore } from "../src/storage.js";
import type { RequestRecord } from "../src/contracts.js";

const account = process.env.TEST_STORAGE_ACCOUNT;
describe.skipIf(!account)("dedicated Azure Table integration", () => {
  it("round-trips conditional creation, paging, ETags and persisted decisions without production records", async () => {
    if (!account || !/^sttest[a-z0-9]{3,18}$/.test(account)) {
      throw new Error("TEST_STORAGE_ACCOUNT must be an explicitly provisioned, dedicated sttest* account.");
    }
    const tenantId = randomUUID();
    const token = new AzureCliCredential();
    const tableUrl = `https://${account}.table.core.windows.net`;
    const table = new TableClient(tableUrl, "Requests", token);
    const store = new AzureStore(tenantId, tableUrl, `https://${account}.blob.core.windows.net`, token);
    await table.createTable();
    const record: RequestRecord = {
      version: 1, id: randomUUID().replaceAll("-", "").repeat(2),
      tenantId, requesterId: randomUUID(), operationId: randomUUID(), budgetId: "integration",
      justification: "Synthetic conditional persistence integration test.", grantKind: "permanent",
      createdAt: new Date().toISOString(), expiresAt: null, status: "pending", etag: "",
    };
    let created = false;
    try {
      expect(await store.create(record)).toBe(true); created = true;
      expect(await store.create(record)).toBe(false);
      const saved = await store.get(tenantId, record.id);
      expect(saved).toMatchObject({ ...record, etag: expect.any(String) });
      expect(saved!.etag).not.toBe("");
      expect((await store.list(tenantId, record.requesterId)).items).toHaveLength(1);
      const decision: RequestRecord = {
        ...saved!, status: "approved", decisionBy: record.requesterId, decisionAt: new Date().toISOString(),
      };
      expect(await store.replace(decision, saved!.etag)).toBe(true);
      expect(await store.replace({ ...decision, status: "rejected" }, saved!.etag)).toBe(false);
      expect(await store.get(tenantId, record.id)).toMatchObject({ status: "approved", decisionBy: record.requesterId });
    } finally {
      if (created) await table.deleteEntity(tenantId, record.id);
    }
  }, 60000);
});
