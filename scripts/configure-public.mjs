import { writeFile } from "node:fs/promises";

const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
for (const key of ["AZURE_TENANT_ID", "SPA_CLIENT_ID", "API_CLIENT_ID"]) {
  if (!guid.test(process.env[key] ?? "") || /^0{8}-0{4}-0{4}-0{4}-0{12}$/.test(process.env[key])) {
    throw new Error(`Missing/invalid ${key}.`);
  }
}
const url = new URL(process.env.API_BASE_URL ?? "");
if (url.protocol !== "https:" || url.origin !== process.env.API_BASE_URL || url.hostname === "configure-your-api.azurewebsites.net") {
  throw new Error("API_BASE_URL must be the actual HTTPS origin, without a trailing slash.");
}
await writeFile(new URL("../demo/dist/demo-config.json", import.meta.url), JSON.stringify({
  version: 1, tenantId: process.env.AZURE_TENANT_ID, clientId: process.env.SPA_CLIENT_ID,
  apiBaseUrl: url.origin, apiScope: `api://${process.env.API_CLIENT_ID}/access_as_user`, serviceId: "cowork", budgets: [],
}, null, 2));
