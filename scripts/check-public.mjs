import { readFile, readdir, stat } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import assert from "node:assert/strict";

const root = new URL("../demo/dist/", import.meta.url);
const config = JSON.parse(await readFile(new URL("demo-config.json", root), "utf8"));
assert.deepEqual(Object.keys(config).sort(), ["apiBaseUrl", "apiScope", "budgets", "clientId", "serviceId", "tenantId", "version"].sort());
assert.deepEqual(config.budgets, []);
await stat(new URL("auth.html", root));
await stat(new URL("staticwebapp.config.json", root));
async function inspect(path) {
  for (const entry of await readdir(path, { withFileTypes: true })) {
    const target = join(path, entry.name);
    assert(!/private|seed|\.pdf$|\.map$|\.test\.|local\.settings|\.env/i.test(entry.name), `Forbidden public file: ${entry.name}`);
    if (entry.isDirectory()) await inspect(target);
    else {
      const content = await readFile(target, "utf8");
      assert(!/AccountKey=|SharedAccessSignature=|-----BEGIN .*PRIVATE KEY-----|[?&]sig=[a-zA-Z0-9%+/]{12}/i.test(content), `Credential material in ${entry.name}`);
      assert(!/admin@m365cpi|LEGAL-DEMO-00|legal\.requester@example\.test/i.test(content), `Private/fixture data in ${entry.name}`);
    }
  }
}
await inspect(fileURLToPath(root));
console.log("Public output contains only the application, callback and public bootstrap configuration.");
