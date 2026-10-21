import { readFile } from "node:fs/promises";
import assert from "node:assert/strict";
import test from "node:test";

const read = (path) => readFile(new URL(`../${path}`, import.meta.url), "utf8");
test("Function runtime has no key-based configuration or identity selectors", async () => {
  const main = await read("infra/main.bicep");
  assert.match(main, /identity: \{ type: 'SystemAssigned' \}/);
  assert.match(main, /authentication: \{ type: 'SystemAssignedIdentity' \}/);
  assert.match(main, /AzureWebJobsStorage__credential', value: 'managedidentity'/);
  assert.equal((main.match(/allowSharedKeyAccess: false/g) ?? []).length, 2);
  assert.equal((main.match(/publicNetworkAccess: 'Disabled'/g) ?? []).length, 2);
  assert.equal((main.match(/bypass: 'None', defaultAction: 'Deny'/g) ?? []).length, 2);
  assert.match(main, /virtualNetworkSubnetId: network.outputs.integrationSubnetId/);
  assert.match(main, /DisableLocalAuth: true/);
  assert.match(main, /Authorization=AAD/);
  assert.equal((main.match(/properties: \{ allow: false \}/g) ?? []).length, 2);
  assert.doesNotMatch(main, /listKeys|AccountKey|SharedAccessSignature|AzureWebJobsStorage__clientId|UserAssignedIdentity/);
  const source = await read("api/src/storage.ts");
  assert.match(source, /new ManagedIdentityCredential\(\)/);
  assert.doesNotMatch(source, /DefaultAzureCredential|AzureNamedKeyCredential|AzureSASCredential|ClientSecretCredential|fromConnectionString/);
});
test("private storage uses separate Flex and endpoint subnets with linked Blob/Table DNS", async () => {
  const network = await read("infra/network.bicep");
  assert.match(network, /serviceName: 'Microsoft.App\/environments'/);
  assert.match(network, /cidrSubnet\(addressPrefix, 27, 0\)/);
  assert.match(network, /cidrSubnet\(addressPrefix, 27, 1\)/);
  assert.match(network, /name: 'private-endpoints'/);
  assert.match(network, /subnet: \{ id: '\$\{vnet.id\}\/subnets\/private-endpoints' \}/);
  assert.match(network, /var services = \[ 'blob', 'table' \]/);
  assert.match(network, /registrationEnabled: false/);
  assert.match(network, /privateDnsZoneId: dnsZones\[endpoint.zoneIndex\].id/);
  assert.equal((network.match(/name: '(?:host-blob|data-blob|data-table)'/g) ?? []).length, 3);
  assert.doesNotMatch(network, /Microsoft.Web\/serverFarms|serviceEndpoints:|virtualNetworkPeerings|publicIPAddresses|natGateways/);
});
test("workflow credentials stay separated and untrusted PRs never deploy", async () => {
  const deploy = await read(".github/workflows/deploy.yml");
  assert.match(deploy, /id-token: write/);
  assert.match(deploy, /sku: flexconsumption/);
  assert.doesNotMatch(deploy, /publish-profile|client-secret|functions-key|pull_request_target/);
  const validate = await read(".github/workflows/validate.yml");
  assert.doesNotMatch(validate, /secrets\.|id-token: write|Azure\/login|pull_request_target/);
  const infra = await read(".github/workflows/infra.yml");
  assert.match(infra, /what-if/); assert.match(infra, /environment: demo-infra/);
  assert.match(infra, /download-artifact@[a-f0-9]{40}/);
});
