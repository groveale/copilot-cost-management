import { beforeAll, describe, expect, it } from "vitest";
import { createLocalJWKSet, exportJWK, generateKeyPair, SignJWT } from "jose";
import { createAuthenticator } from "../src/auth.js";

const tenantId = "10000000-0000-4000-8000-000000000001";
const audience = "10000000-0000-4000-8000-000000000002";
const spaClientId = "10000000-0000-4000-8000-000000000003";
const objectId = "10000000-0000-4000-8000-000000000004";
let privateKey: CryptoKey;
let authenticate: ReturnType<typeof createAuthenticator>;
beforeAll(async () => {
  const pair = await generateKeyPair("RS256");
  privateKey = pair.privateKey;
  authenticate = createAuthenticator({ tenantId, audience, spaClientId },
    createLocalJWKSet({ keys: [{ ...await exportJWK(pair.publicKey), kid: "test", alg: "RS256" }] }));
});
async function token(overrides: Record<string, unknown> = {}, key = privateKey) {
  return new SignJWT({
    tid: tenantId, oid: objectId, ver: "2.0", azp: spaClientId, scp: "access_as_user",
    iss: `https://login.microsoftonline.com/${tenantId}/v2.0`, aud: audience,
    iat: Math.floor(Date.now() / 1000), exp: Math.floor(Date.now() / 1000) + 300, ...overrides,
  }).setProtectedHeader({ alg: "RS256", kid: "test" }).sign(key);
}
describe("delegated API token validation", () => {
  it("returns only validated stable IDs", async () => {
    expect(await authenticate(`Bearer ${await token()}`)).toEqual({ tenantId, objectId });
  });
  it("returns optional verified display hints without using email as the ownership key", async () => {
    expect(await authenticate(`Bearer ${await token({ name: "Example owner", preferred_username: "owner@example.test", email: "different@example.test" })}`))
      .toEqual({ tenantId, objectId, displayName: "Example owner", contactHint: "owner@example.test" });
  });
  it.each([null, "", "some-function-key", "Bearer", "Basic value"])("rejects missing/malformed bearer authentication: %s", async (header) => {
    await expect(authenticate(header)).rejects.toMatchObject({ status: 401 });
  });
  it.each([
    { aud: "https://graph.microsoft.com" }, { aud: spaClientId }, { iss: "https://attacker.test" },
    { tid: objectId }, { azp: objectId }, { scp: "" }, { scp: undefined },
    { idtyp: "app", roles: ["access_as_user"] }, { ver: "1.0" }, { oid: undefined },
    { oid: "not-an-object-id" }, { exp: 1 }, { exp: undefined }, { nbf: Math.floor(Date.now() / 1000) + 1000 },
  ])("rejects unauthorized token claims %j", async (overrides) => {
    await expect(authenticate(`Bearer ${await token(overrides)}`)).rejects.toMatchObject({ status: expect.any(Number) });
  });
  it("rejects a token signed with an untrusted key", async () => {
    const other = await generateKeyPair("RS256");
    await expect(authenticate(`Bearer ${await token({}, other.privateKey)}`)).rejects.toMatchObject({ status: 401 });
  });
});
