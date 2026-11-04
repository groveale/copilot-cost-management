import { beforeEach, expect, it, vi } from "vitest";
import { CertificateObo, ReportingInteractionRequired, validatedClaims } from "../src/reporting-auth.js";
vi.mock("node:crypto", async (original) => ({
  ...await original<typeof import("node:crypto")>(),
  X509Certificate: class {
    publicKey = { asymmetricKeyType: "rsa" }; validFrom = "2026-01-01T00:00:00Z"; validTo = "2099-01-01T00:00:00Z";
    raw = Buffer.from("synthetic-public-certificate");
  },
}));
const id = "10000000-0000-4000-8000-000000000001";
const certId = `https://example.vault.azure.net/certificates/reporting/${"a".repeat(32)}`;
let calls: { url: string; init?: RequestInit }[];
let tokenError: object | null;
let certificateKey: string;
let fetcher: typeof fetch;
beforeEach(() => {
  calls = []; tokenError = null;
  certificateKey = `https://example.vault.azure.net/keys/reporting/${"b".repeat(32)}`;
  fetcher = vi.fn(async (input, init) => {
    const url = String(input); calls.push({ url, init });
    const value = url.includes("/certificates/") ? { cer: Buffer.from("synthetic-public-certificate").toString("base64"), kid: certificateKey } :
      url.includes("/sign?") ? { value: "synthetic_signature" } : tokenError ?? { access_token: "synthetic-delegated-access" };
    return new Response(JSON.stringify(value), { status: tokenError && url.endsWith("/token") ? 400 : 200 });
  }) as typeof fetch;
});
const auth = (uri = certId) => new CertificateObo(id, id, uri, "https://example.sharepoint.com", "Bearer synthetic-api-assertion",
  { getToken: async (scope) => { expect(scope).toBe("https://vault.azure.net/.default"); return { token: "synthetic-vault-access", expiresOnTimestamp: Date.now() + 30000 }; } }, fetcher);
it("uses delegated OBO with a remote certificate-key signature, never private-key/secret retrieval", async () => {
  const obo = auth();
  expect(await obo.token("sharepoint")).toBe("synthetic-delegated-access");
  expect(await obo.token("sharepoint")).toBe("synthetic-delegated-access");
  expect(calls).toHaveLength(3);
  expect(calls.some(({ url }) => url.includes("/secrets/"))).toBe(false);
  const sign = JSON.parse(String(calls[1].init!.body));
  expect(sign).toEqual({ alg: "RS256", value: expect.stringMatching(/^[A-Za-z0-9_-]{43}$/) });
  const body = calls[2].init!.body as URLSearchParams;
  expect(body.get("requested_token_use")).toBe("on_behalf_of");
  expect(body.get("assertion")).toBe("synthetic-api-assertion");
  expect(body.get("scope")).toBe("https://example.sharepoint.com/.default");
  expect(body.has("client_secret")).toBe(false);
  const assertion = body.get("client_assertion")!.split(".");
  expect(JSON.parse(Buffer.from(assertion[0], "base64url").toString())).toMatchObject({ alg: "RS256", x5t: expect.any(String) });
  expect(JSON.parse(Buffer.from(assertion[1], "base64url").toString())).toMatchObject({ iss: id, sub: id, aud: `https://login.microsoftonline.com/${id}/oauth2/v2.0/token` });
  expect(calls.every(({ init }) => init!.redirect === "error")).toBe(true);
});
it("uses separate Graph delegated scope only for server-authoritative user identity resolution", async () => {
  await auth().token("graph");
  expect((calls.at(-1)!.init!.body as URLSearchParams).get("scope")).toBe("https://graph.microsoft.com/.default");
});
it("rejects untrusted certificate/key destinations before forwarding identity tokens", async () => {
  await expect(auth("https://untrusted.example/certificates/test/" + "a".repeat(32)).token("sharepoint")).rejects.toMatchObject({ code: "reporting-not-configured" });
  expect(calls).toHaveLength(0);
  certificateKey = "https://untrusted.example/keys/sign/" + "b".repeat(32);
  await expect(auth().token("sharepoint")).rejects.toThrow("signing-key destination");
  expect(calls).toHaveLength(1);
});
it("surfaces only validated claims in a safe API challenge, without error descriptions or tokens", async () => {
  const claims = JSON.stringify({ access_token: { acrs: { essential: true, value: "c1" } } });
  tokenError = { error: "interaction_required", claims, error_description: "do not leak tenant internals" };
  const error = await auth().token("sharepoint").catch((value: unknown) => value) as ReportingInteractionRequired;
  expect(error).toBeInstanceOf(ReportingInteractionRequired);
  expect(error.code).toBe("reporting-interaction-required");
  expect(error.challenge).toContain(`claims="${Buffer.from(claims).toString("base64")}"`);
  expect(error.challenge).not.toContain("do not leak");
  for (const raw of ["bad\r\nX-Injected: 1", JSON.stringify({ access_token: {}, token: "secret" }), JSON.stringify({ access_token: [] }), "x".repeat(9000)]) {
    expect(validatedClaims(raw)).toBeUndefined();
  }
});
