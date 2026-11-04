import { createHash, randomUUID, X509Certificate } from "node:crypto";
import type { TokenCredential } from "@azure/identity";
import { ApiError, uuid } from "./model.js";

export class ReportingInteractionRequired extends ApiError {
  readonly challenge: string;
  constructor(tenantId: string, claims?: string) {
    super(401, "reporting-interaction-required", "Renew API sign-in to satisfy SharePoint permission-service consent or Conditional Access.");
    this.challenge = `Bearer authorization_uri="https://login.microsoftonline.com/${uuid.parse(tenantId)}/oauth2/v2.0/authorize", error="insufficient_claims"` +
      (claims ? `, claims="${claims}"` : "");
  }
}
export function validatedClaims(raw: unknown): string | undefined {
  if (typeof raw !== "string" || raw.length > 8192) return undefined;
  try {
    const text = raw.trim().startsWith("{") ? raw : Buffer.from(raw, "base64").toString("utf8");
    const value: unknown = JSON.parse(text);
    if (!value || typeof value !== "object" || Array.isArray(value) ||
        !("access_token" in value) || !value.access_token || typeof value.access_token !== "object" ||
        Array.isArray(value.access_token) || Object.keys(value).some((key) => key !== "access_token")) return undefined;
    return Buffer.from(JSON.stringify(value)).toString("base64");
  } catch { return undefined; }
}
export interface ReportingAuth {
  token(resource: "sharepoint" | "graph"): Promise<string>;
}
// Key Vault signs a digest using the nonexportable certificate key. No private key,
// client secret, refresh token or downstream token is persisted.
export class CertificateObo implements ReportingAuth {
  private readonly tokens = new Map<string, string>();
  constructor(private readonly tenantId: string, private readonly clientId: string,
    private readonly certificateId: string, private readonly sharepointOrigin: string,
    private readonly authorization: string | null, private readonly credential: TokenCredential,
    private readonly fetcher: typeof fetch = fetch) {}
  async token(resource: "sharepoint" | "graph"): Promise<string> {
    const cached = this.tokens.get(resource);
    if (cached) return cached;
    if (!this.authorization?.startsWith("Bearer ")) throw new ApiError(401, "unauthorized", "An authenticated delegated API token is required.");
    const certificateUrl = new URL(this.certificateId);
    if (certificateUrl.protocol !== "https:" || !/^[a-z0-9-]+\.vault\.azure\.net$/.test(certificateUrl.hostname) ||
        !/^\/certificates\/[a-zA-Z0-9-]+\/[a-f0-9]{32}$/.test(certificateUrl.pathname) ||
        certificateUrl.search || certificateUrl.hash || certificateUrl.port || certificateUrl.username || certificateUrl.password) {
      throw new ApiError(409, "reporting-not-configured", "Configure a versioned Azure Key Vault certificate URI.");
    }
    const vaultToken = await this.credential.getToken("https://vault.azure.net/.default");
    if (!vaultToken) throw new Error("Key Vault identity token is unavailable.");
    const vaultHeaders = { Authorization: `Bearer ${vaultToken.token}` };
    const certResponse = await this.fetcher(`${certificateUrl.href}?api-version=7.4`, {
      headers: vaultHeaders, redirect: "error", signal: AbortSignal.timeout(30000),
    });
    if (!certResponse.ok) throw new ApiError(503, "reporting-certificate", "The permission-service certificate is unavailable. Check certificate version and scoped Key Vault access.");
    const cert = await certResponse.json() as { cer?: string; kid?: string };
    if (typeof cert.cer !== "string" || typeof cert.kid !== "string") throw new Error("Invalid Key Vault public certificate.");
    const keyUrl = new URL(cert.kid);
    if (keyUrl.origin !== certificateUrl.origin || !/^\/keys\/[a-zA-Z0-9-]+\/[a-f0-9]{32}$/.test(keyUrl.pathname) ||
        keyUrl.search || keyUrl.hash || keyUrl.username || keyUrl.password) throw new Error("Invalid certificate signing-key destination.");
    const certificate = new X509Certificate(Buffer.from(cert.cer, "base64"));
    if (certificate.publicKey.asymmetricKeyType !== "rsa" || Date.parse(certificate.validTo) <= Date.now() ||
        Date.parse(certificate.validFrom) > Date.now()) throw new ApiError(503, "reporting-certificate", "The permission-service RSA certificate is not currently valid.");
    const endpoint = `https://login.microsoftonline.com/${uuid.parse(this.tenantId)}/oauth2/v2.0/token`;
    const now = Math.floor(Date.now() / 1000);
    const encode = (value: object) => Buffer.from(JSON.stringify(value)).toString("base64url");
    const input = `${encode({ alg: "RS256", typ: "JWT", x5t: createHash("sha1").update(certificate.raw).digest("base64url") })}.${encode({
      aud: endpoint, iss: this.clientId, sub: this.clientId, jti: randomUUID(), nbf: now - 30, exp: now + 300,
    })}`;
    const signed = await this.fetcher(`${keyUrl.href}/sign?api-version=7.4`, {
      method: "POST", headers: { ...vaultHeaders, "Content-Type": "application/json" }, redirect: "error",
      signal: AbortSignal.timeout(30000), body: JSON.stringify({ alg: "RS256", value: createHash("sha256").update(input).digest("base64url") }),
    });
    if (!signed.ok) throw new ApiError(503, "reporting-certificate", "Key Vault certificate signing failed. Check the restricted sign role.");
    const signature = await signed.json() as { value?: string };
    if (typeof signature.value !== "string" || !/^[A-Za-z0-9_-]+$/.test(signature.value)) throw new Error("Invalid Key Vault signature.");
    const response = await this.fetcher(endpoint, {
      method: "POST", redirect: "error", signal: AbortSignal.timeout(30000),
      body: new URLSearchParams({
        client_id: this.clientId, client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        client_assertion: `${input}.${signature.value}`, grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        requested_token_use: "on_behalf_of", assertion: this.authorization.slice(7),
        scope: `${resource === "graph" ? "https://graph.microsoft.com" : this.sharepointOrigin}/.default`,
      }),
    });
    const result = await response.json() as { access_token?: string; error?: string; claims?: string };
    if (!response.ok) {
      if (["interaction_required", "invalid_grant", "consent_required"].includes(result.error ?? "") || validatedClaims(result.claims)) {
        throw new ReportingInteractionRequired(this.tenantId, validatedClaims(result.claims));
      }
      throw new ApiError(503, "reporting-obo", "Delegated permission-service token exchange failed. Check certificate registration and downstream delegated consent.");
    }
    if (typeof result.access_token !== "string" || !result.access_token) throw new Error("Delegated token was not returned.");
    this.tokens.set(resource, result.access_token);
    return result.access_token;
  }
}
