import { createRemoteJWKSet, jwtVerify, type JWTVerifyGetKey } from "jose";
import { ApiError, uuid } from "./model.js";
import type { Caller } from "./contracts.js";

export interface AuthSettings { tenantId: string; audience: string; spaClientId: string }
export function createAuthenticator(settings: AuthSettings, keys?: JWTVerifyGetKey) {
  const issuer = `https://login.microsoftonline.com/${settings.tenantId}/v2.0`;
  const resolver = keys ?? createRemoteJWKSet(new URL(
    `https://login.microsoftonline.com/${settings.tenantId}/discovery/v2.0/keys`,
  ));
  return async (authorization: string | null): Promise<Caller> => {
    if (!authorization || !/^Bearer \S+$/i.test(authorization)) {
      throw new ApiError(401, "unauthorized", "A delegated API bearer token is required.");
    }
    let payload;
    try {
      ({ payload } = await jwtVerify(authorization.slice(7), resolver, {
        issuer, audience: settings.audience, algorithms: ["RS256"], requiredClaims: ["exp", "iat", "tid", "oid"],
      }));
    } catch {
      throw new ApiError(401, "unauthorized", "The API access token is invalid or expired.");
    }
    if (payload.tid !== settings.tenantId || payload.ver !== "2.0" ||
        payload.azp !== settings.spaClientId || payload.idtyp === "app" ||
        typeof payload.scp !== "string" || !payload.scp.split(" ").includes("access_as_user") ||
        !uuid.safeParse(payload.oid).success) {
      throw new ApiError(403, "forbidden", "The token is not an authorized delegated API access token.");
    }
    const displayName = typeof payload.name === "string" ? payload.name.trim().slice(0, 200) : "";
    const contact = typeof payload.preferred_username === "string" ? payload.preferred_username : payload.email;
    const contactHint = typeof contact === "string" ? contact.trim().slice(0, 320) : "";
    return { tenantId: settings.tenantId, objectId: uuid.parse(payload.oid),
      ...(displayName ? { displayName } : {}), ...(contactHint ? { contactHint } : {}) };
  };
}
