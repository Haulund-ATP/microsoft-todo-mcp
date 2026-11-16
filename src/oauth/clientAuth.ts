import { timingSafeEqual } from "node:crypto";
import { hashSecret } from "../storage/oauthRepo.js";
import type { OAuthClient } from "../storage/types.js";

/**
 * Client authentication at the token endpoint (RFC 6749 §2.3.1).
 *
 * MCP connectors register either as public clients ("none", PKCE only) or as
 * confidential clients ("client_secret_post" / "client_secret_basic" — the
 * MCP Python SDK's default, which Claude uses). A client registered with a
 * secret must present it on every token request; a public client is
 * identified by client_id alone, PKCE still protects the code exchange.
 */

export const TOKEN_ENDPOINT_AUTH_METHODS = ["none", "client_secret_post", "client_secret_basic"] as const;
export type TokenEndpointAuthMethod = (typeof TOKEN_ENDPOINT_AUTH_METHODS)[number];

export interface ClientCredentials {
  clientId: string;
  clientSecret?: string;
}

/** Form-urlencoding as used inside HTTP Basic client credentials ("+" is a space). */
function formDecode(value: string): string {
  return decodeURIComponent(value.replace(/\+/g, " "));
}

/**
 * Reads the client's credentials from an HTTP Basic Authorization header
 * and/or the request body. Returns an error string when they are missing or
 * contradict each other.
 */
export function readClientCredentials(
  authorizationHeader: string | undefined,
  body: { client_id?: unknown; client_secret?: unknown }
): ClientCredentials | { error: string } {
  const bodyId = typeof body.client_id === "string" && body.client_id !== "" ? body.client_id : undefined;
  const bodySecret = typeof body.client_secret === "string" && body.client_secret !== "" ? body.client_secret : undefined;

  if (authorizationHeader && /^basic\s+/i.test(authorizationHeader)) {
    let decoded: string;
    try {
      decoded = Buffer.from(authorizationHeader.replace(/^basic\s+/i, "").trim(), "base64").toString("utf8");
    } catch {
      return { error: "malformed_basic_auth" };
    }
    const colon = decoded.indexOf(":");
    if (colon <= 0) return { error: "malformed_basic_auth" };
    let id: string;
    let secret: string;
    try {
      id = formDecode(decoded.slice(0, colon));
      secret = formDecode(decoded.slice(colon + 1));
    } catch {
      return { error: "malformed_basic_auth" };
    }
    if (bodyId && bodyId !== id) return { error: "client_id_mismatch" };
    // RFC 6749: a client must not use more than one authentication method per request.
    if (bodySecret) return { error: "multiple_client_auth_methods" };
    return { clientId: id, clientSecret: secret };
  }

  if (!bodyId) return { error: "client_id_missing" };
  return bodySecret ? { clientId: bodyId, clientSecret: bodySecret } : { clientId: bodyId };
}

/**
 * True when the presented secret satisfies the client: a client registered
 * with a secret must present the matching one (constant-time comparison of
 * SHA-256 hashes); a public client needs none.
 */
export function clientSecretAccepted(client: Pick<OAuthClient, "clientSecretHash">, presentedSecret: string | undefined): boolean {
  if (!client.clientSecretHash) return true;
  if (!presentedSecret) return false;
  const expected = Buffer.from(client.clientSecretHash, "hex");
  const actual = Buffer.from(hashSecret(presentedSecret), "hex");
  return expected.length === actual.length && timingSafeEqual(expected, actual);
}
