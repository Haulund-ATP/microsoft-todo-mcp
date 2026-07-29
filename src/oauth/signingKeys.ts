import { importJWK, exportJWK, type JWK, type KeyLike } from "jose";
import { getSecretStore } from "../config/secrets.js";
import { loadEnv } from "../config/env.js";

/**
 * Signing key for the MCP-side OAuth 2.1 authorization server's JWTs
 * (access tokens; ID-token-like claims if ever needed). The private JWK is
 * generated once at provisioning time (see scripts/provision.ps1) and
 * stored as a Key Vault secret named by MCP_OAUTH_SIGNING_KEY_NAME. We use
 * an asymmetric key (ES256) so the public JWKS can be published at
 * /.well-known/jwks.json without exposing signing capability.
 */

let cachedPrivateKey: KeyLike | undefined;
let cachedPublicJwk: JWK | undefined;
let cachedKid: string | undefined;

async function loadPrivateJwk(): Promise<JWK> {
  const env = loadEnv();
  const store = getSecretStore();
  const raw = await store.getSecret(env.MCP_OAUTH_SIGNING_KEY_NAME);
  const jwk = JSON.parse(raw) as JWK & { kid?: string };
  if (!jwk.kid) {
    throw new Error("Signing key JWK in Key Vault must include a stable 'kid'.");
  }
  return jwk;
}

export async function getSigningKey(): Promise<{ key: KeyLike; kid: string }> {
  if (cachedPrivateKey && cachedKid) return { key: cachedPrivateKey, kid: cachedKid };
  const jwk = await loadPrivateJwk();
  cachedPrivateKey = (await importJWK(jwk, "ES256")) as KeyLike;
  cachedKid = jwk.kid as string;

  const { d: _d, ...publicPart } = jwk;
  cachedPublicJwk = { ...publicPart, kid: jwk.kid, alg: "ES256", use: "sig" };

  return { key: cachedPrivateKey, kid: cachedKid };
}

export async function getPublicJwks(): Promise<{ keys: JWK[] }> {
  if (!cachedPublicJwk) await getSigningKey();
  return { keys: cachedPublicJwk ? [cachedPublicJwk] : [] };
}

/** Test-only helper to bypass Key Vault; not used in production code paths. */
export function _setSigningKeyForTests(key: KeyLike, kid: string, publicJwk: JWK): void {
  cachedPrivateKey = key;
  cachedKid = kid;
  cachedPublicJwk = publicJwk;
}

export { exportJWK };
