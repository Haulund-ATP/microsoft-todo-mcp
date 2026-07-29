import { loadEnv } from "../config/env.js";

/**
 * OAuth Protected Resource Metadata (RFC 9728) and Authorization Server
 * Metadata (RFC 8414) / OIDC-discovery-shaped document, both served under
 * /.well-known/. MCP clients (ChatGPT, Claude) use these to discover the
 * authorization server responsible for a given resource (the /mcp
 * endpoint) without any out-of-band configuration.
 */

export function buildProtectedResourceMetadata() {
  const base = loadEnv().PUBLIC_BASE_URL.replace(/\/$/, "");
  return {
    resource: base,
    authorization_servers: [base],
    bearer_methods_supported: ["header"],
    resource_documentation: `${base}/docs`,
  };
}

export function buildAuthorizationServerMetadata() {
  const base = loadEnv().PUBLIC_BASE_URL.replace(/\/$/, "");
  return {
    issuer: base,
    authorization_endpoint: `${base}/oauth/authorize`,
    token_endpoint: `${base}/oauth/token`,
    revocation_endpoint: `${base}/oauth/revoke`,
    registration_endpoint: `${base}/oauth/register`,
    jwks_uri: `${base}/.well-known/jwks.json`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    code_challenge_methods_supported: ["S256"],
    token_endpoint_auth_methods_supported: ["none", "client_secret_basic"],
    scopes_supported: ["mcp"],
    service_documentation: `${base}/docs`,
    // OIDC-discovery-compatible fields, in case a client probes
    // /.well-known/openid-configuration instead of the OAuth-only variant.
    subject_types_supported: ["public"],
    id_token_signing_alg_values_supported: ["ES256"],
  };
}
