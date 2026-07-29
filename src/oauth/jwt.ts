import { SignJWT, jwtVerify, type JWTPayload } from "jose";
import { loadEnv } from "../config/env.js";
import { getSigningKey } from "./signingKeys.js";

export interface McpAccessTokenClaims extends JWTPayload {
  sub: string; // owner claim value (see ADMIN_OWNER_CLAIM_VALUE)
  client_id: string;
  scope: string;
  resource: string;
}

export async function signAccessToken(claims: {
  subject: string;
  clientId: string;
  scope: string;
  resource: string;
}): Promise<string> {
  const env = loadEnv();
  const { key, kid } = await getSigningKey();
  return new SignJWT({ client_id: claims.clientId, scope: claims.scope, resource: claims.resource })
    .setProtectedHeader({ alg: "ES256", kid })
    .setIssuer(env.MCP_OAUTH_ISSUER)
    .setAudience(claims.resource)
    .setSubject(claims.subject)
    .setIssuedAt()
    .setExpirationTime(`${env.MCP_OAUTH_ACCESS_TOKEN_TTL_SECONDS}s`)
    .sign(key);
}

export async function verifyAccessToken(token: string, expectedResource: string): Promise<McpAccessTokenClaims> {
  const env = loadEnv();
  const { key } = await getSigningKey();
  const { payload } = await jwtVerify(token, key, {
    issuer: env.MCP_OAUTH_ISSUER,
    audience: expectedResource,
  });
  return payload as McpAccessTokenClaims;
}
