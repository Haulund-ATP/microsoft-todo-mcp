import { beforeAll, describe, expect, it } from "vitest";
import { exportJWK, generateKeyPair } from "jose";
void generateKeyPair;

process.env.NODE_ENV = "test";
process.env.PUBLIC_BASE_URL = "https://todo.example.test";
process.env.AZURE_TENANT_ID = "test-tenant";
process.env.AZURE_CLIENT_ID = "test-client";
process.env.AZURE_KEY_VAULT_URI = "https://test-vault.vault.azure.net/";
process.env.AZURE_STORAGE_CONNECTION_STRING = "UseDevelopmentStorage=true";
process.env.MCP_OAUTH_SIGNING_KEY_NAME = "mcp-oauth-signing-key";
process.env.MCP_OAUTH_ISSUER = "https://todo.example.test";
process.env.ADMIN_OWNER_CLAIM_VALUE = "test-owner-oid";
process.env.SESSION_COOKIE_SECRET = "a".repeat(32);
process.env.CSRF_COOKIE_SECRET = "b".repeat(32);

describe("access token sign/verify", () => {
  let signAccessToken: typeof import("../../src/oauth/jwt.js").signAccessToken;
  let verifyAccessToken: typeof import("../../src/oauth/jwt.js").verifyAccessToken;

  beforeAll(async () => {
    const { generateKeyPair: genKeyPair } = await import("jose");
    const { privateKey } = await genKeyPair("ES256", { extractable: true });
    const jwk = await exportJWK(privateKey);
    const kid = "test-kid";

    const { _setSigningKeyForTests } = await import("../../src/oauth/signingKeys.js");
    _setSigningKeyForTests(privateKey, kid, { ...jwk, kid, alg: "ES256", use: "sig" });

    const jwtModule = await import("../../src/oauth/jwt.js");
    signAccessToken = jwtModule.signAccessToken;
    verifyAccessToken = jwtModule.verifyAccessToken;
  });

  it("signs and verifies a token with matching issuer/audience/subject", async () => {
    const resource = "https://todo.example.test";
    const token = await signAccessToken({
      subject: "owner-claim-value",
      clientId: "client-123",
      scope: "mcp",
      resource,
    });
    const claims = await verifyAccessToken(token, resource);
    expect(claims.sub).toBe("owner-claim-value");
    expect(claims.client_id).toBe("client-123");
    expect(claims.scope).toBe("mcp");
    expect(claims.aud).toBe(resource);
  });

  it("rejects verification against a different resource (audience)", async () => {
    const token = await signAccessToken({
      subject: "owner-claim-value",
      clientId: "client-123",
      scope: "mcp",
      resource: "https://todo.example.test",
    });
    await expect(verifyAccessToken(token, "https://other.example.test")).rejects.toThrow();
  });

  it("rejects a tampered token", async () => {
    const token = await signAccessToken({
      subject: "owner-claim-value",
      clientId: "client-123",
      scope: "mcp",
      resource: "https://todo.example.test",
    });
    const tampered = token.slice(0, -2) + (token.slice(-2) === "AA" ? "BB" : "AA");
    await expect(verifyAccessToken(tampered, "https://todo.example.test")).rejects.toThrow();
  });
});
