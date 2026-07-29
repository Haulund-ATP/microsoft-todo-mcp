import { expect, test } from "@playwright/test";

test.describe("OAuth discovery documents", () => {
  test("Protected Resource Metadata is well-formed", async ({ request }) => {
    const response = await request.get("/.well-known/oauth-protected-resource");
    expect(response.ok()).toBeTruthy();
    const body = await response.json();
    expect(body.resource).toBeTruthy();
    expect(Array.isArray(body.authorization_servers)).toBe(true);
    expect(body.authorization_servers).toContain(body.resource);
  });

  test("Authorization Server Metadata exposes all required endpoints", async ({ request }) => {
    const response = await request.get("/.well-known/oauth-authorization-server");
    expect(response.ok()).toBeTruthy();
    const body = await response.json();
    for (const field of [
      "issuer",
      "authorization_endpoint",
      "token_endpoint",
      "revocation_endpoint",
      "registration_endpoint",
      "jwks_uri",
    ]) {
      expect(body[field], `missing ${field}`).toBeTruthy();
      expect(String(body[field])).toContain(body.issuer);
    }
    expect(body.code_challenge_methods_supported).toContain("S256");
    expect(body.grant_types_supported).toEqual(expect.arrayContaining(["authorization_code", "refresh_token"]));
  });

  test("OIDC-discovery alias returns the same document shape", async ({ request }) => {
    const response = await request.get("/.well-known/openid-configuration");
    expect(response.ok()).toBeTruthy();
    const body = await response.json();
    expect(body.issuer).toBeTruthy();
  });

  test("JWKS document exposes at least one public signing key, no private material", async ({ request }) => {
    const response = await request.get("/.well-known/jwks.json");
    expect(response.ok()).toBeTruthy();
    const body = await response.json();
    expect(Array.isArray(body.keys)).toBe(true);
    expect(body.keys.length).toBeGreaterThan(0);
    for (const key of body.keys) {
      expect(key.d, "private key component 'd' must never be published in JWKS").toBeUndefined();
      expect(key.kid).toBeTruthy();
    }
  });
});
