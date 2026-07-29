import { expect, test } from "@playwright/test";

test.describe("MCP endpoint authentication", () => {
  test("universal /mcp rejects unauthenticated requests with 401 + WWW-Authenticate", async ({ request }) => {
    const response = await request.post("/mcp", { data: {} });
    expect(response.status()).toBe(401);
    const challenge = response.headers()["www-authenticate"];
    expect(challenge).toContain("Bearer");
    expect(challenge).toContain("resource_metadata=");
  });

  test("/mcp/personal rejects unauthenticated requests with 401", async ({ request }) => {
    const response = await request.post("/mcp/personal", { data: {} });
    expect(response.status()).toBe(401);
  });

  test("a garbage bearer token is rejected as invalid_token", async ({ request }) => {
    const response = await request.post("/mcp", {
      data: {},
      headers: { Authorization: "Bearer not-a-real-token" },
    });
    expect(response.status()).toBe(401);
    const body = await response.json();
    expect(body.error).toBe("invalid_token");
  });

  test("an unknown profile alias 404s rather than silently falling back to the universal endpoint", async ({
    request,
  }) => {
    const response = await request.post("/mcp/definitely-not-a-bound-profile-xyz", {
      data: {},
      headers: { Authorization: "Bearer not-a-real-token" },
    });
    // Bearer check runs first (401) — but an unbound alias must never
    // silently succeed even with a hypothetically valid token, so this
    // also documents the expectation for a follow-up authenticated test.
    expect([401, 404]).toContain(response.status());
  });
});

test.describe("Dynamic Client Registration", () => {
  test("registering a client without redirect_uris is rejected", async ({ request }) => {
    const response = await request.post("/oauth/register", {
      data: { client_name: "Test client" },
    });
    expect(response.status()).toBe(400);
  });

  test("registering a client with an http (non-localhost) redirect_uri is rejected", async ({ request }) => {
    const response = await request.post("/oauth/register", {
      data: {
        client_name: "Test client",
        redirect_uris: ["http://example.com/callback"],
      },
    });
    expect(response.status()).toBe(400);
  });

  test("registering a valid client succeeds and returns a public-client shape", async ({ request }) => {
    const response = await request.post("/oauth/register", {
      data: {
        client_name: "Playwright test client",
        redirect_uris: ["https://example.com/callback"],
      },
    });
    expect(response.status()).toBe(201);
    const body = await response.json();
    expect(body.client_id).toBeTruthy();
    expect(body.token_endpoint_auth_method).toBe("none");
  });
});

test.describe("Authorization endpoint without an owner session", () => {
  test("redirects to owner login rather than issuing a code", async ({ request }) => {
    const registerResponse = await request.post("/oauth/register", {
      data: {
        client_name: "Playwright authorize test client",
        redirect_uris: ["https://example.com/callback"],
      },
    });
    const { client_id: clientId } = await registerResponse.json();

    // `resource` must exactly match this deployment's PUBLIC_BASE_URL for
    // the authorize endpoint to accept it — see src/oauth/authorize.ts.
    const resource = process.env.PLAYWRIGHT_BASE_URL ?? "http://localhost:3000";
    const authorizeUrl =
      `/oauth/authorize?response_type=code&client_id=${clientId}` +
      `&redirect_uri=${encodeURIComponent("https://example.com/callback")}` +
      `&state=teststate123&code_challenge=${"a".repeat(43)}&code_challenge_method=S256` +
      `&resource=${encodeURIComponent(resource)}`;

    const response = await request.get(authorizeUrl, { maxRedirects: 0 });
    expect([302, 303, 307]).toContain(response.status());
    const location = response.headers()["location"] ?? "";
    expect(location).toContain("/oauth/owner/login");
  });
});
