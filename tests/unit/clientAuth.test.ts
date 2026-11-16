import { describe, expect, it, vi, beforeEach } from "vitest";

process.env.NODE_ENV ??= "test";
process.env.PUBLIC_BASE_URL ??= "https://example.test";
process.env.AZURE_TENANT_ID ??= "test-tenant";
process.env.AZURE_CLIENT_ID ??= "test-client";
process.env.AZURE_KEY_VAULT_URI ??= "https://example-kv.test";
process.env.AZURE_STORAGE_CONNECTION_STRING ??= "UseDevelopmentStorage=true";
process.env.MCP_OAUTH_SIGNING_KEY_NAME ??= "test-signing-key";
process.env.MCP_OAUTH_ISSUER ??= "https://example.test";
process.env.ADMIN_OWNER_CLAIM_VALUE ??= "test-owner";
process.env.SESSION_COOKIE_SECRET ??= "test-session-secret-0123456789ab";
process.env.CSRF_COOKIE_SECRET ??= "test-csrf-secret-0123456789abcdef";

const clients = new Map<string, Record<string, unknown>>();
vi.mock("../../src/storage/oauthRepo.js", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../src/storage/oauthRepo.js")>();
  return {
    ...actual,
    OAuthClientsRepo: class {
      async register(c: Record<string, unknown>) {
        clients.set(c.clientId as string, c);
      }
      async get(id: string) {
        return clients.get(id);
      }
    },
    OAuthCodesRepo: class {
      async consume() {
        return { reason: "unknown" };
      }
    },
  };
});

const { readClientCredentials, clientSecretAccepted } = await import("../../src/oauth/clientAuth.js");
const { hashSecret } = await import("../../src/storage/oauthRepo.js");
const { registerRouter } = await import("../../src/oauth/register.js");
const { tokenRouter } = await import("../../src/oauth/token.js");
const { buildAuthorizationServerMetadata } = await import("../../src/oauth/metadata.js");

type Handler = (req: unknown, res: unknown) => Promise<void>;
function routeHandler(router: { stack: Array<{ route?: { path: string; stack: Array<{ handle: Handler }> } }> }, path: string): Handler {
  const layer = router.stack.find((l) => l.route?.path === path)!;
  return layer.route!.stack[layer.route!.stack.length - 1]!.handle;
}

function fakeRes() {
  const res = {
    statusCode: 200,
    body: undefined as unknown,
    status(code: number) {
      res.statusCode = code;
      return res;
    },
    json(b: unknown) {
      res.body = b;
      return res;
    },
    end() {
      return res;
    },
  };
  return res;
}

function fakeReq(body: Record<string, unknown>, authorization?: string) {
  return { body, correlationId: "corr", header: (name: string) => (name.toLowerCase() === "authorization" ? authorization : undefined) };
}

const basic = (id: string, secret: string) => `Basic ${Buffer.from(`${encodeURIComponent(id)}:${encodeURIComponent(secret)}`).toString("base64")}`;

describe("readClientCredentials", () => {
  it("reads client_id (and optional client_secret) from the body", () => {
    expect(readClientCredentials(undefined, { client_id: "a" })).toEqual({ clientId: "a" });
    expect(readClientCredentials(undefined, { client_id: "a", client_secret: "s" })).toEqual({ clientId: "a", clientSecret: "s" });
  });

  it("reads HTTP Basic credentials, form-decoded", () => {
    expect(readClientCredentials(basic("my id", "p:a+ss"), {})).toEqual({ clientId: "my id", clientSecret: "p:a+ss" });
    expect(readClientCredentials(basic("a", "s"), { client_id: "a" })).toEqual({ clientId: "a", clientSecret: "s" });
  });

  it("rejects missing, contradictory or doubled credentials", () => {
    expect(readClientCredentials(undefined, {})).toEqual({ error: "client_id_missing" });
    expect(readClientCredentials(basic("a", "s"), { client_id: "b" })).toEqual({ error: "client_id_mismatch" });
    expect(readClientCredentials(basic("a", "s"), { client_secret: "s" })).toEqual({ error: "multiple_client_auth_methods" });
    expect(readClientCredentials("Basic " + Buffer.from("nocolon").toString("base64"), {})).toEqual({ error: "malformed_basic_auth" });
  });
});

describe("clientSecretAccepted", () => {
  it("lets public clients through without a secret", () => {
    expect(clientSecretAccepted({}, undefined)).toBe(true);
  });

  it("requires the exact secret for confidential clients", () => {
    const client = { clientSecretHash: hashSecret("right") };
    expect(clientSecretAccepted(client, "right")).toBe(true);
    expect(clientSecretAccepted(client, "wrong")).toBe(false);
    expect(clientSecretAccepted(client, undefined)).toBe(false);
  });
});

describe("dynamic client registration", () => {
  const register = routeHandler(registerRouter as never, "/oauth/register");
  const claudeBody = {
    client_name: "Claude",
    redirect_uris: ["https://claude.ai/api/mcp/auth_callback"],
    grant_types: ["authorization_code", "refresh_token"],
    response_types: ["code"],
    scope: "mcp",
  };

  beforeEach(() => clients.clear());

  it("registers Claude's client_secret_post request with a one-time secret, storing only its hash", async () => {
    const res = fakeRes();
    await register(fakeReq({ ...claudeBody, token_endpoint_auth_method: "client_secret_post" }), res);
    expect(res.statusCode).toBe(201);
    const body = res.body as Record<string, unknown>;
    expect(body.token_endpoint_auth_method).toBe("client_secret_post");
    expect(typeof body.client_secret).toBe("string");
    expect(body.client_secret_expires_at).toBe(0);
    const stored = clients.get(body.client_id as string)!;
    expect(stored.clientSecretHash).toBe(hashSecret(body.client_secret as string));
    expect(JSON.stringify(stored)).not.toContain(body.client_secret as string);
  });

  it("keeps public clients secret-free", async () => {
    const res = fakeRes();
    await register(fakeReq({ ...claudeBody, token_endpoint_auth_method: "none" }), res);
    const body = res.body as Record<string, unknown>;
    expect(res.statusCode).toBe(201);
    expect(body.client_secret).toBeUndefined();
    expect(clients.get(body.client_id as string)!.clientSecretHash).toBeUndefined();
  });

  it("still rejects unknown auth methods", async () => {
    const res = fakeRes();
    await register(fakeReq({ ...claudeBody, token_endpoint_auth_method: "private_key_jwt" }), res);
    expect(res.statusCode).toBe(400);
  });

  it("advertises every supported method in the authorization server metadata", () => {
    expect(buildAuthorizationServerMetadata().token_endpoint_auth_methods_supported).toEqual(["none", "client_secret_post", "client_secret_basic"]);
  });
});

describe("token endpoint client authentication", () => {
  const token = routeHandler(tokenRouter as never, "/oauth/token");
  const grant = { grant_type: "authorization_code", code: "c", redirect_uri: "https://claude.ai/api/mcp/auth_callback", code_verifier: "v".repeat(43) };

  beforeEach(() => {
    clients.clear();
    clients.set("conf", { clientId: "conf", clientSecretHash: hashSecret("s3cret") });
    clients.set("pub", { clientId: "pub" });
  });

  it("rejects a confidential client without or with a wrong secret (401 invalid_client)", async () => {
    for (const req of [fakeReq({ ...grant, client_id: "conf" }), fakeReq({ ...grant, client_id: "conf", client_secret: "nope" }), fakeReq(grant, basic("conf", "nope"))]) {
      const res = fakeRes();
      await token(req, res);
      expect(res.statusCode).toBe(401);
      expect(res.body).toEqual({ error: "invalid_client" });
    }
  });

  it("accepts the right secret via body or Basic and moves on to the grant", async () => {
    for (const req of [fakeReq({ ...grant, client_id: "conf", client_secret: "s3cret" }), fakeReq(grant, basic("conf", "s3cret"))]) {
      const res = fakeRes();
      await token(req, res);
      // The fake code store knows no codes, so passing client auth ends at invalid_grant.
      expect(res.statusCode).toBe(400);
      expect(res.body).toMatchObject({ error: "invalid_grant" });
    }
  });

  it("keeps public clients working with client_id only", async () => {
    const res = fakeRes();
    await token(fakeReq({ ...grant, client_id: "pub" }), res);
    expect(res.body).toMatchObject({ error: "invalid_grant" });
  });

  it("rejects requests that identify no client", async () => {
    const res = fakeRes();
    await token(fakeReq(grant), res);
    expect(res.statusCode).toBe(400);
    expect(res.body).toEqual({ error: "invalid_client" });
  });
});
