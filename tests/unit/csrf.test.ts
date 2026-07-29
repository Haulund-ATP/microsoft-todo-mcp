import { describe, it, expect, beforeAll } from "vitest";
import type { Request, Response } from "express";

// loadEnv() validates the full env schema even though csrf.ts only reads
// CSRF_COOKIE_SECRET — fill in the rest with harmless test values so the
// first call in this test file doesn't throw on unrelated missing vars.
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

let issueCsrfToken: typeof import("../../src/admin/csrf.js").issueCsrfToken;
let requireCsrf: typeof import("../../src/admin/csrf.js").requireCsrf;

beforeAll(async () => {
  ({ issueCsrfToken, requireCsrf } = await import("../../src/admin/csrf.js"));
});

function fakeRes(): Response {
  const cookies: Record<string, string> = {};
  return {
    cookie: (name: string, value: string) => {
      cookies[name] = value;
    },
    status: () => ({ json: () => undefined }),
    _cookies: cookies,
  } as unknown as Response;
}

function fakeReq(cookieValue?: string): Request {
  return { cookies: cookieValue ? { mth_csrf: cookieValue } : {} } as unknown as Request;
}

describe("issueCsrfToken", () => {
  it("mints a fresh token when the request has no existing cookie", () => {
    const res = fakeRes();
    const token = issueCsrfToken(fakeReq(), res);
    expect(token).toMatch(/^[0-9a-f]+\.[0-9a-f]+$/);
    expect((res as unknown as { _cookies: Record<string, string> })._cookies.mth_csrf).toBe(token);
  });

  it("reuses an already-valid cookie instead of minting a new one — fixes concurrent page loads invalidating each other", () => {
    const first = issueCsrfToken(fakeReq(), fakeRes());
    const second = issueCsrfToken(fakeReq(first), fakeRes());
    expect(second).toBe(first);
  });

  it("mints a new token if the existing cookie is invalid/tampered", () => {
    const token = issueCsrfToken(fakeReq("garbage.notasignature"), fakeRes());
    expect(token).not.toBe("garbage.notasignature");
  });
});

describe("requireCsrf", () => {
  function run(cookieToken: string | undefined, bodyToken: string | undefined): number {
    let status = 200;
    const req = { cookies: cookieToken ? { mth_csrf: cookieToken } : {}, body: { _csrf: bodyToken } } as unknown as Request;
    const res = {
      status: (code: number) => {
        status = code;
        return { json: () => undefined };
      },
    } as unknown as Response;
    let nextCalled = false;
    requireCsrf(req, res, () => {
      nextCalled = true;
    });
    return nextCalled ? 200 : status;
  }

  it("passes when the two most-recently-issued tokens (same cookie) match", () => {
    const token = issueCsrfToken(fakeReq(), fakeRes());
    expect(run(token, token)).toBe(200);
  });

  it("rejects with 403 when cookie and body tokens differ", () => {
    const token = issueCsrfToken(fakeReq(), fakeRes());
    expect(run(token, `${token}x`)).toBe(403);
  });

  it("rejects with 403 when the cookie token is missing", () => {
    const token = issueCsrfToken(fakeReq(), fakeRes());
    expect(run(undefined, token)).toBe(403);
  });
});
