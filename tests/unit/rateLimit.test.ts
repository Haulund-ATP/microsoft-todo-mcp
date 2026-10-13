import { afterAll, beforeAll, describe, expect, it } from "vitest";
import type { Server } from "node:http";
import type { AddressInfo } from "node:net";
import { _resetEnvCacheForTests } from "../../src/config/env.js";

/**
 * Exercises the real app's limiter wiring. Every request below is rejected
 * by validation or the bearer check before touching storage, so only the
 * 429-or-not outcome matters. Limiter counters are per process, so each
 * test uses its own client address.
 */

let server: Server;
let base: string;

beforeAll(async () => {
  Object.assign(process.env, {
    NODE_ENV: "test",
    PUBLIC_BASE_URL: "https://todo.example",
    AZURE_TENANT_ID: "t",
    AZURE_CLIENT_ID: "c",
    AZURE_KEY_VAULT_URI: "https://kv.example",
    AZURE_STORAGE_ACCOUNT_NAME: "st",
    MCP_OAUTH_SIGNING_KEY_NAME: "k",
    MCP_OAUTH_ISSUER: "https://todo.example",
    ADMIN_OWNER_CLAIM_VALUE: "owner",
    SESSION_COOKIE_SECRET: "x".repeat(32),
    CSRF_COOKIE_SECRET: "y".repeat(32),
    LOG_LEVEL: "error",
  });
  _resetEnvCacheForTests();
  const { createApp } = await import("../../src/http/app.js");
  server = createApp().listen(0);
  base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`;
  // Importing the whole app is slow when the full suite runs in parallel.
}, 60_000);

afterAll(() => server?.close());

async function statuses(path: string, client: string, count: number, init: RequestInit = {}) {
  const out: number[] = [];
  for (let i = 0; i < count; i++) {
    const res = await fetch(base + path, { ...init, headers: { "x-envoy-external-address": client, ...(init.headers ?? {}) } });
    out.push(res.status);
  }
  return out;
}

describe("rate limit wiring", { timeout: 30_000 }, () => {
  it("does not apply the 20/min sign-in limit to /mcp (Claude Code's startup burst)", async () => {
    const result = await statuses("/mcp/nadia", "10.0.0.1", 40);
    expect(result.every((s) => s === 401)).toBe(true);
  });

  it("limits interactive sign-in per client, not across clients", async () => {
    const a = await statuses("/oauth/authorize", "10.0.0.2", 21);
    expect(a.slice(0, 20).every((s) => s === 400)).toBe(true);
    expect(a[20]).toBe(429);
    // Another client behind the same ingress keeps its own budget.
    expect(await statuses("/oauth/authorize", "10.0.0.3", 1)).toEqual([400]);
  });

  it("gives the token endpoint room for parallel refreshes", async () => {
    const result = await statuses("/oauth/token", "10.0.0.4", 40, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: "{}",
    });
    expect(result.every((s) => s === 400)).toBe(true);
  });

  it("leaves discovery and health endpoints unlimited", async () => {
    const result = await statuses("/health", "10.0.0.5", 30);
    expect(result.includes(429)).toBe(false);
  });
});
