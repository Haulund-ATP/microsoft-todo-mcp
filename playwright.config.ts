import { defineConfig } from "@playwright/test";

/**
 * OAuth browser-flow tests. These exercise real HTTP endpoints (discovery
 * documents, unauthenticated /mcp rejection, and — where credentials are
 * available — the owner login redirect chain) against a running instance
 * of this server.
 *
 * Not run as part of default CI (no live Azure resources available there);
 * run locally or in a staging environment with a real/dev Key Vault and
 * Storage account (or Azurite + a local secret source) configured via
 * .env.test. See tests/playwright/README.md.
 */
export default defineConfig({
  testDir: "./tests/playwright",
  timeout: 30_000,
  retries: 0,
  reporter: [["list"]],
  use: {
    baseURL: process.env.PLAYWRIGHT_BASE_URL ?? "http://localhost:3000",
    trace: "retain-on-failure",
  },
});
