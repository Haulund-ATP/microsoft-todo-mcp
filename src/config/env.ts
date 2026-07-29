import { z } from "zod";

/**
 * Central environment configuration. Fails fast at startup if required
 * variables are missing, so misconfiguration is caught before any request
 * is served rather than surfacing as a confusing runtime error.
 *
 * Secret values (client secret, storage connection string, signing key)
 * are read here only as a *fallback* for local development. In Azure, the
 * process should prefer Key Vault via `src/config/secrets.ts`.
 */

const envSchema = z.object({
  NODE_ENV: z.enum(["development", "test", "production"]).default("development"),
  PORT: z.coerce.number().int().positive().default(3000),
  LOG_LEVEL: z.enum(["fatal", "error", "warn", "info", "debug", "trace"]).default("info"),

  PUBLIC_BASE_URL: z.string().url(),
  DEFAULT_TIMEZONE: z.string().default("Europe/Copenhagen"),

  AZURE_TENANT_ID: z.string().min(1),
  AZURE_CLIENT_ID: z.string().min(1),
  AZURE_CLIENT_SECRET: z.string().optional(),

  AZURE_KEY_VAULT_URI: z.string().url(),

  AZURE_STORAGE_ACCOUNT_NAME: z.string().min(1).optional(),
  AZURE_STORAGE_CONNECTION_STRING: z.string().optional(),

  MCP_OAUTH_SIGNING_KEY_NAME: z.string().min(1),
  MCP_OAUTH_ISSUER: z.string().url(),
  MCP_OAUTH_ACCESS_TOKEN_TTL_SECONDS: z.coerce.number().int().positive().default(600),
  MCP_OAUTH_REFRESH_TOKEN_TTL_SECONDS: z.coerce.number().int().positive().default(2592000),
  MCP_OAUTH_AUTH_CODE_TTL_SECONDS: z.coerce.number().int().positive().default(120),

  ADMIN_OWNER_CLAIM_NAME: z.string().min(1).default("oid"),
  ADMIN_OWNER_CLAIM_VALUE: z.string().min(1),

  SESSION_COOKIE_SECRET: z.string().min(16),
  CSRF_COOKIE_SECRET: z.string().min(16),
});

export type AppEnv = z.infer<typeof envSchema>;

let cached: AppEnv | undefined;

export function loadEnv(source: NodeJS.ProcessEnv = process.env): AppEnv {
  if (cached) return cached;
  const parsed = envSchema.safeParse(source);
  if (!parsed.success) {
    const issues = parsed.error.issues
      .map((i) => `  - ${i.path.join(".")}: ${i.message}`)
      .join("\n");
    throw new Error(`Invalid/missing environment configuration:\n${issues}`);
  }
  if (!parsed.data.AZURE_STORAGE_ACCOUNT_NAME && !parsed.data.AZURE_STORAGE_CONNECTION_STRING) {
    throw new Error(
      "Either AZURE_STORAGE_ACCOUNT_NAME (managed identity) or AZURE_STORAGE_CONNECTION_STRING (local dev) must be set."
    );
  }
  cached = parsed.data;
  return cached;
}

/** Test-only: clear the cached env so a test can reload with different values. */
export function _resetEnvCacheForTests(): void {
  cached = undefined;
}
