# Security model

See also `SECURITY.md` for vulnerability reporting, and `docs/oauth.md` for
the detailed OAuth security properties.

## Threat model summary

This is a **single-owner** deployment: exactly one person is meant to be
able to administer connections and use the MCP tools, via ChatGPT/Claude
connectors they control. The primary threats considered:

1. An attacker who can reach the public internet endpoint, but has no
   valid Microsoft account matching `ADMIN_OWNER_CLAIM_VALUE` and no valid
   MCP OAuth token.
2. A leaked MCP access/refresh token (e.g. from a compromised ChatGPT/
   Claude session) being used to call `/mcp*` tools.
3. Compromise of the Storage Account or its backups/exports.
4. A malicious or buggy MCP client attempting cross-account access (e.g.
   passing a different `connection_id` than the one a bound profile
   implies).

## Mitigations by threat

**(1) Unauthenticated attacker:**
- `/mcp*` requires a valid bearer JWT (`requireBearer.ts`); anything else
  gets a 401 with a `WWW-Authenticate` challenge, never a stack trace.
- `/accounts` and `/oauth/authorize`'s approve step require an owner
  session cookie; anyone else is redirected to Microsoft sign-in and then
  rejected if their claim doesn't match.
- `/health`, `/ready`, `/version` are intentionally unauthenticated but
  contain no tenant/account/secret data — see the "no sensitive data" rule
  enforced by code review on `src/http/health.ts`.
- Rate limiting (`src/admin/rateLimit.ts`) on the OAuth and admin surfaces.

**(2) Leaked MCP token:**
- Access tokens are short-lived JWTs (default 10 minutes) — a leaked
  access token has a small window.
- Refresh tokens rotate on every use; reuse of an already-rotated token
  revokes the entire descendant chain (`OAuthTokensRepo.revokeChainFrom`),
  treating replay as a compromise signal.
- Refresh tokens are stored **hashed** (SHA-256) in Table Storage, so a
  storage-level leak does not itself yield usable bearer tokens.
- `/oauth/revoke` lets the owner invalidate a specific refresh token
  on demand.

**(3) Storage compromise:**
- OAuth authorization codes and refresh tokens: hashed at rest (see
  above).
- MSAL token caches: AES-256-GCM encrypted before being written to Blob
  Storage, with the connection id bound in as AAD — a cache blob cannot be
  silently reattached to a different connection record. The
  data-encryption key lives in Key Vault, accessed only via the
  Container App's user-assigned managed identity (least-privilege: Key
  Vault Secrets User, Storage Table/Blob Data Contributor — no broader
  roles).
- No shared-key/connection-string credentials are used by the running
  application in production (`AZURE_STORAGE_CONNECTION_STRING` is a local-
  dev-only fallback, explicitly gated on `NODE_ENV !== "production"` in
  `src/storage/clients.ts`).

**(4) Cross-account confusion:**
- See `docs/multi-account.md` — bound profiles reject a mismatched
  `connection_id` outright rather than silently using the one from the
  route or the one from the argument.

## Defense in depth on the admin surface

- CSRF: double-submit signed cookie token (`src/admin/csrf.ts`), required
  on every state-changing `/accounts` and `/admin/connect/*` POST.
- Cookies: `HttpOnly`, `Secure`, `SameSite=Lax` (owner session) or `Strict`
  (CSRF token cookie itself, since it's read by the page's own forms
  only).
- CSP via Helmet: `default-src 'self'`, no third-party scripts, no framing
  (`frame-ancestors 'none'`).
- `trust proxy` is enabled for the single hop from Container Apps ingress,
  not blindly for arbitrary `X-Forwarded-*` chains.

## What is deliberately NOT implemented

- Multi-user consent / per-user data isolation — this is a single-owner
  tool. If you need multiple distinct human users each with their own
  Microsoft accounts and no shared visibility, this project's data model
  (one owner claim, N connections all manageable by that one owner) is the
  wrong starting point without further work.
- IP allow-listing / network isolation (Private Endpoints) — not applied,
  since the whole point is public reachability for ChatGPT/Claude's own
  infrastructure to call in; consider this if your threat model changes.

## Logging redaction

`src/logging/logger.ts` maintains an explicit `pino` `redact` path list
(`req.headers.authorization`, `req.headers.cookie`, `*.token`, `*.code`,
`*.title`, `*.description`, `*.email`, etc.) and every log call site is
expected to pass only the fields in `LogFields` — treat any log call that
passes raw request/response bodies or Graph payloads as a bug to fix, not
a style nit.
