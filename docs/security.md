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

## Known data-integrity risk: checklist items on shared, non-owned lists

**Symptom (reported and reproduced):** on a Microsoft To Do list that is
*shared* and *not owned* by the connected account (`isOwner: false`),
checklist items added or updated via `add_checklist_item`/
`update_checklist_item` were later observed missing, and the parent task's
`body` reset to blank — both triggered by an unrelated, later edit (by any
participant, including via this server itself) to the same task.

**Root-cause investigation (2026-07-29):** the request shapes this server
sends were audited line by line against the current Microsoft Graph API
reference:

- `updateTask`/`completeTask`/`reopenTask` PATCH only the fields explicitly
  provided (`src/graph/todoApi.ts`) — never `checklistItems`, and never
  `body` unless the caller passed one.
- `add_checklist_item`/`update_checklist_item`/`delete_checklist_item` only
  ever call the dedicated `.../checklistItems` (or `.../checklistItems/{id}`)
  sub-resource endpoint — never the parent task endpoint.
- Per Microsoft's own [`todoTask` resource reference](https://learn.microsoft.com/en-us/graph/api/resources/todotask),
  `checklistItems` is a **navigation property** (a separate child
  resource collection), not an inline field — a PATCH to the task itself
  cannot touch it even in principle.
- Per the [`checklistItem` resource reference](https://learn.microsoft.com/en-us/graph/api/resources/checklistitem),
  there is no ETag, `cTag`, or version field on this resource for this
  server to have mishandled — the API exposes none.
- These request-shape properties are locked in as regression tests in
  `tests/unit/todoApi.test.ts`.

No bug was found in how this server constructs Graph requests. Multiple
independent, longstanding user reports (e.g. on
[Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/5219543/microsoft-to-do-shared-list-is-not-syncing))
describe shared Microsoft To Do lists losing data or falling out of sync
across participants — Microsoft To Do's shared-list backend (built on
Exchange Online, per Microsoft's own To Do API overview) has a
long-documented history of this class of issue, independent of any
specific client. The symptom pattern here (checklist items *and* body both
reset together, triggered by a subsequent unrelated edit) is consistent
with a **server-side full-task resync overwriting a stale replica**,
which no client-side request-shaping can prevent or detect — the API
gives no version/ETag signal to guard against it.

**Mitigation implemented pending a Microsoft-side fix:** `add_checklist_item`,
`update_checklist_item`, and `delete_checklist_item` now call
`getTaskList` first and refuse to proceed (`shared_list_write_blocked`
error) when the target list is shared and not owned by the current
connection, unless the caller explicitly passes
`acknowledge_shared_list_risk: true` (see `src/tools/sharedListGuard.ts`).
This does not fix the underlying Microsoft-side risk — it stops the tool
from silently returning "success" on writes that Microsoft's sync may
later discard, and requires explicit, informed opt-in instead.

**If you hit this:** avoid checklist-item writes (from this server or any
other client) on shared lists you don't own where possible; prefer
per-person lists with `linkedResource`/task-level sharing instead of
relying on shared-list sync for anything you can't afford to lose. Use the
`list_checklist_items` read tool to verify actual server-side state before
and after any write, rather than trusting a write's own success response.

## Logging redaction

`src/logging/logger.ts` maintains an explicit `pino` `redact` path list
(`req.headers.authorization`, `req.headers.cookie`, `*.token`, `*.code`,
`*.title`, `*.description`, `*.email`, etc.) and every log call site is
expected to pass only the fields in `LogFields` — treat any log call that
passes raw request/response bodies or Graph payloads as a bug to fix, not
a style nit.
