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

## Checklist items on shared lists: read-after-write verification, not a write block

**Symptom originally reported:** on a Microsoft To Do list that is *shared*
and *not owned* by the connected account (`isOwner: false`), checklist items
added or updated via `add_checklist_item`/`update_checklist_item` were later
observed missing, and the parent task's `body` reset to blank.

**2026-07-29 investigation, round 1 (superseded):** an earlier version of
this document concluded this was a Microsoft-backend sync issue, based on
the request shapes this server sends and independent forum reports, and
shipped a mitigation that blocked `add_checklist_item`/`update_checklist_item`/
`delete_checklist_item` on any shared, non-owned list unless the caller
passed `acknowledge_shared_list_risk: true`. That conclusion and that
mitigation were both wrong in scope:

- The "no ETag on todoTask" claim was **factually incorrect** — Graph's own
  documented response example for updating a task
  (https://learn.microsoft.com/en-us/graph/api/todotask-update) shows
  `"@odata.etag": "W/\"...\""` on the response. The earlier check only read
  the resource's "Properties" table, which omits OData protocol annotations
  like `@odata.etag` — it never inspected an actual response. This is now
  captured on every task read (`TaskItem.etag`, `src/graph/todoApi.ts`).
- Blocking writes on every shared, non-owned list made a core use case of
  this tool — shared lists — unusable without an extra opt-in flag on every
  call, based on forum reports rather than a reproduction against this
  server's own code paths.
- checklistItem responses, by contrast, are confirmed (via the documented
  create-response example) to carry no `@odata.etag` or
  `lastModifiedDateTime` — only `@odata.context`, `displayName`,
  `createdDateTime`, `isChecked`, `id`. That part of the original finding
  held up under the same actual-response check.

**Current approach:** normal checklist and task writes — `create_task`,
`update_task`, `complete_task`, `reopen_task`, `add_checklist_item`,
`update_checklist_item`, `delete_checklist_item` — work unconditionally on
shared lists, regardless of ownership. There is no `isOwner` gate and no
required acknowledgement flag. Instead:

- Every checklist write (`add_checklist_item`/`update_checklist_item`/
  `delete_checklist_item`) re-reads the full checklist afterward
  (`src/graph/verification.ts`) with a few short, bounded retries, and
  returns a `verificationStatus` of `verified`, `delayed`, or `inconsistent`
  alongside the fresh checklist, its count, and the task's `etag`/
  `lastModifiedDateTime`. An `inconsistent` result never triggers an
  automatic retry of the write, an automatic replacement item, or deletion
  of anything else — it's surfaced as a warning with the observed item IDs
  so the caller (and the human on the other end) can decide what to do.
- `update_task`/`complete_task`/`reopen_task` read the checklist immediately
  before and after the task-level operation and report whether any
  previously-visible checklist item IDs went missing, without blocking the
  operation or attempting any reconstruction.
- Every Graph call carries a generated `client-request-id` and
  `return-client-request-id: true` (`src/graph/diagnostics.ts`), and the
  Graph-assigned `request-id` (when returned) is captured for diagnostic
  logging — so a genuine Microsoft-side inconsistency can be pinpointed by
  request ID rather than inferred from a forum thread.
- `get_task_with_checklist` (`src/tools/readTools.ts`) always performs fresh
  Graph reads (list + task + checklist, never cached) so a caller — or a
  second participant on the same shared list — can independently confirm
  actual server-side state.

**Root-cause status:** no bug has been found in how this server constructs
Graph requests for task or checklist operations (see the field-scoping
regression tests in `tests/unit/todoApi.test.ts`) — every write is scoped to
exactly the resource it targets and never touches an unrelated one. Whether
Microsoft's shared-list backend itself drops checklist items has **not**
been proven or disproven here: doing so requires two real Microsoft accounts
sharing a list, both reading the same task directly via Graph after a write,
which is outside what a single-account deployment/investigation can verify
end-to-end. `scripts/live-checklist-test.ts` is an opt-in, single-account
live check that exercises the same request/verification path against a real,
dedicated test list — see "Live checklist verification" in
`docs/operations.md` for how to run it, and for the two-account manual
procedure if a second participant is available to compare notes.

**If you observe a discrepancy:** use `get_task_with_checklist` from both
sides (or `list_checklist_items` plus `get_task`) to capture each account's
own Graph view, note the `verificationStatus` and any `graphRequestId`
values logged for the write in question, and compare — per the classification
rule above, only attribute it to Microsoft's backend if the writer's own
fresh Graph read is correct while a different participant's fresh Graph read
(not a To Do client's cached view) disagrees.

## Logging redaction

`src/logging/logger.ts` maintains an explicit `pino` `redact` path list
(`req.headers.authorization`, `req.headers.cookie`, `*.token`, `*.code`,
`*.title`, `*.description`, `*.email`, etc.) and every log call site is
expected to pass only the fields in `LogFields` — treat any log call that
passes raw request/response bodies or Graph payloads as a bug to fix, not
a style nit.
