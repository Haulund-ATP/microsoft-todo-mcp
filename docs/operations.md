# Operations

## Finding the owner's identity claim value

`ADMIN_OWNER_CLAIM_VALUE` must be a stable claim from the deployment
owner's Microsoft account id token — by default `oid` (object id), which
does not change even if the account's email/UPN is renamed.

To find it before the first deploy:

1. Temporarily set `ADMIN_OWNER_CLAIM_VALUE=DISCOVER` (any placeholder) and
   deploy, or run the server locally against a dev tenant.
2. Attempt an owner login (`/oauth/owner/login`); it will be denied, but
   the audit log (`admin.login_denied`) is emitted — however it
   intentionally does **not** log the claim value itself (no PII in logs).
3. Instead, use `scripts/provision.ps1`, which prompts you interactively
   to paste the value — retrieve it ahead of time via `az ad signed-in-user
   show --query id` (for a work/school account) or by inspecting the
   decoded id token from a manual OAuth test against the Entra app
   registration (for a personal Microsoft account, use
   `https://jwt.ms` after a manual sign-in test — never paste a real token
   into a third-party site in a production incident, only for initial
   setup with a throwaway test).
4. Update the Key Vault secret `admin-owner-claim-value` directly if you
   need to change it later:
   ```powershell
   az keyvault secret set --vault-name <kv-name> --name admin-owner-claim-value --value <new-value>
   ```
   Then restart the Container App revision so the new value takes effect
   (env values referencing Key Vault secrets are read at container start).

## Rotating secrets

| Secret | How to rotate | Blast radius if delayed |
|---|---|---|
| `mcp-oauth-signing-key` | Generate a new ES256 JWK (see `provision.ps1`'s `New-EcP256JwkJson`), update the Key Vault secret, restart the app. All previously issued MCP access tokens become invalid immediately (clients re-authorize via refresh token). | Low — short-lived tokens anyway. |
| `session-cookie-secret` / `csrf-cookie-secret` | Rotate via `az keyvault secret set` + restart. All active owner sessions are invalidated. | Low. |
| `msal-cache-encryption-key` | **Do not rotate casually** — rotating without re-encrypting existing cache blobs makes them undecryptable, forcing re-authentication of every connection. If rotating, plan a maintenance window and re-run the Microsoft sign-in for each connection afterward. | High if done carelessly. |
| `graph-client-secret` | Rotate on the Entra app registration first, then update the Key Vault secret, then restart. | Medium — Graph calls fail with 401 until updated. |

## Reading audit logs

Audit events (`src/logging/audit.ts`) are emitted as structured `pino`
log lines with `"audit": true` and an `action` field (see the
`AuditAction` union for the full list). In Azure, if `enableLogAnalytics`
is turned on, query with:

```kusto
ContainerAppConsoleLogs_CL
| where Log_s has "\"audit\":true"
| project TimeGenerated, Log_s
| order by TimeGenerated desc
```

Without Log Analytics, use `az containerapp logs show --follow` for
live tailing, or `az containerapp logs show` for a recent window.

## Diagnosing a stuck/erroring connection

1. Check `/accounts` — the status badge (`active` / `expired` / `revoked`
   / `consent_required` / `error`) is the first signal.
2. `consent_required` — a tenant admin needs to grant consent; the admin
   consent URL is included in the error text surfaced both in `/accounts`
   activity and in any tool call that hits it (`src/tools/errors.ts`).
3. `error` — check `errorDetail` on the connection record (only generic
   markers like `silent_token_acquisition_failed` are stored, never raw
   Graph error bodies) and the correlated structured logs by
   `correlationId`.
4. When in doubt, use the "Reauth" action on `/accounts` — it re-runs the
   Microsoft sign-in and updates the connection in place.

## Manual verification: shared-list checklist behavior (requires two real accounts)

`tests/unit/todoApi.test.ts` and `tests/unit/sharedListGuard.test.ts` cover
everything that can be verified against a mocked Graph client: request
shapes, field-scoping, and the shared-list write guard. They cannot
reproduce Microsoft's own shared-list sync behavior between two real
participants — that requires two real Microsoft accounts sharing a real
list, which can't be automated in CI without live credentials. See
`docs/security.md` for the known-risk writeup. To manually verify:

1. Share a Microsoft To Do list from Account A to Account B (in the To Do
   app, not via this server).
2. Connect Account A here, bind a profile, `add_checklist_item` twice on a
   task in the shared list (pass `acknowledge_shared_list_risk: true`).
3. `list_checklist_items` — confirm both are present.
4. In the To Do app as Account B, add a third checklist item manually.
5. `list_checklist_items` again as Account A — confirm all three are
   present (this is the step known to sometimes fail due to Microsoft's
   sync — if it does, that's the documented known risk, not this server).
6. Repeat with the write coming from Account B and the read from Account A
   to check the reverse direction.

## Known TODOs left for a human / live-Azure step

- **Exact managed-identity RBAC role assignment propagation timing**: role
  assignments can take a few minutes to propagate; if the very first
  deploy's container fails readiness checks with authorization errors
  against Key Vault/Storage, wait ~5 minutes and restart the revision
  before assuming something is misconfigured.
- **Admin-UI form for pre-registering a static OAuth client** (see
  `docs/oauth.md` "pre-registered" path) is not built — currently requires
  directly inserting a row into the `oauthclients` Table via `az storage
  entity insert` or the Storage Explorer if you need this fallback.
- **Log Analytics shared key wiring** in
  `infra/modules/containerAppsEnvironment.bicep` is left as an empty
  string by default (see the comment in `infra/main.bicep`) to avoid
  passing a plaintext key through Bicep parameters; if you enable
  `enableLogAnalytics`, wire the shared key via a `listKeys()` reference
  or a follow-up `az containerapp env update` call.
- **CIMD (Client ID Metadata Documents) support** is not implemented — see
  `docs/oauth.md` for why DCR was chosen instead, and what a future CIMD
  add-on would touch.
