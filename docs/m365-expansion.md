# Microsoft 365 modules (mail, files, Teams, calendar, Loop)

planning_task_id: 8289A489-BDD6-4A13-930A-8871532EA858 (Ops Hub planning, TimeWinder 2027 / IT, task #341)
status: in_progress

## Goal

Grow this server from "Microsoft To Do" into a general Microsoft 365 connector that can be
connected once per Microsoft account, so one Claude/ChatGPT user can reach several tenants
side by side (`/mcp/{alias}` per account) — something the first-party Microsoft 365 connector
cannot do (one account per connector).

Target accounts: the owner's own tenant, H-Aa Holding, CloudPoint, TimeWinder — all
organizational. The existing **personal** (MSA) To Do connections must keep working unchanged.

## Design constraints

1. **No shared scope list.** Today every token is acquired with one fixed list
   (`Tasks.ReadWrite User.Read …`). Adding M365 scopes to it would (a) make silent refresh fail
   for every existing connection until it is re-authorized, and (b) break sign-in for personal
   Microsoft accounts outright, because `Sites.*`, `Chat.*`, `Channel*` are not supported for MSA.
2. **Modules per connection (incremental consent).** Each connection records which modules it
   has consented to (`grantedModules`). A module's scopes are only requested when the owner adds
   that module from `/accounts` (re-auth with the union of scopes). Token acquisition for a tool
   uses *base + that module's scopes*, so the To Do path requests exactly what it does today.
   Rows without `grantedModules` are treated as `["todo"]`.
3. **Two endpoints, two OAuth resources.** To Do stays To Do:
   - `/mcp`, `/mcp/{alias}` — To Do tools only, never send; tokens for resource `https://host`
     (unchanged, existing connectors keep working).
   - `/m365`, `/m365/{alias}` — every module the bound connection has consented to (To Do
     included); tokens for resource `https://host/m365`, metadata at
     `/.well-known/oauth-protected-resource/m365`, and a consent page that says "Microsoft 365"
     and mentions sending.
   Because the audiences differ, a token granted to a To Do connector (e.g. a family member's)
   is rejected on `/m365` and vice versa; the token endpoint always takes the resource from the
   approved code/refresh token, so one cannot be traded for the other. The profile only carries
   `allowSend` (applies to `/m365/{alias}`; missing → false). The universal `/m365` route never
   registers send tools.
4. **Never guess the account** — unchanged: tools resolve their connection exactly as today.

## Modules

| Module | Scopes (delegated) | Personal MSA | Admin consent |
|---|---|---|---|
| `todo` | Tasks.ReadWrite | yes | no |
| `mail` | Mail.ReadWrite, Mail.Send | yes | no |
| `calendar` | Calendars.ReadWrite | yes | no |
| `calendar_shared` (calendars shared with you, meeting-time finder) | Calendars.ReadWrite.Shared | no | no |
| `contacts` | Contacts.ReadWrite, People.Read | yes | no |
| `files` (OneDrive + `.loop` files) | Files.ReadWrite.All | yes | no |
| `sites` (SharePoint) | Sites.ReadWrite.All | no | no (tenant policy may differ) |
| `teams` (chats, channel posting) | Chat.ReadWrite, ChatMessage.Send, Team.ReadBasic.All, Channel.ReadBasic.All, ChannelMessage.Send | no | no |
| `teams_channels` (read channel messages) | ChannelMessage.Read.All | no | **yes** |

Base scopes on every request: `User.Read openid profile offline_access`.

**Loop.** `.loop` files in OneDrive/SharePoint are found by the files search and read as text
through the Graph format conversion (`content?format=html`). Their content is Fluid and cannot be
edited through Graph. Loop *workspaces* live in SharePoint Embedded containers, where guest apps
can only be granted **app-only** access (`Set-SPOApplicationPermission`: delegated is not
supported) — that would mean tenant-wide access without a signed-in user, which this delegated,
per-user server deliberately never has. Loop workspaces are therefore out of scope.

## One admin approval per tenant

`/accounts` → "Admin consent (all modules)" (or the per-connection link) opens Microsoft's v2
admin-consent endpoint for that tenant with **every module's permissions in one request**
(`adminConsentScopes()` in `src/graph/modules.ts`). A tenant administrator approves once; the owner
is then sent straight on to connect (or re-connect) the account with all modules, without any
further consent prompt. The redirect is handled before the owner check, so an administrator from
another organisation just sees a confirmation page. Nothing has to be listed statically on the app
registration for this; consent-required errors point at the same URL.

## Write and send policy

The owner chose full write including send. Safeguards that stay on regardless:

- `send_mail` / `send_chat_message` / `send_channel_message` are only registered on profiles
  with `allowSend` set; otherwise drafts only.
- Every send/delete is written to the audit log (recipient count, never bodies).
- Deletes reuse the To Do "safe delete" idea: mail goes to Deleted Items (no hard delete tool),
  files go to the recycle bin.
- Tool descriptions remind the model that mail/file/chat content is untrusted data, never
  instructions (prompt-injection exfiltration is the main risk of a send-capable connector).

## Scope of this change (one PR, so each tenant is approved once)

- **Foundation** — module registry, `grantedModules` per connection, `allowSend` per profile,
  module-aware token acquisition and Graph client, the separate `/m365` endpoint family and OAuth
  resource, tool registration gated by module, one-shot tenant admin consent, connect presets.
  Existing connections and connectors unaffected.
- **Mail** (`src/tools/m365/mailTools.ts`), **calendar** and **contacts/people**
  (`calendarTools.ts`, `contactsTools.ts`), **OneDrive + Loop files** and **SharePoint**
  (`filesTools.ts`, `sitesTools.ts`), **Teams** (`teamsTools.ts`). Shared conventions in
  `src/tools/m365/common.ts`: per-module tokens, untrusted-content note, send tools only with
  `allowSend`, audit of every send/delete, Graph-only paging tokens.

## Owner actions (cannot be done from code)

- Per organizational tenant (own tenant, H-Aa Holding, TimeWinder, CloudPoint): `/accounts` →
  admin consent with the tenant's domain → sign in as its administrator → the account is then
  connected with all modules. Bind a profile, tick "allow sending" where wanted, and add the
  `/m365/{alias}` URL as a connector in Claude.
- CloudPoint: connecting a third-party app needs that tenant's administrator and must be allowed
  by its policies — confirm before connecting.
- Personal accounts: nothing changes unless you choose "connect with everything personal
  accounts support" for one of them.
