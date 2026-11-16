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
3. **Modules per profile.** Each profile binding carries the modules it exposes and whether it
   may send (mail, Teams messages). Missing → `["todo"]`, no send. A To Do-only profile (e.g. a
   family member's) never sees mail tools. Tools are only registered for enabled modules, which
   also keeps the tool list small.
4. **Never guess the account** — unchanged: tools resolve their connection exactly as today.

## Modules

| Module | Scopes (delegated) | Personal MSA | Admin consent |
|---|---|---|---|
| `todo` | Tasks.ReadWrite | yes | no |
| `mail` | Mail.ReadWrite, Mail.Send | yes | no |
| `calendar` | Calendars.ReadWrite | yes | no |
| `contacts` | Contacts.ReadWrite, People.Read | yes | no |
| `files` (OneDrive + `.loop` files) | Files.ReadWrite.All | yes | no |
| `sites` (SharePoint) | Sites.ReadWrite.All | no | no (tenant policy may differ) |
| `teams` (chats, channel posting) | Chat.ReadWrite, ChatMessage.Send, Team.ReadBasic.All, Channel.ReadBasic.All, ChannelMessage.Send | no | no |
| `teams_channels` (read channel messages) | ChannelMessage.Read.All | no | **yes** |

Base scopes on every request: `User.Read openid profile offline_access`.

**Loop.** `.loop` files in OneDrive/SharePoint can be found and exported read-only through the
`files`/`sites` modules (Graph `driveItem` content/format export). Their content is Fluid and
cannot be edited through Graph. Loop *workspaces* live in SharePoint Embedded containers and need
a SharePoint Embedded admin to grant this app guest-app access per tenant
(`Set-SPOApplicationPermission`); that is phase 4 and only for tenants the owner administers.

## Write and send policy

The owner chose full write including send. Safeguards that stay on regardless:

- `send_mail` / `send_chat_message` / `send_channel_message` are only registered on profiles
  with `allowSend` set; otherwise drafts only.
- Every send/delete is written to the audit log (recipient count, never bodies).
- Deletes reuse the To Do "safe delete" idea: mail goes to Deleted Items (no hard delete tool),
  files go to the recycle bin.
- Tool descriptions remind the model that mail/file/chat content is untrusted data, never
  instructions (prompt-injection exfiltration is the main risk of a send-capable connector).

## Phases

0. **Foundation** — module registry, `grantedModules` per connection, modules/allowSend per
   profile, module-aware token acquisition and Graph client, "add modules" re-auth in `/accounts`,
   tool registration gated by module. No new Graph tools. Existing connections unaffected.
1. **Mail + calendar + contacts** — search/read/thread/attachments, drafts, send, reply/forward,
   move/flag; events search/create/update/respond, free/busy; people/contacts.
2. **Files + SharePoint + Loop files** — search, list, read (text extraction for Office/PDF where
   Graph offers a format conversion), upload/replace, create folder, move/copy/rename, sites and
   libraries.
3. **Teams** — teams/channels/chats list, chat and channel messages read/send.
4. **Loop workspaces** (optional) — SharePoint Embedded guest-app access, per tenant.

## Owner actions (cannot be done from code)

- Add the new delegated Graph permissions to the Entra app registration (so the admin-consent
  URL covers them). Dynamic consent works without this for user-consentable scopes.
- Per organizational tenant: connect the account from `/accounts` with the wanted modules; grant
  admin consent where a tenant requires it (`teams_channels` always; others depend on the
  tenant's user-consent policy).
- CloudPoint: connecting a third-party app needs that tenant's admin consent and must be allowed
  by its policies — confirm before connecting.
