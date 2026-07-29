# Claude custom connector setup

> **Beta caveat:** Claude's custom connector / remote MCP support is
> actively evolving. Verify against Anthropic's current documentation for
> exact steps, especially around Dynamic Client Registration behavior and
> where connectors are configured (personal account vs. Team/Enterprise
> workspace settings).

## Prerequisites

Same as `docs/chatgpt-setup.md`: a reachable deployment, at least one
connected Microsoft account, and profile bindings set up from `/accounts`.

## Adding a connector

In Claude (claude.ai or Claude Desktop): **Settings -> Connectors -> Add
custom connector**, then:

1. **Name**: a human-readable label (see the three examples below for
   suggested names).
2. **Remote MCP server URL**: the specific `/mcp/...` endpoint for this
   connector — same three examples as ChatGPT.
3. Claude discovers the OAuth endpoints the same way as ChatGPT (Protected
   Resource Metadata -> Authorization Server Metadata -> Dynamic Client
   Registration -> authorize -> token), since this server exposes the same
   standards-based discovery documents to both clients.
4. You'll go through the same owner-login + one-click-allow screen as in
   the ChatGPT flow.

## The three profile examples

| Claude connector name | Remote MCP server URL | Notes |
|---|---|---|
| Microsoft To Do – Privat | `https://todo.h-aa.dk/mcp/personal` | Bound to your personal Microsoft account connection |
| Microsoft To Do – Arbejde | `https://todo.h-aa.dk/mcp/work` | Bound to your work/organizational connection |
| Microsoft To Do – Universal | `https://todo.h-aa.dk/mcp` | Requires an explicit `connection_id` per tool call |

## Known Claude-specific behavior / limitations

- Claude surfaces tool `isError` results (including the structured JSON
  error payloads this server returns for consent-required/throttled/
  not-found cases) directly in the conversation; phrasing your own prompts
  to ask Claude to "check the connection status if a tool fails" works
  well given `list_connections`/`get_active_connection` are always
  available.
- As of this writing, Claude's custom connectors are most reliably
  configurable from a personal claude.ai account; Team/Enterprise
  workspace-level connector management may have additional admin-approval
  steps not modeled by this project (this server has no concept of "Claude
  workspace admin" — it only knows about its own single owner claim).
- Tool annotations (`readOnlyHint`, `destructiveHint`) are set on every
  tool (`src/tools/*.ts`) so a client that surfaces "this may modify data"
  warnings (e.g. before `delete_task`, `delete_task_list`,
  `delete_checklist_item`) can do so — verify Claude's current UI actually
  renders these hints if you rely on them for user-facing warnings.

## Differences worth calling out between ChatGPT and Claude here

- Both discover the same standards-based OAuth metadata from this server,
  so the *server-side* behavior is identical for both clients — any
  differences are purely in each client's own connector UI/UX, not in
  anything this project special-cases.
- Neither client is assumed to support more than one active OAuth session
  per distinct Server URL, which is why three distinct URLs (not query
  parameters or headers) are used to represent the three profiles.
