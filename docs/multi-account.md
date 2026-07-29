# Multi-account and multi-profile model

## Two concepts: connections and profiles

- **Connection** (`src/storage/types.ts` `Connection`): one authenticated
  Microsoft account — personal or organizational — with its own encrypted
  MSAL token cache, status, and metadata. Created by connecting a new
  account from `/accounts` (`src/admin/connectRouter.ts`,
  `src/graph/upstreamOAuth.ts`).
- **Profile** (`ProfileBinding`): a URL-safe alias bound to exactly one
  connection, exposed as its own MCP route: `/mcp/{alias}`. Created/edited
  from `/accounts` (`src/storage/profilesRepo.ts`).

A connection can exist without any profile bound to it (reachable only via
the universal `/mcp` endpoint with an explicit `connection_id`). A profile
alias can only ever point at one connection at a time, but you can rebind
it later (e.g. if you replace your work account).

## Reserved aliases

`mcp`, `accounts`, `health`, `ready`, `version`, `.well-known` can never be
used as a profile alias — `isValidAlias()` in `src/storage/profilesRepo.ts`
enforces this, and the route layer (`src/profiles/router.ts`) also checks
`RESERVED_ALIASES` defensively before ever calling into profile resolution.

## Why the server never guesses between accounts

Every tool handler resolves its Graph connection via
`src/tools/context.ts` `resolveConnectionId()`:

- On a bound route (`/mcp/personal`, `/mcp/work`, `/mcp/{alias}`), the
  connection is fixed by the route itself. If a tool call somehow also
  passes a *different* `connection_id`, that's rejected as an error rather
  than silently overridden — a bound profile can never be redirected to
  another account mid-call.
- On the universal `/mcp` route, there is no bound connection, so every
  tool requires an explicit `connection_id` argument
  (`connectionIdInputShape` in `src/tools/accounts.ts`). Calling a tool
  without one raises `UnresolvedConnectionError`, whose message tells the
  caller (the LLM) to run `list_connections`/`list_profiles` first.

This means there is no code path where "create_task" could accidentally
land in the wrong Microsoft account because the server made an assumption
about which one was meant.

## Typical setup

```
1. /accounts -> "Connect a new Microsoft account" -> sign in with personal MSA
2. /accounts -> bind profile alias "personal" -> that connection
3. /accounts -> "Connect a new Microsoft account" -> sign in with work account
4. /accounts -> bind profile alias "work" -> that connection
5. Add three connectors in ChatGPT/Claude: personal, work, and universal
   (see docs/chatgpt-setup.md / docs/claude-setup.md)
```

## Renaming, revoking, deleting

- **Rename** a connection's alias (cosmetic, does not affect profile
  bindings, which reference `connectionId` not alias) — form on
  `/accounts`.
- **Reauth** — re-runs the Microsoft sign-in for an existing connection,
  updating its cached tokens/metadata in place rather than creating a
  duplicate connection (`completeMicrosoftSignIn(..., reauthConnectionId)`
  in `src/graph/upstreamOAuth.ts`).
- **Revoke** — marks the connection `status: "revoked"` without deleting
  its data; tool calls against it will fail with a clear error, but the
  admin can still see history/last-used before deciding to fully delete.
- **Delete** — permanently removes the connection row and its encrypted
  token-cache blob (`ConnectionsRepo.delete`). Any profile still bound to
  it will 404 until rebound.
