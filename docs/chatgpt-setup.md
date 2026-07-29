# ChatGPT custom connector setup

> **Beta caveat:** ChatGPT's custom connector / MCP support has moved
> quickly and details (exact menu names, whether OAuth discovery is fully
> automatic, whether Dynamic Client Registration is always attempted) may
> have changed since this was written. Verify against OpenAI's current
> documentation if anything below doesn't match what you see in the UI.

## Prerequisites

- The server deployed and reachable at `PUBLIC_BASE_URL` (e.g.
  `https://todo.h-aa.dk`), passing `scripts/test-production.ps1`.
- At least one Microsoft account connected via `/accounts`, with a profile
  bound (e.g. `personal` -> your personal Microsoft account).

## Adding a connector

In ChatGPT: **Settings -> Connectors -> Add custom connector** (exact path
varies by ChatGPT plan/version), then:

1. **Server URL**: point it at the specific profile endpoint you want this
   connector to represent — see the three examples below. Do **not** point
   two separate ChatGPT connectors at the same profile URL if you want them
   to behave independently in ChatGPT's UI (each connector maps 1:1 to a
   Server URL there).
2. ChatGPT will fetch `/.well-known/oauth-protected-resource` from that
   URL, discover the authorization server (this same origin), then fetch
   `/.well-known/oauth-authorization-server` to find the authorize/token/
   registration endpoints.
3. ChatGPT performs Dynamic Client Registration against `/oauth/register`
   (see `docs/oauth.md` for why DCR was chosen as the primary path).
4. You'll be redirected to `/oauth/authorize`; if you're not already logged
   in as the deployment owner, you'll be bounced through
   `/oauth/owner/login` (a Microsoft sign-in) first, then shown a one-click
   "Allow" screen.
5. After authorizing, ChatGPT holds an access token + refresh token scoped
   to that specific `/mcp/...` resource.

## The three profile examples

| ChatGPT connector name | Server URL | Notes |
|---|---|---|
| Microsoft To Do – Privat | `https://todo.h-aa.dk/mcp/personal` | Bound to your personal Microsoft account connection |
| Microsoft To Do – Arbejde | `https://todo.h-aa.dk/mcp/work` | Bound to your work/organizational connection |
| Microsoft To Do – Universal | `https://todo.h-aa.dk/mcp` | Not bound to any single account — every tool call must include an explicit `connection_id` (get one via `list_connections`) |

Bind the `personal`/`work` profile aliases first from `/accounts` (see
`docs/multi-account.md`) before adding the corresponding ChatGPT connector
— an unbound alias returns HTTP 404 from `/mcp/{alias}`.

## Known ChatGPT-specific behavior / limitations

- ChatGPT's connector UI generally expects the **first** tool call in a
  conversation to succeed cleanly; if a profile is unbound or a connection
  is `consent_required`, expect ChatGPT to surface the tool's structured
  error text (`src/tools/errors.ts`) rather than retry automatically — you
  may need to tell the user to re-authorize.
- Refresh-token lifetime as configured
  (`MCP_OAUTH_REFRESH_TOKEN_TTL_SECONDS`, default 30 days) determines how
  often ChatGPT will need to re-run the authorize flow for a given
  connector; there's no server-side push to force early reauth short of
  revoking the token from `/accounts`.
- ChatGPT connectors are per-workspace/per-account in ChatGPT's own model;
  if multiple people use the same ChatGPT workspace, they'd all be sharing
  this single-owner deployment's connections — this project assumes a
  single human end-user matching `ADMIN_OWNER_CLAIM_VALUE`.
