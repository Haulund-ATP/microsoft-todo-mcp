# Playwright OAuth flow tests

These tests exercise real HTTP endpoints against a **running instance** of
this server. They are not run in default CI (`.github/workflows/ci.yml`)
because they require a real (or emulated) Azure Key Vault + Storage
Account — the discovery endpoints (`/.well-known/jwks.json`) and the
owner-login flow both need a working `SecretStore`/`ConnectionsRepo`.

## Running locally

1. Provision a dev-tier Key Vault + Storage Account (or point at Azurite +
   a local secret source you've adapted `src/config/secrets.ts` for).
2. Seed the required secrets (see `scripts/provision.ps1`).
3. Copy `.env.example` to `.env.test` and fill in real dev values.
4. Start the server against that config:
   ```bash
   node --env-file=.env.test dist/index.js
   ```
5. In another terminal:
   ```bash
   PLAYWRIGHT_BASE_URL=http://localhost:3000 npm run test:e2e
   ```

## What is covered

- `oauth-discovery.spec.ts` — Protected Resource Metadata, Authorization
  Server Metadata, and JWKS documents are well-formed and internally
  consistent (issuer/endpoints all point at the same base URL).
- `mcp-auth.spec.ts` — `/mcp`, `/mcp/personal` reject unauthenticated
  requests with 401 and a `WWW-Authenticate` challenge; dynamic client
  registration followed by an authorize-without-login redirects to owner
  login rather than silently issuing a code.

## What is intentionally NOT covered here

A full round-trip through a real Microsoft Entra ID consent screen isn't
automated — that would require either a scripted headless login against a
real Microsoft account (fragile, and a credentials-handling liability in a
public repo) or a mocked Entra ID, which isn't the standards conformance
this suite is trying to verify. Instead, `src/graph/upstreamOAuth.ts`'s
pure logic (PKCE generation, admin-consent-error detection) is covered by
the vitest unit tests where possible; the actual live redirect exchange is
a manual step in `docs/operations.md`'s deployment checklist.
