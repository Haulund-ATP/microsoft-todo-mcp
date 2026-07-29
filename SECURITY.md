# Security Policy

## Supported versions

This project follows a rolling-release model on the `main` branch. Only the
latest deployed revision is supported; there are no maintained LTS branches.

## Reporting a vulnerability

Please **do not** open a public GitHub issue for security vulnerabilities.

Instead, report privately using one of these channels, in order of
preference:

1. **GitHub Security Advisories** — open a draft advisory via
   "Security" → "Report a vulnerability" on this repository. This is the
   preferred channel: it is private by default and lets us collaborate on a
   fix before disclosure.
2. **Email** — send details to the address listed in the repository owner's
   GitHub profile, with subject line `SECURITY: microsoft-todo-mcp`. Include
   `[email protected]` in encrypted form if you have PGP; otherwise plain
   email is acceptable for an initial report (avoid including live secrets,
   tokens, or personal data in the report itself — describe the issue and
   reproduction steps instead).

Please include:

- A description of the vulnerability and its impact.
- Steps to reproduce, or a proof-of-concept if available.
- The affected component (upstream Microsoft OAuth flow, MCP-side OAuth 2.1
  authorization server, token/cache storage, admin UI, Graph tool handlers,
  infrastructure/Bicep, CI/CD).
- Whether the issue requires authentication, and at what privilege level.

## Response process

- **Acknowledgement:** within 3 business days.
- **Initial assessment (severity, affected versions):** within 7 business
  days.
- **Fix or mitigation:** timeline depends on severity; critical
  authentication/token issues are prioritized and typically addressed within
  14 days, with a temporary mitigation (e.g., disabling an endpoint,
  rotating signing keys) if a full fix takes longer.
- **Disclosure:** we prefer coordinated disclosure. We will credit reporters
  (unless anonymity is requested) in the release notes / advisory once a fix
  is shipped.

## Scope

In scope:

- The MCP server application code in `src/` (OAuth flows, token storage,
  crypto, Graph tool handlers, admin UI, routing).
- The infrastructure-as-code in `infra/` as it pertains to security posture
  (RBAC scope, network exposure, secret handling).
- The CI/CD workflows in `.github/workflows/`.

Out of scope:

- Vulnerabilities in upstream dependencies without a demonstrated exploit
  path through this project (report those upstream instead, though we do
  want to know if a dependency vulnerability is directly exploitable here).
- Vulnerabilities that require compromising a user's own Microsoft account,
  Azure subscription, or GitHub account credentials outside of this
  project's control.
- Denial-of-service reports that rely purely on volumetric traffic against a
  scale-to-zero Container Apps deployment (expected/rate-limited behavior).

## Security design notes relevant to reporters

- Two distinct OAuth systems exist: (1) this server acting as an OAuth 2.1
  **authorization server** toward ChatGPT/Claude, and (2) this server acting
  as an OAuth **client** toward Microsoft Entra ID/Graph. Please specify
  which one your report concerns — see `docs/oauth.md`.
- MSAL token caches are encrypted at rest with AES-256-GCM before being
  persisted to Azure Table/Blob Storage; the data-encryption key is held in
  Key Vault and accessed only via user-assigned managed identity.
- The admin page (`/accounts`) is gated on a single, configurable, stable
  identity claim (`ADMIN_OWNER_CLAIM_VALUE`), not email. A report that this
  gate can be bypassed is treated as critical.
