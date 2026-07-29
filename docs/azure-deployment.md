# Azure deployment

## Prerequisites

- Azure subscription with permission to create resource groups, Container
  Apps, Key Vault, Storage Accounts, and role assignments.
- Azure CLI with the Bicep extension (`az bicep install`).
- A Microsoft Entra app registration for upstream Graph access:
  - Sign-in audience: **Accounts in any organizational directory and
    personal Microsoft accounts** (`AzureADandPersonalMicrosoftAccount`).
  - Redirect URI: `https://<your-domain>/oauth/microsoft/callback` (web
    platform).
  - API permissions: `Tasks.ReadWrite`, `User.Read` (delegated).
  - A client secret (stored in Key Vault, never in this repo).
- Node.js 22+ and Docker locally if you plan to build images yourself
  (`scripts/deploy.ps1 -Build`), or a CI pipeline that does it for you
  (`.github/workflows/deploy.yml`).

## First-time provisioning

```powershell
./scripts/login.ps1
./scripts/validate-environment.ps1
$region = ./scripts/select-region.ps1

./scripts/provision.ps1 `
    -ResourceGroupName rg-microsoft-todo-mcp `
    -Location $region `
    -ResourceSuffix prod01 `
    -AzureClientId <graph-app-registration-client-id> `
    -AzureTenantId <your-tenant-id> `
    -PublicBaseUrl https://todo.h-aa.dk `
    -BudgetNotificationEmail you@example.com
```

`provision.ps1`:

1. Creates the resource group.
2. Deploys `infra/main.bicep` **without** a custom domain bound yet (the
   Container App gets its default `*.azurecontainerapps.io` FQDN first —
   see `docs/dns.md` for why domain binding is a separate step).
3. Generates and seeds Key Vault secrets that have no safe external
   source (OAuth signing key, cookie secrets, cache encryption key),
   skipping any that already exist so re-running is safe.
4. Prompts you once for the Graph app registration's client secret and the
   owner's identity claim value (see `docs/operations.md` for how to find
   the latter).

## Deploying a real image

```powershell
./scripts/deploy.ps1 `
    -ResourceGroupName rg-microsoft-todo-mcp `
    -ContainerAppName todomcp-app-prod01 `
    -Image ghcr.io/<you>/microsoft-todo-mcp:sha-abc123 `
    -Build
```

Or let `.github/workflows/deploy.yml` do this on every push to `main`
(after `configure-github-oidc.ps1` has been run once).

## Custom domain

See `docs/dns.md` for the full DNS story; short version:

```powershell
./scripts/configure-domain.ps1 `
    -ResourceGroupName rg-microsoft-todo-mcp `
    -ContainerAppEnvironmentName todomcp-env-prod01 `
    -ContainerAppName todomcp-app-prod01 `
    -DomainName todo.h-aa.dk `
    -Location $region
```

## GitHub Actions OIDC (no stored client secret)

```powershell
./scripts/configure-github-oidc.ps1 `
    -RepoOwner <you> -RepoName microsoft-todo-mcp `
    -SubscriptionId <sub-id> `
    -ResourceGroupName rg-microsoft-todo-mcp `
    -ContainerAppName todomcp-app-prod01 `
    -Region $region
```

This creates a dedicated deployment-only app registration scoped to
`Contributor` on just the one resource group, with a federated credential
for `repo:<owner>/<repo>:environment:production` — no client secret is ever
stored in GitHub.

## Cost estimate (approximate, West/North Europe pricing, 2026)

| Resource | Assumption | Est. monthly cost |
|---|---|---|
| Container Apps (consumption) | Scale-to-zero, a few hundred requests/day, minimal always-on | ~5–10 DKK |
| Storage Account (Table + Blob) | A handful of connections, KB-sized cache blobs | < 1 DKK |
| Key Vault | Standard tier, low operation count | ~2 DKK |
| Log Analytics (optional, `enableLogAnalytics=true`) | 30-day retention, low volume | ~5–10 DKK if enabled |
| **Total (Log Analytics off)** | | **~10–15 DKK/month** |
| **Total (Log Analytics on)** | | **~20–25 DKK/month** |

The budget alert (`infra/modules/budget.bicep`, default 25 DKK/month) fires
at 80% and 100% of the configured amount via email — adjust
`-BudgetAmount`/`budgetAmount` per your own pricing region.

## Teardown

```powershell
az group delete --name rg-microsoft-todo-mcp --yes --no-wait
```

Key Vault is soft-delete-protected by default; to fully purge it (freeing
the name for reuse):

```powershell
az keyvault purge --name <key-vault-name> --location <region>
```
