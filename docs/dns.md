# DNS and custom domain

## Why domain binding is a separate step from provisioning

Azure Container Apps managed certificates require **domain control
validation** (a CNAME pointing at the environment, plus a TXT record proving
you control the zone) *before* Azure will issue a certificate. If
`infra/main.bicep` tried to bind `todo.h-aa.dk` on the very first deploy —
before those DNS records exist — the deployment would fail. So:

1. `scripts/provision.ps1` deploys first **without** `customDomainName`
   set, giving you a working `*.azurecontainerapps.io` URL immediately.
2. You create the DNS records below, pointing at that environment.
3. `scripts/configure-domain.ps1` requests the managed certificate and
   binds the custom domain, once DNS has propagated.

## Records to create

Run `scripts/configure-domain.ps1` first — it prints the exact values for
your environment — but in general, at your DNS provider for `h-aa.dk`:

| Type | Name | Value |
|---|---|---|
| CNAME | `todo` | `<container-app-name>.<environment-default-domain>` |
| TXT | `asuid.todo` | `<customDomainVerificationId from the Container App>` |

`<environment-default-domain>` looks like
`whitecliff-12345678.westeurope.azurecontainerapps.io`; the exact value is
printed by `az containerapp env show`.

Propagation can take anywhere from a few minutes to a few hours depending
on your provider and existing TTLs — lower the TTL on these records ahead
of time if you want faster iteration.

## Verifying propagation

```powershell
Resolve-DnsName todo.h-aa.dk -Type CNAME
Resolve-DnsName asuid.todo.h-aa.dk -Type TXT
```

Both should resolve before you continue past the confirmation prompt in
`configure-domain.ps1`.

## Certificate lifecycle

Azure Container Apps managed certificates auto-renew as long as the
validating DNS records remain in place — do not remove the `asuid.todo`
TXT record after initial validation, some renewal flows re-check it.

## If you ever change the environment's default domain

(For example, after fully tearing down and re-provisioning.) The
`asuid.todo` TXT record's *value* is tied to the Container App's
`customDomainVerificationId`, which stays stable across redeploys within
the same Container App resource, but changes if you delete and recreate
the Container App itself. Re-run `configure-domain.ps1` in that case.
