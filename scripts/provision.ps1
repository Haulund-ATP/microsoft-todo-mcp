#Requires -Version 7.0
<#
.SYNOPSIS
    First-time provisioning: creates the resource group, generates the
    MCP OAuth signing key + cookie secrets, seeds Key Vault, and deploys
    infra/main.bicep (without a custom domain — add that afterwards via
    configure-domain.ps1 once DNS is validated).

.DESCRIPTION
    Idempotent-ish: safe to re-run. Secret generation is skipped if the
    corresponding Key Vault secret already exists, so re-running this
    script does not rotate secrets by accident.

.PARAMETER ResourceGroupName
.PARAMETER Location
    Azure region, e.g. from ./scripts/select-region.ps1.
.PARAMETER ResourceSuffix
    Short suffix used in globally-unique resource names (Key Vault, Storage).
.PARAMETER AzureClientId
    The upstream Microsoft Entra app registration client id (Graph access).
.PARAMETER AzureTenantId
.PARAMETER PublicBaseUrl
.PARAMETER BudgetNotificationEmail
.PARAMETER ContainerImage
    Initial image to deploy, e.g. a placeholder or the first CI-built tag.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$ResourceGroupName,
    [Parameter(Mandatory)] [string]$Location,
    [Parameter(Mandatory)] [string]$ResourceSuffix,
    [Parameter(Mandatory)] [string]$AzureClientId,
    [Parameter(Mandatory)] [string]$AzureTenantId,
    [Parameter(Mandatory)] [string]$PublicBaseUrl,
    [Parameter(Mandatory)] [string]$BudgetNotificationEmail,
    [string]$ContainerImage = 'mcr.microsoft.com/k8se/quickstart:latest',
    [string]$AdminOwnerClaimName = 'oid'
)

$ErrorActionPreference = 'Stop'
$repoRoot = Join-Path $PSScriptRoot '..'

Write-Host "== Ensuring resource group $ResourceGroupName in $Location ==" -ForegroundColor Cyan
az group create --name $ResourceGroupName --location $Location --output none

$keyVaultName = ('todomcpkv' + $ResourceSuffix).Substring(0, [Math]::Min(24, ('todomcpkv' + $ResourceSuffix).Length))

Write-Host "`n== Deploying infrastructure (no custom domain yet) ==" -ForegroundColor Cyan
$deployOutput = az deployment group create `
    --resource-group $ResourceGroupName `
    --template-file (Join-Path $repoRoot 'infra' 'main.bicep') `
    --parameters `
        resourceSuffix=$ResourceSuffix `
        location=$Location `
        publicBaseUrl=$PublicBaseUrl `
        azureClientId=$AzureClientId `
        azureTenantId=$AzureTenantId `
        containerImage=$ContainerImage `
        budgetNotificationEmail=$BudgetNotificationEmail `
        adminOwnerClaimName=$AdminOwnerClaimName `
    --output json | ConvertFrom-Json

$outputs = $deployOutput.properties.outputs
Write-Host "Container App FQDN: $($outputs.containerAppFqdn.value)" -ForegroundColor Green
Write-Host "Key Vault URI:      $($outputs.keyVaultUri.value)" -ForegroundColor Green

function Set-SecretIfMissing {
    param([string]$VaultName, [string]$SecretName, [scriptblock]$Generator)
    $existing = az keyvault secret show --vault-name $VaultName --name $SecretName --output json 2>$null
    if ($existing) {
        Write-Host "[skip] Secret '$SecretName' already exists." -ForegroundColor Yellow
        return
    }
    $value = & $Generator
    az keyvault secret set --vault-name $VaultName --name $SecretName --value $value --output none
    Write-Host "[ok] Set secret '$SecretName'." -ForegroundColor Green
}

function New-RandomBase64Url {
    param([int]$Bytes = 32)
    $buffer = [byte[]]::new($Bytes)
    [System.Security.Cryptography.RandomNumberGenerator]::Fill($buffer)
    return [Convert]::ToBase64String($buffer).TrimEnd('=').Replace('+', '-').Replace('/', '_')
}

function New-EcP256JwkJson {
    # Generates a P-256 (ES256) key pair and serializes it as a JWK with a
    # stable kid, matching what src/oauth/signingKeys.ts expects.
    $ecdsa = [System.Security.Cryptography.ECDsa]::Create([System.Security.Cryptography.ECCurve+NamedCurves]::nistP256)
    $params = $ecdsa.ExportParameters($true)
    $b64url = { param($bytes) [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
    $kid = New-RandomBase64Url -Bytes 8
    $jwk = [ordered]@{
        kty = 'EC'
        crv = 'P-256'
        x   = (& $b64url $params.Q.X)
        y   = (& $b64url $params.Q.Y)
        d   = (& $b64url $params.D)
        kid = $kid
    }
    return ($jwk | ConvertTo-Json -Compress)
}

Write-Host "`n== Seeding Key Vault secrets (skipped if already present) ==" -ForegroundColor Cyan
Set-SecretIfMissing -VaultName $keyVaultName -SecretName 'mcp-oauth-signing-key' -Generator { New-EcP256JwkJson }
Set-SecretIfMissing -VaultName $keyVaultName -SecretName 'session-cookie-secret' -Generator { New-RandomBase64Url -Bytes 32 }
Set-SecretIfMissing -VaultName $keyVaultName -SecretName 'csrf-cookie-secret' -Generator { New-RandomBase64Url -Bytes 32 }
Set-SecretIfMissing -VaultName $keyVaultName -SecretName 'msal-cache-encryption-key' -Generator { New-RandomBase64Url -Bytes 32 }
Set-SecretIfMissing -VaultName $keyVaultName -SecretName 'graph-client-secret' -Generator {
    Write-Host "graph-client-secret has no safe auto-generated value — create a client secret on the Entra app registration and paste it now." -ForegroundColor Yellow
    Read-Host -AsSecureString "Paste the Graph app registration client secret" | ConvertFrom-SecureString -AsPlainText
}
Set-SecretIfMissing -VaultName $keyVaultName -SecretName 'admin-owner-claim-value' -Generator {
    Write-Host "admin-owner-claim-value should be the deployment owner's stable Microsoft identity claim (see docs/operations.md for how to find it)." -ForegroundColor Yellow
    Read-Host "Paste the owner's oid/sub claim value"
}

Write-Host "`n== Done. Next: run ./scripts/deploy.ps1 to push a real image, then ./scripts/configure-domain.ps1 for the custom domain. ==" -ForegroundColor Green
