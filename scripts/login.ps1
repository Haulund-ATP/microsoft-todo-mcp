#Requires -Version 7.0
<#
.SYNOPSIS
    Logs in to Azure CLI and selects the subscription to deploy into.

.DESCRIPTION
    Thin wrapper around `az login` that also verifies the CLI version and
    lets the caller pick a subscription interactively (or pass -SubscriptionId
    non-interactively for CI use, though CI should use OIDC instead — see
    configure-github-oidc.ps1 and .github/workflows/deploy.yml).

.PARAMETER SubscriptionId
    Optional. If provided, skips the interactive subscription picker.

.EXAMPLE
    ./scripts/login.ps1
.EXAMPLE
    ./scripts/login.ps1 -SubscriptionId 00000000-0000-0000-0000-000000000000
#>
[CmdletBinding()]
param(
    [string]$SubscriptionId
)

$ErrorActionPreference = 'Stop'

function Assert-AzCli {
    if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
        throw "Azure CLI ('az') was not found on PATH. Install it: https://learn.microsoft.com/cli/azure/install-azure-cli"
    }
    $versionJson = az version --output json | ConvertFrom-Json
    Write-Host "Azure CLI version: $($versionJson.'azure-cli')" -ForegroundColor Cyan
}

Assert-AzCli

$account = az account show --output json 2>$null | ConvertFrom-Json
if (-not $account) {
    Write-Host "Not logged in. Launching 'az login'..." -ForegroundColor Yellow
    az login --output none
}

if ($SubscriptionId) {
    az account set --subscription $SubscriptionId
} else {
    $subs = az account list --output json | ConvertFrom-Json
    if ($subs.Count -eq 0) {
        throw "No subscriptions visible to this account."
    } elseif ($subs.Count -eq 1) {
        az account set --subscription $subs[0].id
    } else {
        Write-Host "Available subscriptions:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $subs.Count; $i++) {
            Write-Host "  [$i] $($subs[$i].name) ($($subs[$i].id))"
        }
        $choice = Read-Host "Select subscription index"
        az account set --subscription $subs[[int]$choice].id
    }
}

$current = az account show --output json | ConvertFrom-Json
Write-Host "Using subscription: $($current.name) ($($current.id))" -ForegroundColor Green
