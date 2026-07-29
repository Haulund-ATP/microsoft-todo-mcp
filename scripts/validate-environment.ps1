#Requires -Version 7.0
<#
.SYNOPSIS
    Validates local tooling and required environment/config values before
    running provision.ps1 or deploy.ps1.

.DESCRIPTION
    Fails fast with clear, actionable messages rather than letting a
    provisioning run partially fail deep into an ARM/Bicep deployment.
#>
[CmdletBinding()]
param(
    [string]$EnvFilePath = (Join-Path $PSScriptRoot '..' '.env')
)

$ErrorActionPreference = 'Stop'
$failures = New-Object System.Collections.Generic.List[string]

function Test-Tool {
    param([string]$Name, [string]$InstallHint)
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        $failures.Add("Missing required tool '$Name'. $InstallHint")
    } else {
        Write-Host "[ok] $Name found: $((Get-Command $Name).Source)" -ForegroundColor Green
    }
}

Write-Host "== Checking required CLI tools ==" -ForegroundColor Cyan
Test-Tool -Name 'az' -InstallHint 'https://learn.microsoft.com/cli/azure/install-azure-cli'
Test-Tool -Name 'node' -InstallHint 'Install Node.js 22+: https://nodejs.org'
Test-Tool -Name 'npm' -InstallHint 'Ships with Node.js.'
Test-Tool -Name 'docker' -InstallHint 'https://docs.docker.com/get-docker/ (only required for local image builds/tests)'
Test-Tool -Name 'gh' -InstallHint 'https://cli.github.com/ (only required for configure-github-oidc.ps1)'

Write-Host "`n== Checking Bicep ==" -ForegroundColor Cyan
try {
    $bicepVersion = az bicep version 2>&1
    Write-Host "[ok] $bicepVersion" -ForegroundColor Green
} catch {
    $failures.Add("Bicep CLI not available via 'az bicep'. Run: az bicep install")
}

Write-Host "`n== Checking Node.js version ==" -ForegroundColor Cyan
$nodeVersion = (node --version) -replace '^v', ''
$majorVersion = [int]($nodeVersion.Split('.')[0])
if ($majorVersion -lt 22) {
    $failures.Add("Node.js 22+ is required; found $nodeVersion.")
} else {
    Write-Host "[ok] Node.js $nodeVersion" -ForegroundColor Green
}

Write-Host "`n== Checking .env / .env.example ==" -ForegroundColor Cyan
$exampleFile = Join-Path $PSScriptRoot '..' '.env.example'
if (-not (Test-Path $exampleFile)) {
    $failures.Add(".env.example is missing from the repo root; this indicates a broken checkout.")
}
if (-not (Test-Path $EnvFilePath)) {
    Write-Host "[warn] No .env found at $EnvFilePath (expected for CI/production; required for local dev)." -ForegroundColor Yellow
} else {
    Write-Host "[ok] .env found at $EnvFilePath" -ForegroundColor Green
    $envContent = Get-Content $EnvFilePath -Raw
    foreach ($placeholder in @('replace-me')) {
        if ($envContent -match [regex]::Escape($placeholder)) {
            Write-Host "[warn] .env still contains '$placeholder' placeholders — fill in real values before running the app." -ForegroundColor Yellow
        }
    }
}

Write-Host "`n== Checking Azure login state ==" -ForegroundColor Cyan
$account = az account show --output json 2>$null | ConvertFrom-Json
if (-not $account) {
    $failures.Add("Not logged in to Azure CLI. Run ./scripts/login.ps1 first.")
} else {
    Write-Host "[ok] Logged in as $($account.user.name) on subscription $($account.name)" -ForegroundColor Green
}

if ($failures.Count -gt 0) {
    Write-Host "`n== Validation FAILED ==" -ForegroundColor Red
    foreach ($f in $failures) { Write-Host " - $f" -ForegroundColor Red }
    exit 1
}

Write-Host "`n== Validation passed ==" -ForegroundColor Green
