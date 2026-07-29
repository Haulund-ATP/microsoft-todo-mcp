#Requires -Version 7.0
<#
.SYNOPSIS
    Builds (optionally), pushes, and deploys a new container image to the
    existing Container App. Intended for both local ad-hoc deploys and as
    the logic .github/workflows/deploy.yml drives via OIDC.

.PARAMETER ResourceGroupName
.PARAMETER ContainerAppName
.PARAMETER Image
    Full image reference to deploy, e.g. ghcr.io/replace-me/microsoft-todo-mcp:sha-abc123.
.PARAMETER Build
    If set, builds and pushes the image locally via `docker build` + `docker push`
    before deploying it (requires Docker and registry login already configured).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$ResourceGroupName,
    [Parameter(Mandatory)] [string]$ContainerAppName,
    [Parameter(Mandatory)] [string]$Image,
    [switch]$Build
)

$ErrorActionPreference = 'Stop'
$repoRoot = Join-Path $PSScriptRoot '..'

if ($Build) {
    Write-Host "== Building image $Image ==" -ForegroundColor Cyan
    docker build -t $Image $repoRoot
    if ($LASTEXITCODE -ne 0) { throw "docker build failed." }

    Write-Host "== Pushing image $Image ==" -ForegroundColor Cyan
    docker push $Image
    if ($LASTEXITCODE -ne 0) { throw "docker push failed." }
}

Write-Host "== Updating Container App '$ContainerAppName' to image $Image ==" -ForegroundColor Cyan
az containerapp update `
    --name $ContainerAppName `
    --resource-group $ResourceGroupName `
    --image $Image `
    --output json | Out-Null

Write-Host "== Waiting for the new revision to become healthy ==" -ForegroundColor Cyan
$fqdn = az containerapp show --name $ContainerAppName --resource-group $ResourceGroupName --query "properties.configuration.ingress.fqdn" --output tsv

$deadline = (Get-Date).AddMinutes(3)
$healthy = $false
while ((Get-Date) -lt $deadline) {
    try {
        $response = Invoke-WebRequest -Uri "https://$fqdn/health" -UseBasicParsing -TimeoutSec 5
        if ($response.StatusCode -eq 200) { $healthy = $true; break }
    } catch {
        Start-Sleep -Seconds 5
    }
}

if (-not $healthy) {
    throw "Deployment did not become healthy within 3 minutes. Check 'az containerapp logs show --name $ContainerAppName --resource-group $ResourceGroupName'."
}

Write-Host "== Deployed successfully. https://$fqdn/health is responding. ==" -ForegroundColor Green
