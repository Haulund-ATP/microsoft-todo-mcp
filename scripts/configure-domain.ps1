#Requires -Version 7.0
<#
.SYNOPSIS
    Walks through binding a custom domain (e.g. todo.h-aa.dk) to the
    Container App: prints the DNS records you need to create, waits for
    them to be visible, then creates the managed certificate and binds
    the domain via infra/modules/managedCertificate.bicep + a follow-up
    `az containerapp hostname` call.

.DESCRIPTION
    This is intentionally a two-phase process (see docs/dns.md):
      1. Deploy without a custom domain (provision.ps1 default).
      2. Create the CNAME + TXT (asuid) records at your DNS provider,
         confirm they resolve, THEN run this script to request the
         managed certificate and bind the domain — Azure will not issue
         a certificate for a domain it cannot yet validate.

.PARAMETER ResourceGroupName
.PARAMETER ContainerAppEnvironmentName
.PARAMETER ContainerAppName
.PARAMETER DomainName
    e.g. todo.h-aa.dk
.PARAMETER Location
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$ResourceGroupName,
    [Parameter(Mandatory)] [string]$ContainerAppEnvironmentName,
    [Parameter(Mandatory)] [string]$ContainerAppName,
    [Parameter(Mandatory)] [string]$DomainName,
    [Parameter(Mandatory)] [string]$Location
)

$ErrorActionPreference = 'Stop'
$repoRoot = Join-Path $PSScriptRoot '..'

$envInfo = az containerapp env show --name $ContainerAppEnvironmentName --resource-group $ResourceGroupName --output json | ConvertFrom-Json
$defaultDomain = $envInfo.properties.defaultDomain
$verificationId = az containerapp show --name $ContainerAppName --resource-group $ResourceGroupName --query "properties.customDomainVerificationId" --output tsv

Write-Host "== Create these DNS records at your DNS provider, then re-run this script ==" -ForegroundColor Cyan
Write-Host ""
Write-Host "  CNAME  $DomainName  ->  $ContainerAppName.$defaultDomain"
Write-Host "  TXT    asuid.$DomainName  ->  $verificationId"
Write-Host ""
$confirm = Read-Host "Have these records been created and propagated? (y/N)"
if ($confirm -ne 'y') {
    Write-Host "Aborting. Re-run this script once DNS has propagated (can take up to a few hours)." -ForegroundColor Yellow
    exit 0
}

Write-Host "`n== Requesting managed certificate for $DomainName ==" -ForegroundColor Cyan
az deployment group create `
    --resource-group $ResourceGroupName `
    --template-file (Join-Path $repoRoot 'infra' 'modules' 'managedCertificate.bicep') `
    --parameters environmentName=$ContainerAppEnvironmentName domainName=$DomainName location=$Location `
    --output none

Write-Host "== Binding hostname to the Container App ==" -ForegroundColor Cyan
az containerapp hostname add --hostname $DomainName --name $ContainerAppName --resource-group $ResourceGroupName --output none
az containerapp hostname bind --hostname $DomainName --name $ContainerAppName --resource-group $ResourceGroupName --environment $ContainerAppEnvironmentName --output none

Write-Host "`n== Done. https://$DomainName should now serve the app once the certificate finishes issuing (a few minutes). ==" -ForegroundColor Green
