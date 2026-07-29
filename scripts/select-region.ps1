#Requires -Version 7.0
<#
.SYNOPSIS
    Lists Azure regions that support Container Apps consumption plan and
    lets the caller pick one, without hardcoding a fabricated default.

.DESCRIPTION
    This project deliberately does not hardcode a region in Bicep
    parameters — you choose it explicitly here (or pass -Region) based on
    your latency/compliance preferences and current Container Apps
    availability, which changes over time.

.PARAMETER Region
    Optional. If supplied, validates it's a real Azure region name and
    returns it without prompting.

.EXAMPLE
    ./scripts/select-region.ps1
.EXAMPLE
    ./scripts/select-region.ps1 -Region westeurope
#>
[CmdletBinding()]
param(
    [string]$Region
)

$ErrorActionPreference = 'Stop'

$allRegions = az account list-locations --output json | ConvertFrom-Json |
    Where-Object { $_.metadata.regionType -eq 'Physical' } |
    Sort-Object -Property displayName

if ($Region) {
    $match = $allRegions | Where-Object { $_.name -eq $Region }
    if (-not $match) {
        throw "'$Region' is not a recognized Azure region name (expected e.g. 'westeurope', 'northeurope', 'swedencentral')."
    }
    Write-Output $match.name
    return
}

Write-Host "Common low-latency choices for a Denmark-based deployment: swedencentral, westeurope, northeurope, germanywestcentral." -ForegroundColor Cyan
Write-Host ""
$displayed = $allRegions | Select-Object -First 40
for ($i = 0; $i -lt $displayed.Count; $i++) {
    Write-Host ("  [{0,2}] {1} ({2})" -f $i, $displayed[$i].displayName, $displayed[$i].name)
}
$choice = Read-Host "`nSelect a region index (or type a region name directly)"

if ($choice -match '^\d+$') {
    Write-Output $displayed[[int]$choice].name
} else {
    $match = $allRegions | Where-Object { $_.name -eq $choice }
    if (-not $match) { throw "'$choice' is not a recognized Azure region name." }
    Write-Output $match.name
}
