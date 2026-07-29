#Requires -Version 7.0
<#
.SYNOPSIS
    Configures GitHub Actions OIDC federated credentials against an Entra
    app registration (a *deployment* identity, distinct from the app
    registration used for upstream Graph OAuth) and sets the GitHub
    Actions repository VARIABLES (not secrets) that .github/workflows/deploy.yml
    reads.

.DESCRIPTION
    Per the security model documented in docs/security.md, deploy.yml uses
    `permissions: id-token: write` + `azure/login@v2` with a federated
    credential — no client secret is stored in GitHub. Tenant id,
    subscription id, client id, resource group, container app name, and
    region are not secrets (they're not sufficient on their own to access
    anything), so they are stored as repository VARIABLES.

.PARAMETER RepoOwner
.PARAMETER RepoName
.PARAMETER AppRegistrationName
    Display name for the deployment-only app registration to create (or reuse).
.PARAMETER SubscriptionId
.PARAMETER ResourceGroupName
.PARAMETER ContainerAppName
.PARAMETER Region
.PARAMETER Environment
    GitHub Environment name that deploy.yml targets. Default: production.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$RepoOwner,
    [Parameter(Mandatory)] [string]$RepoName,
    [string]$AppRegistrationName = 'microsoft-todo-mcp-deploy',
    [Parameter(Mandatory)] [string]$SubscriptionId,
    [Parameter(Mandatory)] [string]$ResourceGroupName,
    [Parameter(Mandatory)] [string]$ContainerAppName,
    [Parameter(Mandatory)] [string]$Region,
    [string]$Environment = 'production'
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "GitHub CLI ('gh') is required. Install: https://cli.github.com/"
}

Write-Host "== Creating (or reusing) app registration '$AppRegistrationName' ==" -ForegroundColor Cyan
$existingApp = az ad app list --display-name $AppRegistrationName --output json | ConvertFrom-Json
if ($existingApp.Count -gt 0) {
    $app = $existingApp[0]
    Write-Host "[skip] App registration already exists: $($app.appId)" -ForegroundColor Yellow
} else {
    $app = az ad app create --display-name $AppRegistrationName --output json | ConvertFrom-Json
    az ad sp create --id $app.appId --output none
    Write-Host "[ok] Created app registration: $($app.appId)" -ForegroundColor Green
}

Write-Host "`n== Assigning Contributor on the resource group (scoped, not subscription-wide) ==" -ForegroundColor Cyan
$scope = "/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroupName"
az role assignment create `
    --assignee $app.appId `
    --role "Contributor" `
    --scope $scope `
    --output none

Write-Host "`n== Creating federated credential for GitHub Actions OIDC ==" -ForegroundColor Cyan
$subject = "repo:$($RepoOwner)/$($RepoName):environment:$Environment"
$federatedCredentialJson = @{
    name = "github-actions-$Environment"
    issuer = "https://token.actions.githubusercontent.com"
    subject = $subject
    audiences = @("api://AzureADTokenExchange")
} | ConvertTo-Json -Compress

$tempFile = New-TemporaryFile
Set-Content -Path $tempFile -Value $federatedCredentialJson -NoNewline
az ad app federated-credential create --id $app.appId --parameters "@$tempFile" --output none
Remove-Item $tempFile

Write-Host "`n== Setting GitHub repository variables (not secrets) ==" -ForegroundColor Cyan
$tenantId = (az account show --query tenantId --output tsv)

gh variable set AZURE_TENANT_ID --repo "$RepoOwner/$RepoName" --body $tenantId
gh variable set AZURE_SUBSCRIPTION_ID --repo "$RepoOwner/$RepoName" --body $SubscriptionId
gh variable set AZURE_CLIENT_ID --repo "$RepoOwner/$RepoName" --body $app.appId
gh variable set AZURE_RESOURCE_GROUP --repo "$RepoOwner/$RepoName" --body $ResourceGroupName
gh variable set CONTAINER_APP_NAME --repo "$RepoOwner/$RepoName" --body $ContainerAppName
gh variable set AZURE_REGION --repo "$RepoOwner/$RepoName" --body $Region

Write-Host "`n== Done. deploy.yml can now authenticate via OIDC with no stored client secret. ==" -ForegroundColor Green
