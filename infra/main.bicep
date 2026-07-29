targetScope = 'resourceGroup'

@description('Short, unique-ish suffix used to keep globally-unique resource names apart across environments, e.g. "prod" or a random 6-char string.')
@minLength(3)
@maxLength(10)
param resourceSuffix string

@description('Azure region for all resources. Left as a plain string (not a hardcoded fabricated value) — pick with scripts/select-region.ps1.')
param location string

@description('Public base URL of the deployment, e.g. https://todo.h-aa.dk.')
param publicBaseUrl string

@description('Custom domain name to bind to the Container App, e.g. todo.h-aa.dk. Leave empty to deploy without a custom domain first (recommended for the first deploy, before DNS/cert validation).')
param customDomainName string = ''

@description('Upstream Microsoft Entra multi-tenant + personal-account app registration client id used for Graph access. Not a secret, but still not a fabricated value — set this after app registration.')
param azureClientId string

@description('Tenant id of the Azure subscription/directory this is deployed into (used for display/tagging only; the Graph OAuth authority is always /common regardless).')
param azureTenantId string

@description('Container image reference to deploy, e.g. ghcr.io/replace-me/microsoft-todo-mcp:sha-xxxxx.')
param containerImage string

@description('Container registry server (leave empty for anonymous/public pull).')
param registryServer string = ''

@description('Container registry username (leave empty for anonymous/public pull).')
param registryUsername string = ''

@secure()
@description('Container registry password/PAT (leave empty for anonymous/public pull).')
param registryPassword string = ''

@description('Email address for budget alert notifications.')
param budgetNotificationEmail string

@description('Monthly budget amount in the subscription\'s billing currency.')
param budgetAmount int = 25

@description('Deploy an (optional) Log Analytics workspace and wire it to the Container Apps environment.')
param enableLogAnalytics bool = false

@description('Owner identity claim name used to gate the /accounts admin page, e.g. "oid".')
param adminOwnerClaimName string = 'oid'

var namePrefix = 'todomcp'
var managedIdentityName = '${namePrefix}-id-${resourceSuffix}'
var keyVaultName = take('${namePrefix}kv${resourceSuffix}', 24)
var storageAccountName = take(toLower('${namePrefix}st${resourceSuffix}'), 24)
var logAnalyticsName = '${namePrefix}-log-${resourceSuffix}'
var environmentName = '${namePrefix}-env-${resourceSuffix}'
var containerAppName = '${namePrefix}-app-${resourceSuffix}'
var budgetName = '${namePrefix}-budget-${resourceSuffix}'

module identity 'modules/managedIdentity.bicep' = {
  name: 'managedIdentity'
  params: {
    name: managedIdentityName
    location: location
  }
}

module keyVault 'modules/keyVault.bicep' = {
  name: 'keyVault'
  params: {
    name: keyVaultName
    location: location
    tenantId: azureTenantId
    managedIdentityPrincipalId: identity.outputs.principalId
  }
}

module storage 'modules/storage.bicep' = {
  name: 'storage'
  params: {
    name: storageAccountName
    location: location
    managedIdentityPrincipalId: identity.outputs.principalId
  }
}

module logAnalytics 'modules/logAnalytics.bicep' = if (enableLogAnalytics) {
  name: 'logAnalytics'
  params: {
    name: logAnalyticsName
    location: location
  }
}

module environment 'modules/containerAppsEnvironment.bicep' = {
  name: 'environment'
  params: {
    name: environmentName
    location: location
    logAnalyticsWorkspaceId: enableLogAnalytics ? logAnalytics.outputs.id : ''
    logAnalyticsCustomerId: enableLogAnalytics ? logAnalytics.outputs.customerId : ''
    // Shared key is intentionally NOT wired here to avoid plaintext key flow
    // through parameters; see docs/operations.md for the documented manual
    // step (or extend with a listKeys() reference at deploy time) if you
    // enable Log Analytics.
    logAnalyticsSharedKey: ''
  }
}

module containerApp 'modules/containerApp.bicep' = {
  name: 'containerApp'
  params: {
    name: containerAppName
    location: location
    environmentId: environment.outputs.id
    managedIdentityId: identity.outputs.id
    managedIdentityClientId: identity.outputs.clientId
    image: containerImage
    publicBaseUrl: publicBaseUrl
    customDomainName: customDomainName
    keyVaultUri: keyVault.outputs.uri
    storageAccountName: storage.outputs.name
    azureClientId: azureClientId
    azureTenantId: azureTenantId
    adminOwnerClaimName: adminOwnerClaimName
    registryServer: registryServer
    registryUsername: registryUsername
    registryPassword: registryPassword
    logAnalyticsWorkspaceId: enableLogAnalytics ? logAnalytics.outputs.id : ''
  }
}

module budget 'modules/budget.bicep' = {
  name: 'budget'
  params: {
    name: budgetName
    amount: budgetAmount
    notificationEmail: budgetNotificationEmail
  }
}

output containerAppFqdn string = containerApp.outputs.fqdn
output keyVaultUri string = keyVault.outputs.uri
output storageAccountName string = storage.outputs.name
output managedIdentityClientId string = identity.outputs.clientId
output managedIdentityPrincipalId string = identity.outputs.principalId
