@description('Name of the Container App.')
param name string

@description('Azure region.')
param location string

@description('Resource id of the Container Apps managed environment.')
param environmentId string

@description('Resource id of the user-assigned managed identity to attach to the app.')
param managedIdentityId string

@description('Client id of the user-assigned managed identity (passed to the app as AZURE_CLIENT_ID is the Graph app; this is separate — used for DefaultAzureCredential selection).')
param managedIdentityClientId string

@description('Container image reference, e.g. ghcr.io/org/microsoft-todo-mcp:sha.')
param image string

@description('Target port the app listens on inside the container.')
param targetPort int = 3000

@description('Public base URL of the deployment, e.g. https://todo.h-aa.dk.')
param publicBaseUrl string

@description('Custom domain name to bind (e.g. todo.h-aa.dk), or empty to skip custom domain binding on this deployment.')
param customDomainName string = ''

@description('Key Vault URI, used both as an env var and as the source for Key-Vault-backed secrets below.')
param keyVaultUri string

@description('Storage account name (Table + Blob), accessed via managed identity.')
param storageAccountName string

@description('Upstream Microsoft Entra multi-tenant app registration client id (Graph).')
param azureClientId string

@description('Upstream Microsoft Entra tenant id used for the /common authority display purposes only (actual authority is always /common in code).')
param azureTenantId string

@description('Owner claim name, e.g. oid.')
param adminOwnerClaimName string = 'oid'

@description('Container registry server, e.g. ghcr.io. Leave empty for public/anonymous pull.')
param registryServer string = ''

@description('Container registry username. Leave empty for public/anonymous pull.')
param registryUsername string = ''

@secure()
@description('Container registry password/PAT. Leave empty for public/anonymous pull.')
param registryPassword string = ''

@description('Optional Log Analytics workspace resource id, used only for output/documentation purposes here.')
param logAnalyticsWorkspaceId string = ''

var hasRegistryAuth = !empty(registryServer) && !empty(registryUsername)
var hasCustomDomain = !empty(customDomainName)

resource containerApp 'Microsoft.App/containerApps@2024-03-01' = {
  name: name
  location: location
  tags: !empty(logAnalyticsWorkspaceId) ? { logAnalyticsWorkspaceId: logAnalyticsWorkspaceId } : {}
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${managedIdentityId}': {}
    }
  }
  properties: {
    managedEnvironmentId: environmentId
    configuration: {
      activeRevisionsMode: 'Single'
      ingress: {
        external: true
        targetPort: targetPort
        transport: 'auto'
        allowInsecure: false
        customDomains: hasCustomDomain
          ? [
              {
                name: customDomainName
                bindingType: 'SniEnabled'
                certificateId: null // bound separately once the managed certificate resource exists; see docs/dns.md
              }
            ]
          : []
      }
      registries: hasRegistryAuth
        ? [
            {
              server: registryServer
              username: registryUsername
              passwordSecretRef: 'registry-password'
            }
          ]
        : []
      secrets: concat(
        hasRegistryAuth
          ? [
              {
                name: 'registry-password'
                value: registryPassword
              }
            ]
          : [],
        [
          {
            name: 'admin-owner-claim-value'
            keyVaultUrl: '${keyVaultUri}secrets/admin-owner-claim-value'
            identity: managedIdentityId
          }
          {
            name: 'session-cookie-secret'
            keyVaultUrl: '${keyVaultUri}secrets/session-cookie-secret'
            identity: managedIdentityId
          }
          {
            name: 'csrf-cookie-secret'
            keyVaultUrl: '${keyVaultUri}secrets/csrf-cookie-secret'
            identity: managedIdentityId
          }
        ]
      )
    }
    template: {
      revisionSuffix: uniqueString(image)
      scale: {
        minReplicas: 0
        maxReplicas: 1
        rules: [
          {
            name: 'http-scale'
            http: {
              metadata: {
                concurrentRequests: '20'
              }
            }
          }
        ]
      }
      containers: [
        {
          name: 'microsoft-todo-mcp'
          image: image
          resources: {
            cpu: json('0.25')
            memory: '0.5Gi'
          }
          env: [
            { name: 'NODE_ENV', value: 'production' }
            { name: 'PORT', value: string(targetPort) }
            { name: 'PUBLIC_BASE_URL', value: publicBaseUrl }
            { name: 'AZURE_TENANT_ID', value: azureTenantId }
            { name: 'AZURE_CLIENT_ID', value: azureClientId }
            { name: 'AZURE_KEY_VAULT_URI', value: keyVaultUri }
            { name: 'AZURE_STORAGE_ACCOUNT_NAME', value: storageAccountName }
            { name: 'AZURE_CLIENT_ID_MANAGED_IDENTITY', value: managedIdentityClientId }
            { name: 'ADMIN_OWNER_CLAIM_NAME', value: adminOwnerClaimName }
            { name: 'DEFAULT_TIMEZONE', value: 'Europe/Copenhagen' }
            { name: 'MCP_OAUTH_ISSUER', value: publicBaseUrl }
            { name: 'MCP_OAUTH_SIGNING_KEY_NAME', value: 'mcp-oauth-signing-key' }
            { name: 'ADMIN_OWNER_CLAIM_VALUE', secretRef: 'admin-owner-claim-value' }
            { name: 'SESSION_COOKIE_SECRET', secretRef: 'session-cookie-secret' }
            { name: 'CSRF_COOKIE_SECRET', secretRef: 'csrf-cookie-secret' }
          ]
          probes: [
            {
              type: 'Liveness'
              httpGet: { path: '/health', port: targetPort }
              initialDelaySeconds: 5
              periodSeconds: 30
            }
            {
              type: 'Readiness'
              httpGet: { path: '/ready', port: targetPort }
              initialDelaySeconds: 5
              periodSeconds: 15
            }
            {
              type: 'Startup'
              httpGet: { path: '/health', port: targetPort }
              initialDelaySeconds: 2
              periodSeconds: 5
              failureThreshold: 10
            }
          ]
        }
      ]
    }
  }
}

output fqdn string = containerApp.properties.configuration.ingress.fqdn
output name string = containerApp.name
