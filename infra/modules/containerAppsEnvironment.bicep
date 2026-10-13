@description('Name of the Container Apps environment.')
param name string

@description('Azure region.')
param location string

@description('Name of an existing Log Analytics workspace in this resource group, or empty string to skip Container Apps -> Log Analytics wiring.')
param logAnalyticsWorkspaceName string = ''

var useLogAnalytics = !empty(logAnalyticsWorkspaceName)

// The shared key is read here with listKeys() at deploy time rather than
// passed in as a parameter, so it never flows through parameter files or
// deployment history in plaintext.
resource workspace 'Microsoft.OperationalInsights/workspaces@2023-09-01' existing = if (useLogAnalytics) {
  name: useLogAnalytics ? logAnalyticsWorkspaceName : 'unused'
}

resource environment 'Microsoft.App/managedEnvironments@2024-03-01' = {
  name: name
  location: location
  properties: {
    appLogsConfiguration: useLogAnalytics
      ? {
          destination: 'log-analytics'
          logAnalyticsConfiguration: {
            customerId: workspace.properties.customerId
            sharedKey: workspace.listKeys().primarySharedKey
          }
        }
      : null
    zoneRedundant: false
  }
}

output id string = environment.id
output defaultDomain string = environment.properties.defaultDomain
