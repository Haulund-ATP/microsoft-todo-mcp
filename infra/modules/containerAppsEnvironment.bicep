@description('Name of the Container Apps environment.')
param name string

@description('Azure region.')
param location string

@description('Log Analytics workspace id, or empty string to skip Container Apps -> Log Analytics wiring.')
param logAnalyticsWorkspaceId string = ''

@description('Log Analytics customer id (workspace GUID), required when logAnalyticsWorkspaceId is set.')
param logAnalyticsCustomerId string = ''

@secure()
@description('Log Analytics shared key, required when logAnalyticsWorkspaceId is set.')
param logAnalyticsSharedKey string = ''

var useLogAnalytics = !empty(logAnalyticsWorkspaceId)

resource environment 'Microsoft.App/managedEnvironments@2024-03-01' = {
  name: name
  location: location
  properties: {
    appLogsConfiguration: useLogAnalytics
      ? {
          destination: 'log-analytics'
          logAnalyticsConfiguration: {
            customerId: logAnalyticsCustomerId
            sharedKey: logAnalyticsSharedKey
          }
        }
      : null
    zoneRedundant: false
  }
}

output id string = environment.id
output defaultDomain string = environment.properties.defaultDomain
