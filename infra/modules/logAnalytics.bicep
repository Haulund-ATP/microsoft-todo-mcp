@description('Name of the Log Analytics workspace.')
param name string

@description('Azure region.')
param location string

@description('Retention in days. Kept short to control cost on a low-traffic hobby deployment.')
param retentionInDays int = 30

resource workspace 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: name
  location: location
  properties: {
    sku: {
      name: 'PerGB2018'
    }
    retentionInDays: retentionInDays
  }
}

output id string = workspace.id
output customerId string = workspace.properties.customerId
