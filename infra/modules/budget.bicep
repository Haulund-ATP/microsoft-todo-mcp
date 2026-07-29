@description('Name of the budget resource.')
param name string

@description('Monthly budget amount. Billed in the subscription\'s native currency (see docs/azure-deployment.md for the DKK 25/month cost estimate).')
param amount int = 25

@description('Email address to notify at 80% and 100% of budget.')
param notificationEmail string

@description('Start date for the budget\'s monthly time period, defaulting to the first of the current month.')
param startDate string = '${utcNow('yyyy-MM')}-01'

// Deployed at resource-group scope (this module is included from main.bicep,
// which itself targets the resource group), so no explicit filter is
// needed — the budget automatically applies to this resource group.
resource budget 'Microsoft.Consumption/budgets@2023-11-01' = {
  name: name
  properties: {
    category: 'Cost'
    amount: amount
    timeGrain: 'Monthly'
    timePeriod: {
      startDate: startDate
    }
    notifications: {
      actual_80_percent: {
        enabled: true
        operator: 'GreaterThanOrEqualTo'
        threshold: 80
        contactEmails: [
          notificationEmail
        ]
        thresholdType: 'Actual'
      }
      actual_100_percent: {
        enabled: true
        operator: 'GreaterThanOrEqualTo'
        threshold: 100
        contactEmails: [
          notificationEmail
        ]
        thresholdType: 'Actual'
      }
    }
  }
}
