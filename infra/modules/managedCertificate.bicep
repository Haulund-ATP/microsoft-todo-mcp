@description('Name of the Container Apps managed environment the certificate belongs to.')
param environmentName string

@description('Custom domain the certificate covers, e.g. todo.h-aa.dk.')
param domainName string

@description('Azure region (must match the managed environment).')
param location string

// A managed certificate can only be created once domain ownership has been
// validated (CNAME + TXT records in place — see docs/dns.md). Provision
// this AFTER scripts/configure-domain.ps1 has confirmed DNS validation,
// otherwise deployment will fail with a validation error.
resource environment 'Microsoft.App/managedEnvironments@2024-03-01' existing = {
  name: environmentName
}

resource managedCertificate 'Microsoft.App/managedEnvironments/managedCertificates@2024-03-01' = {
  parent: environment
  name: replace(domainName, '.', '-')
  location: location
  properties: {
    subjectName: domainName
    domainControlValidation: 'CNAME'
  }
}

output certificateId string = managedCertificate.id
