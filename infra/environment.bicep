targetScope = 'subscription'

param resourceGroupName string = 'rg-copilot-cost-demo-uks'
param prefix string = 'copilotcost'
param location string = 'uksouth'
param tenantId string
param spaClientId string
param apiClientId string

resource group 'Microsoft.Resources/resourceGroups@2024-03-01' = {
  name: resourceGroupName
  location: location
  tags: { application: 'copilot-cost-demo', environment: 'demo' }
}
module identities './bootstrap.bicep' = {
  name: 'github-identities'
  scope: group
  params: { prefix: prefix, location: location }
}
module resources './main.bicep' = {
  name: 'application-resources'
  scope: group
  params: {
    prefix: prefix
    location: location
    tenantId: tenantId
    spaClientId: spaClientId
    apiClientId: apiClientId
  }
}
output resources object = resources.outputs
output identities object = identities.outputs
