targetScope = 'resourceGroup'

param vaultName string
param location string
param tags object
param tenantId string
param endpointSubnetId string
param vnetName string

// Dedicated signing vault: the runtime never receives a certificate private key.
// Certificate creation/registration and privileged RBAC remain operator steps.
resource vault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  name: vaultName
  location: location
  tags: tags
  properties: {
    tenantId: tenantId
    sku: { family: 'A', name: 'standard' }
    enableRbacAuthorization: true
    enableSoftDelete: true
    softDeleteRetentionInDays: 90
    enablePurgeProtection: true
    publicNetworkAccess: 'Disabled'
    networkAcls: { bypass: 'None', defaultAction: 'Deny', ipRules: [], virtualNetworkRules: [] }
  }
}
resource vnet 'Microsoft.Network/virtualNetworks@2024-05-01' existing = { name: vnetName }
resource dns 'Microsoft.Network/privateDnsZones@2024-06-01' = {
  name: 'privatelink.vaultcore.azure.net'
  location: 'global'
  tags: tags
}
resource link 'Microsoft.Network/privateDnsZones/virtualNetworkLinks@2024-06-01' = {
  parent: dns
  name: '${vnetName}-reporting'
  location: 'global'
  properties: { registrationEnabled: false, virtualNetwork: { id: vnet.id } }
}
resource endpoint 'Microsoft.Network/privateEndpoints@2024-05-01' = {
  name: '${vaultName}-pe'
  location: location
  tags: tags
  properties: {
    subnet: { id: endpointSubnetId }
    privateLinkServiceConnections: [
      { name: 'reporting-vault', properties: { privateLinkServiceId: vault.id, groupIds: [ 'vault' ] } }
    ]
  }
}
resource zone 'Microsoft.Network/privateEndpoints/privateDnsZoneGroups@2024-05-01' = {
  parent: endpoint
  name: 'vault'
  properties: { privateDnsZoneConfigs: [ { name: 'vault', properties: { privateDnsZoneId: dns.id } } ] }
}
output name string = vault.name
output id string = vault.id
