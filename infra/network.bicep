targetScope = 'resourceGroup'

param prefix string
param suffix string
param location string
param tags object
param hostStorageName string
param businessStorageName string
@description('Isolated demo address space; review for overlap before adding any peering or VPN.')
param addressPrefix string = '10.84.0.0/24'

resource host 'Microsoft.Storage/storageAccounts@2023-05-01' existing = { name: hostStorageName }
resource business 'Microsoft.Storage/storageAccounts@2023-05-01' existing = { name: businessStorageName }

resource integrationNsg 'Microsoft.Network/networkSecurityGroups@2024-05-01' = {
  name: '${prefix}-functions-nsg-${suffix}'
  location: location
  tags: tags
  properties: { securityRules: [] }
}
resource endpointNsg 'Microsoft.Network/networkSecurityGroups@2024-05-01' = {
  name: '${prefix}-endpoints-nsg-${suffix}'
  location: location
  tags: tags
  properties: { securityRules: [] }
}
resource vnet 'Microsoft.Network/virtualNetworks@2024-05-01' = {
  name: '${prefix}-vnet-${suffix}'
  location: location
  tags: tags
  properties: {
    addressSpace: { addressPrefixes: [ addressPrefix ] }
    subnets: [
      {
        name: 'functions'
        properties: {
          addressPrefix: cidrSubnet(addressPrefix, 27, 0)
          networkSecurityGroup: { id: integrationNsg.id }
          delegations: [
            {
              name: 'flex-consumption'
              properties: { serviceName: 'Microsoft.App/environments' }
            }
          ]
        }
      }
      {
        name: 'private-endpoints'
        properties: {
          addressPrefix: cidrSubnet(addressPrefix, 27, 1)
          networkSecurityGroup: { id: endpointNsg.id }
          privateEndpointNetworkPolicies: 'Enabled'
        }
      }
    ]
  }
}
var services = [ 'blob', 'table' ]
resource dnsZones 'Microsoft.Network/privateDnsZones@2024-06-01' = [for service in services: {
  name: 'privatelink.${service}.core.windows.net'
  location: 'global'
  tags: tags
}]
resource dnsLinks 'Microsoft.Network/privateDnsZones/virtualNetworkLinks@2024-06-01' = [for (service, i) in services: {
  parent: dnsZones[i]
  name: '${vnet.name}-link'
  location: 'global'
  properties: {
    registrationEnabled: false
    virtualNetwork: { id: vnet.id }
  }
}]
var endpoints = [
  { name: 'host-blob', accountName: host.name, accountId: host.id, service: 'blob', zoneIndex: 0 }
  { name: 'data-blob', accountName: business.name, accountId: business.id, service: 'blob', zoneIndex: 0 }
  { name: 'data-table', accountName: business.name, accountId: business.id, service: 'table', zoneIndex: 1 }
]
resource privateEndpoints 'Microsoft.Network/privateEndpoints@2024-05-01' = [for endpoint in endpoints: {
  name: '${prefix}-${endpoint.name}-pe-${suffix}'
  location: location
  tags: tags
  properties: {
    subnet: { id: '${vnet.id}/subnets/private-endpoints' }
    privateLinkServiceConnections: [
      {
        name: endpoint.name
        properties: {
          privateLinkServiceId: endpoint.accountId
          groupIds: [ endpoint.service ]
        }
      }
    ]
  }
}]
resource zoneGroups 'Microsoft.Network/privateEndpoints/privateDnsZoneGroups@2024-05-01' = [for (endpoint, i) in endpoints: {
  parent: privateEndpoints[i]
  name: 'storage'
  properties: {
    privateDnsZoneConfigs: [
      {
        name: endpoint.service
        properties: { privateDnsZoneId: dnsZones[endpoint.zoneIndex].id }
      }
    ]
  }
}]

output vnetName string = vnet.name
output integrationSubnetId string = '${vnet.id}/subnets/functions'
output endpointSubnetId string = '${vnet.id}/subnets/private-endpoints'
output privateEndpointConfiguration array = [for (endpoint, i) in endpoints: {
  name: privateEndpoints[i].name
  accountName: endpoint.accountName
  service: endpoint.service
}]
output privateDnsZoneNames array = [for service in services: 'privatelink.${service}.core.windows.net']
