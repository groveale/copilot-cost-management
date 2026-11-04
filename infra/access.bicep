targetScope = 'resourceGroup'

param functionName string
param functionPrincipalId string
param hostStorageName string
param businessStorageName string
param insightsName string
param infraPrincipalId string
param appPrincipalId string
@description('Optional authorized seed operator object ID; no key credentials.')
param seedOperatorId string = ''
@description('Dedicated reporting signing vault from main outputs; empty disables reporting signing access.')
param reportingVaultName string = ''

resource api 'Microsoft.Web/sites@2024-04-01' existing = { name: functionName }
resource host 'Microsoft.Storage/storageAccounts@2023-05-01' existing = { name: hostStorageName }
resource business 'Microsoft.Storage/storageAccounts@2023-05-01' existing = { name: businessStorageName }
resource tableService 'Microsoft.Storage/storageAccounts/tableServices@2023-05-01' existing = {
  parent: business
  name: 'default'
}
resource requests 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' existing = {
  parent: tableService
  name: 'Requests'
}
resource index 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' existing = {
  parent: tableService
  name: 'RequesterBudgetIndex'
}
resource onboarding 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' existing = {
  parent: tableService
  name: 'OnboardingRequests'
}
resource syncJobs 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' existing = {
  parent: tableService
  name: 'SyncJobs'
}
resource blobs 'Microsoft.Storage/storageAccounts/blobServices@2023-05-01' existing = {
  parent: business
  name: 'default'
}
resource metadata 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' existing = {
  parent: blobs
  name: 'metadata'
}
resource syncCheckpoints 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' existing = {
  parent: blobs
  name: 'sync-checkpoints'
}
resource insights 'Microsoft.Insights/components@2020-02-02' existing = { name: insightsName }
resource reportingVault 'Microsoft.KeyVault/vaults@2023-07-01' existing = if (!empty(reportingVaultName)) {
  name: reportingVaultName
}
resource reportingSignerRole 'Microsoft.Authorization/roleDefinitions@2022-04-01' = if (!empty(reportingVaultName)) {
  name: guid(resourceGroup().id, 'copilot-reporting-certificate-sign')
  properties: {
    roleName: '${resourceGroup().name} reporting certificate signer'
    description: 'Read public certificates and sign digests only in the dedicated reporting vault. No secret/private-key retrieval or key management.'
    type: 'CustomRole'
    assignableScopes: [ resourceGroup().id ]
    permissions: [
      {
        actions: []
        notActions: []
        dataActions: [ 'Microsoft.KeyVault/vaults/certificates/read', 'Microsoft.KeyVault/vaults/keys/sign/action' ]
        notDataActions: []
      }
    ]
  }
}
resource reportingSigner 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!empty(reportingVaultName)) {
  scope: reportingVault
  name: guid(reportingVault.id, functionPrincipalId, 'reporting-certificate-signer')
  properties: {
    principalId: functionPrincipalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: reportingSignerRole.id
  }
}

var blobOwner = 'b7e6dc6d-f1e8-4753-8033-0f276bb0955b'
var blobContributor = 'ba92f5b4-2d11-453d-a403-e96b0029c9fe'
var tableContributor = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'
var metricsPublisher = '3913510d-42f4-4e42-8a64-420c390055eb'
var contributor = 'b24988ac-6180-42a0-ab88-20f7382dd24c'
var websiteContributor = 'de139f84-1756-47ae-9be6-808fbbe84772'
var runtimeId = functionPrincipalId

// HTTP-only Functions host needs blob ownership; there are no queue, durable or table triggers.
resource hostAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: host
  name: guid(host.id, runtimeId, blobOwner)
  properties: {
    principalId: runtimeId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', blobOwner)
  }
}
resource requestsAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: requests
  name: guid(requests.id, runtimeId, tableContributor)
  properties: {
    principalId: runtimeId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', tableContributor)
  }
}
resource indexAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: index
  name: guid(index.id, runtimeId, tableContributor)
  properties: {
    principalId: runtimeId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', tableContributor)
  }
}
resource metadataAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: metadata
  name: guid(metadata.id, runtimeId, blobContributor)
  properties: {
    principalId: runtimeId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', blobContributor)
  }
}
resource onboardingAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: onboarding
  name: guid(onboarding.id, runtimeId, tableContributor)
  properties: {
    principalId: runtimeId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', tableContributor)
  }
}
resource syncJobsAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: syncJobs
  name: guid(syncJobs.id, runtimeId, tableContributor)
  properties: {
    principalId: runtimeId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', tableContributor)
  }
}
resource syncCheckpointAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: syncCheckpoints
  name: guid(syncCheckpoints.id, runtimeId, blobContributor)
  properties: {
    principalId: runtimeId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', blobContributor)
  }
}
resource telemetryAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: insights
  name: guid(insights.id, runtimeId, metricsPublisher)
  properties: {
    principalId: runtimeId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', metricsPublisher)
  }
}
resource infraAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(resourceGroup().id, infraPrincipalId, contributor)
  properties: {
    principalId: infraPrincipalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', contributor)
  }
}
resource appAccess 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: api
  name: guid(api.id, appPrincipalId, websiteContributor)
  properties: {
    principalId: appPrincipalId
    principalType: 'ServicePrincipal'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', websiteContributor)
  }
}
resource seedIndex 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!empty(seedOperatorId)) {
  scope: index
  name: guid(index.id, seedOperatorId, tableContributor)
  properties: {
    principalId: seedOperatorId
    principalType: 'User'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', tableContributor)
  }
}
resource seedMetadata 'Microsoft.Authorization/roleAssignments@2022-04-01' = if (!empty(seedOperatorId)) {
  scope: metadata
  name: guid(metadata.id, seedOperatorId, blobContributor)
  properties: {
    principalId: seedOperatorId
    principalType: 'User'
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', blobContributor)
  }
}
