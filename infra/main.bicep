targetScope = 'resourceGroup'

param prefix string = 'copilotcost'
param location string = 'uksouth'
param staticLocation string = 'westeurope'
param tenantId string
param spaClientId string
@description('API registration client ID. Configure v2 access tokens in Entra.')
param apiClientId string
@secure()
@description('Comma-separated onboarding administrator object IDs. Empty disables administration.')
param onboardingAdminObjectIds string = ''
@description('Explicitly enable additive private bootstrap import; disable after initial seeding.')
param bootstrapEnabled bool = false
@secure()
@description('SharePoint sync siteId, policyListId, balanceListId and jobListId. Empty disables sync.')
param syncDestination object = {}
@minValue(1)
@maxValue(8)
param syncDefaultConcurrency int = 3
@minValue(1)
@maxValue(8)
param syncMaxConcurrency int = 8
@minValue(1)
@maxValue(3650)
param syncJobRetentionDays int = 30
@secure()
@description('Explicit reporting opt-in: enabled, siteUrl, adminGroupId and versioned certificateId. Contains no private key or secret.')
param reportingSecurity object = {}

var suffix = uniqueString(resourceGroup().id, prefix)
var tags = { application: 'copilot-cost-demo', environment: 'demo' }

resource host 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: 'sthost${take(suffix, 14)}'
  location: location
  tags: tags
  kind: 'StorageV2'
  sku: { name: 'Standard_LRS' }
  properties: {
    allowSharedKeyAccess: false
    allowBlobPublicAccess: false
    defaultToOAuthAuthentication: true
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
    publicNetworkAccess: 'Disabled'
    networkAcls: { bypass: 'None', defaultAction: 'Deny', ipRules: [], virtualNetworkRules: [] }
  }
}
resource hostBlobs 'Microsoft.Storage/storageAccounts/blobServices@2023-05-01' = {
  parent: host
  name: 'default'
}
resource packages 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = {
  parent: hostBlobs
  name: 'packages'
  properties: { publicAccess: 'None' }
}
resource business 'Microsoft.Storage/storageAccounts@2023-05-01' = {
  name: 'stdata${take(suffix, 14)}'
  location: location
  tags: tags
  kind: 'StorageV2'
  sku: { name: 'Standard_LRS' }
  properties: {
    allowSharedKeyAccess: false
    allowBlobPublicAccess: false
    defaultToOAuthAuthentication: true
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
    publicNetworkAccess: 'Disabled'
    networkAcls: { bypass: 'None', defaultAction: 'Deny', ipRules: [], virtualNetworkRules: [] }
  }
}
resource businessBlobs 'Microsoft.Storage/storageAccounts/blobServices@2023-05-01' = {
  parent: business
  name: 'default'
}
resource metadata 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = {
  parent: businessBlobs
  name: 'metadata'
  properties: { publicAccess: 'None' }
}
resource syncCheckpoints 'Microsoft.Storage/storageAccounts/blobServices/containers@2023-05-01' = {
  parent: businessBlobs
  name: 'sync-checkpoints'
  properties: { publicAccess: 'None' }
}
resource tables 'Microsoft.Storage/storageAccounts/tableServices@2023-05-01' = {
  parent: business
  name: 'default'
}
resource requests 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' = {
  parent: tables
  name: 'Requests'
}
resource index 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' = {
  parent: tables
  name: 'RequesterBudgetIndex'
}
resource onboarding 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' = {
  parent: tables
  name: 'OnboardingRequests'
}
resource syncJobs 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-05-01' = {
  parent: tables
  name: 'SyncJobs'
}
module network './network.bicep' = {
  name: 'private-storage-network'
  params: {
    prefix: prefix
    suffix: suffix
    location: location
    tags: tags
    hostStorageName: host.name
    businessStorageName: business.name
  }
}
var reportingEnabled = reportingSecurity.?enabled ?? false
var signingVaultName = '${take(prefix, 8)}-rpt-${take(suffix, 10)}'
module reporting './reporting.bicep' = if (reportingEnabled) {
  name: 'reporting-certificate-vault'
  params: {
    vaultName: signingVaultName
    location: location
    tags: tags
    tenantId: tenantId
    endpointSubnetId: network.outputs.endpointSubnetId
    vnetName: network.outputs.vnetName
  }
}
resource workspace 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {
  name: '${prefix}-logs-${suffix}'
  location: location
  tags: tags
  properties: {
    retentionInDays: 30
    sku: { name: 'PerGB2018' }
    workspaceCapping: { dailyQuotaGb: json('0.1') }
  }
}
resource insights 'Microsoft.Insights/components@2020-02-02' = {
  name: '${prefix}-insights-${suffix}'
  location: location
  kind: 'web'
  tags: tags
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: workspace.id
    DisableLocalAuth: true
    RetentionInDays: 30
  }
}
resource site 'Microsoft.Web/staticSites@2023-12-01' = {
  name: '${prefix}-web-${suffix}'
  location: staticLocation
  tags: tags
  sku: { name: 'Free', tier: 'Free' }
  properties: {
    allowConfigFileUpdates: true
    stagingEnvironmentPolicy: 'Disabled'
    enterpriseGradeCdnStatus: 'Disabled'
  }
}
resource plan 'Microsoft.Web/serverfarms@2024-04-01' = {
  name: '${prefix}-plan-${suffix}'
  location: location
  tags: tags
  kind: 'functionapp'
  sku: { name: 'FC1', tier: 'FlexConsumption' }
  properties: { reserved: true }
}
resource api 'Microsoft.Web/sites@2024-04-01' = {
  name: '${prefix}-api-${suffix}'
  location: location
  kind: 'functionapp,linux'
  tags: tags
  identity: { type: 'SystemAssigned' }
  properties: {
    serverFarmId: plan.id
    httpsOnly: true
    virtualNetworkSubnetId: network.outputs.integrationSubnetId
    siteConfig: {
      minTlsVersion: '1.2'
      ftpsState: 'Disabled'
      cors: {
        allowedOrigins: [ 'https://${site.properties.defaultHostname}' ]
        supportCredentials: false
      }
      appSettings: [
        { name: 'AzureWebJobsStorage__accountName', value: host.name }
        { name: 'AzureWebJobsStorage__credential', value: 'managedidentity' }
        { name: 'APPLICATIONINSIGHTS_CONNECTION_STRING', value: insights.properties.ConnectionString }
        { name: 'APPLICATIONINSIGHTS_AUTHENTICATION_STRING', value: 'Authorization=AAD' }
        { name: 'TENANT_ID', value: tenantId }
        { name: 'SPA_CLIENT_ID', value: spaClientId }
        { name: 'API_AUDIENCE', value: apiClientId }
        { name: 'BUSINESS_TABLE_URL', value: business.properties.primaryEndpoints.table }
        { name: 'BUSINESS_BLOB_URL', value: business.properties.primaryEndpoints.blob }
        { name: 'ONBOARDING_ADMIN_OBJECT_IDS', value: onboardingAdminObjectIds }
        { name: 'BOOTSTRAP_ENABLED', value: bootstrapEnabled ? 'true' : 'false' }
        { name: 'SYNC_SITE_ID', value: syncDestination.?siteId ?? '' }
        { name: 'SYNC_POLICY_LIST_ID', value: syncDestination.?policyListId ?? '' }
        { name: 'SYNC_BALANCE_LIST_ID', value: syncDestination.?balanceListId ?? '' }
        { name: 'SYNC_JOB_LIST_ID', value: syncDestination.?jobListId ?? '' }
        { name: 'SYNC_DEFAULT_CONCURRENCY', value: string(syncDefaultConcurrency) }
        { name: 'SYNC_MAX_CONCURRENCY', value: string(syncMaxConcurrency) }
        { name: 'SYNC_JOB_RETENTION_DAYS', value: string(syncJobRetentionDays) }
        { name: 'REPORTING_SECURITY_ENABLED', value: reportingEnabled ? 'true' : 'false' }
        { name: 'REPORTING_SITE_URL', value: reportingSecurity.?siteUrl ?? '' }
        { name: 'REPORTING_ADMIN_GROUP_ID', value: string(reportingSecurity.?adminGroupId ?? '') }
        { name: 'REPORTING_CERTIFICATE_ID', value: reportingSecurity.?certificateId ?? '' }
        { name: 'REPORTING_MAX_SCOPES', value: '5000' }
        { name: 'REPORTING_SCOPE_RESERVE', value: '250' }
      ]
    }
    functionAppConfig: {
      runtime: { name: 'node', version: '24' }
      deployment: {
        storage: {
          type: 'blobContainer'
          value: '${host.properties.primaryEndpoints.blob}${packages.name}'
          authentication: { type: 'SystemAssignedIdentity' }
        }
      }
      scaleAndConcurrency: {
        maximumInstanceCount: 40
        instanceMemoryMB: 512
        triggers: { http: { perInstanceConcurrency: 8 } }
      }
    }
  }
}
resource scm 'Microsoft.Web/sites/basicPublishingCredentialsPolicies@2024-04-01' = {
  parent: api
  name: 'scm'
  properties: { allow: false }
}
resource ftp 'Microsoft.Web/sites/basicPublishingCredentialsPolicies@2024-04-01' = {
  parent: api
  name: 'ftp'
  properties: { allow: false }
}

output functionName string = api.name
output reportingVaultName string = reportingEnabled ? signingVaultName : ''
output functionPrincipalId string = api.identity.principalId
output functionId string = api.id
output apiUrl string = 'https://${api.properties.defaultHostName}'
output siteName string = site.name
output siteUrl string = 'https://${site.properties.defaultHostname}'
output hostStorageName string = host.name
output businessStorageName string = business.name
output tableUrl string = business.properties.primaryEndpoints.table
output blobUrl string = business.properties.primaryEndpoints.blob
output insightsName string = insights.name
output vnetName string = network.outputs.vnetName
output integrationSubnetId string = network.outputs.integrationSubnetId
output endpointSubnetId string = network.outputs.endpointSubnetId
output privateEndpointConfiguration array = network.outputs.privateEndpointConfiguration
output privateDnsZoneNames array = network.outputs.privateDnsZoneNames
