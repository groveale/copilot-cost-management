targetScope = 'resourceGroup'

param prefix string = 'copilotcost'
param location string = 'uksouth'
@description('Exact sub_claim_prefix returned by the repository GitHub OIDC configuration, including immutable IDs when enabled.')
param repositorySubjectPrefix string = 'repo:groveale@17550498/copilot-cost-management@1427172317'

resource infra 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: '${prefix}-github-infra'
  location: location
}
resource app 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: '${prefix}-github-app'
  location: location
}
resource infraWhatIfTrust 'Microsoft.ManagedIdentity/userAssignedIdentities/federatedIdentityCredentials@2023-01-31' = {
  parent: infra
  name: 'github-whatif'
  properties: {
    issuer: 'https://token.actions.githubusercontent.com'
    subject: '${repositorySubjectPrefix}:environment:demo-whatif'
    audiences: [ 'api://AzureADTokenExchange' ]
  }
}
resource infraApplyTrust 'Microsoft.ManagedIdentity/userAssignedIdentities/federatedIdentityCredentials@2023-01-31' = {
  parent: infra
  name: 'github-infra'
  // Managed Identity rejects concurrent federated-credential writes.
  dependsOn: [ infraWhatIfTrust ]
  properties: {
    issuer: 'https://token.actions.githubusercontent.com'
    subject: '${repositorySubjectPrefix}:environment:demo-infra'
    audiences: [ 'api://AzureADTokenExchange' ]
  }
}
resource appTrust 'Microsoft.ManagedIdentity/userAssignedIdentities/federatedIdentityCredentials@2023-01-31' = {
  parent: app
  name: 'github-app'
  properties: {
    issuer: 'https://token.actions.githubusercontent.com'
    subject: '${repositorySubjectPrefix}:environment:demo'
    audiences: [ 'api://AzureADTokenExchange' ]
  }
}
output infraClientId string = infra.properties.clientId
output infraPrincipalId string = infra.properties.principalId
output appClientId string = app.properties.clientId
output appPrincipalId string = app.properties.principalId
