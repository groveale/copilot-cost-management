using './main.bicep'

param prefix = 'copilotcost'
param location = 'uksouth'
param staticLocation = 'westeurope'
param tenantId = readEnvironmentVariable('AZURE_TENANT_ID')
param spaClientId = readEnvironmentVariable('SPA_CLIENT_ID')
param apiClientId = readEnvironmentVariable('API_CLIENT_ID')
