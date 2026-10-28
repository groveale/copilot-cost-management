using './main.bicep'

param prefix = 'copilotcost'
param location = 'uksouth'
param staticLocation = 'westeurope'
param tenantId = readEnvironmentVariable('AZURE_TENANT_ID')
param spaClientId = readEnvironmentVariable('SPA_CLIENT_ID')
param apiClientId = readEnvironmentVariable('API_CLIENT_ID')
param bootstrapEnabled = bool(readEnvironmentVariable('BOOTSTRAP_ENABLED', 'false'))
param syncDefaultConcurrency = int(readEnvironmentVariable('SYNC_DEFAULT_CONCURRENCY', '3'))
param syncMaxConcurrency = int(readEnvironmentVariable('SYNC_MAX_CONCURRENCY', '8'))
param syncJobRetentionDays = int(readEnvironmentVariable('SYNC_JOB_RETENTION_DAYS', '30'))
