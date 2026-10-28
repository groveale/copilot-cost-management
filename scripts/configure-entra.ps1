param(
  [Parameter(Mandatory)][string] $SpaClientId,
  [string] $ApiDisplayName = 'Copilot Cost Management Requests API',
  [string] $SiteUrl,
  [switch] $EnableSync,
  [switch] $Apply
)
$ErrorActionPreference = 'Stop'
function Invoke-AzJson([string[]] $Arguments) {
  $output = & az @Arguments --output json
  if ($LASTEXITCODE -ne 0) { throw 'Entra operation failed.' }
  if ($output) { return ($output | ConvertFrom-Json -Depth 100) }
}
function Update-App([string] $ObjectId, [object] $Patch) {
  $file = [IO.Path]::GetTempFileName()
  try {
    $Patch | ConvertTo-Json -Depth 100 | Set-Content -LiteralPath $file -Encoding utf8
    Invoke-AzJson @('rest', '--method', 'PATCH', '--url', "https://graph.microsoft.com/v1.0/applications/$ObjectId", '--body', "@$file") | Out-Null
  } finally { Remove-Item -LiteralPath $file }
}
$spa = Invoke-AzJson @('ad', 'app', 'show', '--id', $SpaClientId)
$matches = @(Invoke-AzJson @('ad', 'app', 'list', '--filter', "displayName eq '$($ApiDisplayName.Replace("'", "''"))'"))
if ($matches.Count -gt 1) { throw 'Multiple API registrations match; resolve ambiguity before continuing.' }
if ($SiteUrl) {
  $uri = [uri] $SiteUrl
  if ($uri.Scheme -ne 'https' -or $uri.GetLeftPart([UriPartial]::Authority) -ne $SiteUrl) {
    throw 'SiteUrl must be the actual HTTPS origin without a trailing slash.'
  }
}
if (!$Apply) {
  Write-Output 'Preview only: create/reuse a single-tenant API registration; expose v2 access_as_user; preauthorize this SPA; preserve existing Graph permissions; ensure enterprise application.'
  if ($EnableSync) { Write-Output 'Also add delegated sync read scopes and Sites.Selected to the SPA. Consent and the explicit site write grant are separate operator steps.' }
  if ($SiteUrl) { Write-Output "Also preserve existing SPA redirects and add $SiteUrl/auth.html and the localhost callback." }
  Write-Output "Existing matching API registrations: $($matches.Count). Rerun with -Apply after review."
  return
}
$api = if ($matches.Count -eq 1) { $matches[0] } else {
  Invoke-AzJson @('ad', 'app', 'create', '--display-name', $ApiDisplayName, '--sign-in-audience', 'AzureADMyOrg')
}
if ($api.signInAudience -ne 'AzureADMyOrg') { throw 'Existing API registration is not single-tenant.' }
$scopes = @($api.api.oauth2PermissionScopes | Where-Object { $null -ne $_ })
$scope = @($scopes | Where-Object { $_.value -eq 'access_as_user' })
if ($scope.Count -gt 1) { throw 'API has duplicate access_as_user scopes.' }
if ($scope.Count -eq 0) {
  $scopeId = [guid]::NewGuid().ToString()
  $scopes += @{
    id = $scopeId; value = 'access_as_user'; type = 'User'; isEnabled = $true
    adminConsentDisplayName = 'Use the Copilot budget requests API'
    adminConsentDescription = 'Submit your own requests and manage decisions for budgets you own.'
    userConsentDisplayName = 'Use the Copilot budget requests API'
    userConsentDescription = 'Submit your own requests and manage decisions for budgets you own.'
  }
} else {
  if (!$scope[0].isEnabled) { throw 'Existing API scope is disabled; review it before continuing.' }
  $scopeId = $scope[0].id
}
$preauthorized = @($api.api.preAuthorizedApplications | Where-Object { $null -ne $_ -and $_.appId -ne $SpaClientId })
$currentSpa = @($api.api.preAuthorizedApplications | Where-Object { $_.appId -eq $SpaClientId })
$permissionIds = @($currentSpa.delegatedPermissionIds) + @($scopeId) | Where-Object { $_ } | Select-Object -Unique
$preauthorized += @{ appId = $SpaClientId; delegatedPermissionIds = @($permissionIds) }
$apiPatch = @{
  requestedAccessTokenVersion = 2
  oauth2PermissionScopes = @($scopes)
}
Update-App $api.id @{
  identifierUris = @(@($api.identifierUris) + @("api://$($api.appId)") | Where-Object { $_ } | Select-Object -Unique)
  api = $apiPatch
}
# Entra validates preauthorization against already-persisted permissions.
Update-App $api.id @{ api = @{ preAuthorizedApplications = @($preauthorized) } }
$permissions = @($spa.requiredResourceAccess | Where-Object { $null -ne $_ -and $_.resourceAppId -ne $api.appId })
$existingApiPermission = @($spa.requiredResourceAccess | Where-Object { $_.resourceAppId -eq $api.appId })
$access = @($existingApiPermission.resourceAccess | Where-Object { $null -ne $_ -and $_.id -ne $scopeId })
$access += @{ id = $scopeId; type = 'Scope' }
$permissions += @{ resourceAppId = $api.appId; resourceAccess = @($access) }
if ($EnableSync) {
  $graphId = '00000003-0000-0000-c000-000000000000'
  $graph = @(Invoke-AzJson @('ad', 'sp', 'list', '--filter', "appId eq '$graphId'"))
  if ($graph.Count -ne 1) { throw 'The Microsoft Graph service principal could not be uniquely resolved.' }
  $wanted = @('CopilotCostManagement-Policy.Read.All', 'CopilotCostManagement-Assignment.Read.All',
    'CopilotCostManagement-UserData.Read.All', 'GroupMember.Read.All', 'User.ReadBasic.All', 'Sites.Selected')
  $currentGraph = @($permissions | Where-Object resourceAppId -EQ $graphId)
  $graphAccess = @($currentGraph.resourceAccess | Where-Object { $null -ne $_ })
  foreach ($value in $wanted) {
    $advertised = @($graph[0].oauth2PermissionScopes | Where-Object { $_.value -eq $value -and $_.isEnabled })
    if ($advertised.Count -ne 1) { throw "Delegated sync permission $value is not advertised in this tenant. Do not substitute application permissions." }
    if (-not @($graphAccess | Where-Object { $_.id -eq $advertised[0].id -and $_.type -eq 'Scope' }).Count) {
      $graphAccess += @{ id = $advertised[0].id; type = 'Scope' }
    }
  }
  $permissions = @($permissions | Where-Object resourceAppId -NE $graphId) + @(@{ resourceAppId = $graphId; resourceAccess = $graphAccess })
}
$spaPatch = @{ requiredResourceAccess = @($permissions) }
if ($SiteUrl) {
  $spaPatch.spa = @{
    redirectUris = @(@($spa.spa.redirectUris) + @("$SiteUrl/auth.html", 'http://localhost:5173/auth.html') | Select-Object -Unique)
  }
}
Update-App $spa.id $spaPatch
$principals = @(Invoke-AzJson @('ad', 'sp', 'list', '--filter', "appId eq '$($api.appId)'"))
if ($principals.Count -eq 0) { Invoke-AzJson @('ad', 'sp', 'create', '--id', $api.appId) | Out-Null }
Write-Output "API_CLIENT_ID=$($api.appId)"
Write-Output 'API configured; existing SPA Graph permissions were retained. API scope is preauthorized for this SPA only.'
