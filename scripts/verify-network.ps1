param(
  [Parameter(Mandatory)][string] $ResourceGroup,
  [Parameter(Mandatory)][string] $DeploymentName
)
$ErrorActionPreference = 'Stop'
function Invoke-AzJson([string[]] $Arguments) {
  $output = & az @Arguments --output json
  if ($LASTEXITCODE -ne 0) { throw "Azure verification query failed: $($Arguments[0..1] -join ' ')." }
  return ($output | ConvertFrom-Json -Depth 100)
}
function Require([bool] $Condition, [string] $Message) {
  if (!$Condition) { throw $Message }
}
$deployment = Invoke-AzJson @('deployment', 'group', 'show', '--resource-group', $ResourceGroup, '--name', $DeploymentName)
Require ($deployment.properties.provisioningState -eq 'Succeeded') 'The infrastructure deployment did not succeed.'
$outputs = $deployment.properties.outputs
Require (!!$outputs.integrationSubnetId.value) 'Deployment has no private-network outputs.'
$function = Invoke-AzJson @('resource', 'show', '--ids', $outputs.functionId.value, '--api-version', '2024-04-01')
Require ($function.identity.type -eq 'SystemAssigned') 'Function runtime must retain its system-assigned identity.'
Require ($function.properties.virtualNetworkSubnetId -eq $outputs.integrationSubnetId.value) 'Function is not integrated with the expected subnet.'
Require ($function.properties.functionAppConfig.deployment.storage.authentication.type -eq 'SystemAssignedIdentity') 'Function package storage is not using the system-assigned identity.'
$subnet = Invoke-AzJson @('network', 'vnet', 'subnet', 'show', '--ids', $outputs.integrationSubnetId.value)
Require (@($subnet.delegations.serviceName) -contains 'Microsoft.App/environments') 'Incorrect Flex subnet delegation.'
foreach ($name in @($outputs.hostStorageName.value, $outputs.businessStorageName.value)) {
  $account = Invoke-AzJson @('storage', 'account', 'show', '--resource-group', $ResourceGroup, '--name', $name)
  Require ($account.publicNetworkAccess -eq 'Disabled') "Public networking is not disabled on $name."
  Require ($account.allowSharedKeyAccess -eq $false -and $account.allowBlobPublicAccess -eq $false) "Key/anonymous access is enabled on $name."
  Require ($account.networkRuleSet.defaultAction -eq 'Deny' -and $account.networkRuleSet.bypass -eq 'None') "Unexpected public firewall bypass on $name."
}
foreach ($endpoint in $outputs.privateEndpointConfiguration.value) {
  $pe = Invoke-AzJson @('network', 'private-endpoint', 'show', '--resource-group', $ResourceGroup, '--name', $endpoint.name)
  Require ($pe.provisioningState -eq 'Succeeded') "Private endpoint $($endpoint.name) did not provision."
  Require ($pe.subnet.id -eq $outputs.endpointSubnetId.value) "Private endpoint $($endpoint.name) is in the wrong subnet."
  Require (@($pe.privateLinkServiceConnections).Count -eq 1) "Unexpected connections on $($endpoint.name)."
  Require ($pe.privateLinkServiceConnections[0].privateLinkServiceConnectionState.status -eq 'Approved') "Private endpoint $($endpoint.name) is not approved."
  Require (@($pe.privateLinkServiceConnections[0].groupIds) -contains $endpoint.service) "Incorrect storage subresource on $($endpoint.name)."
  $nic = Invoke-AzJson @('network', 'nic', 'show', '--ids', $pe.networkInterfaces[0].id)
  $ips = @($nic.ipConfigurations.privateIPAddress)
  Require ($ips.Count -gt 0) "Private endpoint $($endpoint.name) has no private IP."
  $dns = Invoke-AzJson @('network', 'private-dns', 'record-set', 'a', 'show', '--resource-group', $ResourceGroup,
    '--zone-name', "privatelink.$($endpoint.service).core.windows.net", '--name', $endpoint.accountName)
  $dnsIps = @($dns.aRecords.ipv4Address)
  foreach ($ip in $ips) { Require ($dnsIps -contains $ip) "Private DNS does not match the NIC for $($endpoint.name)." }
  Write-Output "Approved private endpoint and matching DNS: $($endpoint.name) ($($ips -join ', '))."
}
foreach ($zone in $outputs.privateDnsZoneNames.value) {
  $links = @(Invoke-AzJson @('network', 'private-dns', 'link', 'vnet', 'list', '--resource-group', $ResourceGroup, '--zone-name', $zone))
  $matches = @($links | Where-Object { $_.virtualNetwork.id -eq ($outputs.integrationSubnetId.value -replace '/subnets/[^/]+$', '') })
  Require ($matches.Count -eq 1 -and $matches[0].virtualNetworkLinkState -eq 'Completed') "Private DNS zone $zone is not linked to the VNet."
}
Write-Output 'Private network control-plane checks passed. Runtime DNS/data-plane access and private seeding require separate in-network verification.'
