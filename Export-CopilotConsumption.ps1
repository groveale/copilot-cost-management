#Requires -Version 7.4
<#
.SYNOPSIS
Retains Copilot consumption snapshots as policy JSON and a combined reporting CSV.
.DESCRIPTION
Run Setup interactively once, then schedule Export under the same Windows account.
Graph PowerShell manages the delegated token cache. Export uses a supervised
device-code worker: it never launches a browser and aborts if the SDK requests
interaction. See docs\consumption-export.md for permissions and reporting limits.
#>
[CmdletBinding()]
param(
    [ValidateSet('Setup', 'Export', 'Worker')]
    [string]$Mode = 'Export',
    [guid]$TenantId,
    [guid]$ClientId,
    [string]$Account,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'CopilotCostManagement\ConsumptionExports'),
    [ValidateRange(10, 600)][int]$AuthenticationTimeoutSeconds = 90,
    [ValidateRange(5, 120)][int]$RequestTimeoutSeconds = 30,
    [ValidateRange(0, 10)][int]$MaxRetryCount = 3
)

$script:ConsumptionScriptPath = $PSCommandPath
$script:ConsumptionRepositoryPath = $PSScriptRoot
$script:ConsumptionProtocol = $false
$script:ConsumptionScopes = @(
    'CopilotCostManagement-Policy.Read.All',
    'CopilotCostManagement-Assignment.Read.All',
    'CopilotCostManagement-UserData.Read.All',
    'GroupMember.Read.All',
    'User.ReadBasic.All'
)
$script:ConsumptionCsvColumns = @(
    'schemaVersion', 'runId', 'tenantId', 'collectionStartedAtUtc', 'observedAtUtc',
    'discoveryPolicyId', 'discoveryPolicyName', 'sourceAssignedGroupId', 'userId',
    'userDisplayName', 'userPrincipalName', 'serviceId', 'returnedSpendingPolicyId',
    'policyMatch', 'validationStatus', 'consumedQuantity', 'unit', 'remainingQuantityKind',
    'remainingQuantityValue', 'isConsumptionAllowed', 'accessDecisionReason', 'sourceLastModifiedDateTime'
)

function Get-ConsumptionTime { [DateTimeOffset]::UtcNow.ToString('o') }

function Stop-ConsumptionOperation {
    param([string]$Category, [string]$Message)
    $exception = [InvalidOperationException]::new($Message)
    $exception.Data['ConsumptionCategory'] = $Category
    throw $exception
}

function Get-ConsumptionFailure {
    param([System.Management.Automation.ErrorRecord]$Record)
    $exception = $Record.Exception
    while ($exception.InnerException -and -not $exception.Data.Contains('ConsumptionCategory')) {
        $exception = $exception.InnerException
    }
    $category = 'OperationFailed'
    if ($exception.Data.Contains('ConsumptionCategory')) { $category = $exception.Data['ConsumptionCategory'] }
    $message = $exception.Message -replace '(?i)Bearer\s+\S+', 'Bearer [redacted]'
    $message = $message -replace 'eyJ[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+\.[a-zA-Z0-9_-]+', '[redacted token]'
    [ordered]@{ category = $category; message = $message.Substring(0, [Math]::Min(600, $message.Length)) }
}

function Write-ConsumptionEvent {
    param([hashtable]$Event)
    if ($script:ConsumptionProtocol) {
        [Console]::Out.WriteLine('COPILOT_EXPORT:' + (ConvertTo-Json -InputObject $Event -Compress -Depth 10))
    }
}

function Write-ConsumptionJson {
    param([string]$Path, [object]$Value)
    $temporary = "$Path.$([guid]::NewGuid().ToString('N')).tmp"
    try {
        $json = ConvertTo-Json -InputObject $Value -Depth 100
        [IO.File]::WriteAllText($temporary, $json, [Text.UTF8Encoding]::new($false))
        [IO.File]::Move($temporary, $Path, $true)
    } finally {
        if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
    }
}

function Assert-ConsumptionConfiguration {
    param([hashtable]$Config)
    if ($Config.tenantId -eq [guid]::Empty.ToString() -or $Config.clientId -eq [guid]::Empty.ToString() -or
        [string]::IsNullOrWhiteSpace($Config.account)) {
        Stop-ConsumptionOperation 'InvalidConfiguration' 'TenantId, ClientId and the expected Account UPN are required.'
    }
    $destination = [IO.Path]::GetFullPath($Config.outputPath)
    $repository = [IO.Path]::GetFullPath($script:ConsumptionRepositoryPath).TrimEnd('\', '/')
    if ($destination.Equals($repository, [StringComparison]::OrdinalIgnoreCase) -or
        $destination.StartsWith($repository + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        Stop-ConsumptionOperation 'UnsafeOutputPath' 'Store user-level exports outside the repository and its public assets.'
    }
    $ancestor = $destination
    while ($ancestor) {
        if ((Test-Path -LiteralPath $ancestor) -and
            ((Get-Item -LiteralPath $ancestor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            Stop-ConsumptionOperation 'UnsafeOutputPath' 'The export destination must not traverse a symbolic link or junction.'
        }
        $ancestor = [IO.Path]::GetDirectoryName($ancestor)
    }
}

function Connect-ConsumptionGraph {
    param([hashtable]$Config)
    try {
        Import-Module Microsoft.Graph.Authentication -MinimumVersion 2.37.0 -ErrorAction Stop -WarningAction Stop
    } catch {
        Stop-ConsumptionOperation 'DependencyMissing' 'Install Microsoft.Graph.Authentication 2.37.0 or later for this Windows account.'
    }
    Write-ConsumptionEvent @{ type = 'phase'; phase = 'authentication' }
    try {
        Connect-MgGraph -TenantId $Config.tenantId -ClientId $Config.clientId `
            -Scopes $script:ConsumptionScopes -ContextScope CurrentUser -UseDeviceCode `
            -ClientTimeout $Config.requestTimeoutSeconds -NoWelcome -ErrorAction Stop | Out-Null
    } catch {
        Stop-ConsumptionOperation 'AuthenticationFailed' 'Graph sign-in failed. Check consent and Entra sign-in logs; run Setup interactively if necessary.'
    }
    $context = Get-MgContext
    if (-not $context -or $context.AuthType -ne 'Delegated' -or
        $context.ContextScope -ne 'CurrentUser' -or
        $context.TenantId -ne $Config.tenantId -or $context.ClientId -ne $Config.clientId -or
        $context.Account -ne $Config.account) {
        Stop-ConsumptionOperation 'AccountMismatch' 'Expected a persistent delegated connection for the configured tenant, client and account. Use a dedicated Windows profile or explicitly correct the Graph sign-in.'
    }
    foreach ($scope in $script:ConsumptionScopes) {
        if ($scope -notin @($context.Scopes | ForEach-Object { $_ -replace '^https://graph.microsoft.com/', '' })) {
            Stop-ConsumptionOperation 'MissingPermission' "The delegated connection is missing $scope. Review consent and run Setup again."
        }
    }
    Write-ConsumptionEvent @{ type = 'phase'; phase = 'processing' }
}

function Resolve-ConsumptionGraphUri {
    param([string]$Uri, [string]$ExpectedPath)
    $parsed = $null
    if (-not [uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsed)) {
        if (-not $Uri.StartsWith('/') -or $Uri.StartsWith('//')) {
            Stop-ConsumptionOperation 'InvalidContinuation' 'Graph returned an invalid route.'
        }
        $parsed = [uri]("https://graph.microsoft.com$Uri")
    }
    if ($parsed.Scheme -ne 'https' -or $parsed.Host -ne 'graph.microsoft.com' -or
        -not $parsed.IsDefaultPort -or $parsed.UserInfo -or $parsed.Fragment -or
        $parsed.AbsolutePath -cne $ExpectedPath) {
        Stop-ConsumptionOperation 'InvalidContinuation' 'Refusing a Graph route outside the requested HTTPS collection.'
    }
    $parsed.AbsoluteUri
}

function Invoke-ConsumptionGraphGet {
    param([string]$Uri, [hashtable]$Config)
    for ($attempt = 0; ; $attempt++) {
        Write-ConsumptionEvent @{ type = 'phase'; phase = 'request' }
        try {
            $result = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -ErrorAction Stop
            Write-ConsumptionEvent @{ type = 'phase'; phase = 'processing' }
            return $result
        } catch {
            $response = $_.Exception.Response
            $status = if ($response) { [int]$response.StatusCode } else { 0 }
            if ($status -in @(429, 503, 504) -and $attempt -lt $Config.maxRetryCount) {
                $delay = [Math]::Min(60, [Math]::Pow(2, $attempt + 1))
                if ($response.Headers.RetryAfter.Delta) {
                    $delay = [Math]::Max(1, [Math]::Ceiling($response.Headers.RetryAfter.Delta.TotalSeconds))
                } elseif ($response.Headers.RetryAfter.Date) {
                    $delay = [Math]::Max(1, [Math]::Ceiling(($response.Headers.RetryAfter.Date - [DateTimeOffset]::UtcNow).TotalSeconds))
                }
                if ($delay -gt 60) {
                    Stop-ConsumptionOperation 'Throttled' "Graph requested a retry after $delay seconds; rerun later."
                }
                Write-ConsumptionEvent @{ type = 'phase'; phase = 'retry'; delaySeconds = $delay }
                Start-Sleep -Seconds $delay
                continue
            }
            $requestId = if ($response -and $response.Headers.Contains('request-id')) {
                @($response.Headers.GetValues('request-id'))[0] -replace '[^a-zA-Z0-9-]', ''
            } else { 'unavailable' }
            $category = if ($status -eq 401) { 'AuthenticationFailed' } elseif ($status -eq 403) { 'AuthorizationFailed' } else { 'GraphRequestFailed' }
            Stop-ConsumptionOperation $category "Graph GET failed (HTTP $status; request-id $requestId). Check permissions, preview access and service availability."
        }
    }
}

function Get-ConsumptionCollection {
    param([string]$Path, [hashtable]$Config)
    $items = [Collections.Generic.List[object]]::new()
    $visited = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $next = $Path
    try {
        while ($next) {
            $uri = Resolve-ConsumptionGraphUri $next $Path
            if (-not $visited.Add($uri)) { Stop-ConsumptionOperation 'InvalidContinuation' 'Graph repeated a continuation link.' }
            $page = Invoke-ConsumptionGraphGet $uri $Config
            if ($page -isnot [Collections.IDictionary] -or -not $page.Contains('value') -or
                $page.value -isnot [System.Collections.IList]) {
                Stop-ConsumptionOperation 'InvalidResponse' 'Graph collection did not contain a value array.'
            }
            foreach ($item in $page.value) { $items.Add($item) }
            $next = $page['@odata.nextLink']
            if ($null -ne $next -and ($next -isnot [string] -or [string]::IsNullOrWhiteSpace($next))) {
                Stop-ConsumptionOperation 'InvalidContinuation' 'Graph returned an invalid continuation link.'
            }
        }
        @{ items = $items.ToArray(); complete = $true; error = $null }
    } catch {
        @{ items = $items.ToArray(); complete = $false; error = (Get-ConsumptionFailure $_) }
    }
}

function Test-ConsumptionId {
    param([object]$Value)
    $parsed = [guid]::Empty
    $Value -is [string] -and [guid]::TryParse($Value, [ref]$parsed) -and $parsed -ne [guid]::Empty
}

function Test-ConsumptionNumber {
    param([object]$Value)
    ($Value -is [double] -or $Value -is [single] -or $Value -is [decimal] -or
        $Value -is [int] -or $Value -is [long]) -and
        [double]::IsFinite([double]$Value) -and $Value -ge 0
}

function ConvertTo-ConsumptionUtcTime {
    param([object]$Value)
    if ($Value -is [DateTimeOffset]) { return $Value.ToUniversalTime().ToString('o') }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { $Value = [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
        return [DateTimeOffset]::new($Value).ToUniversalTime().ToString('o')
    }
    $parsed = [DateTimeOffset]::MinValue
    if ($Value -is [string] -and [DateTimeOffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) {
        return $parsed.ToUniversalTime().ToString('o')
    }
    Stop-ConsumptionOperation 'InvalidSourceTimestamp' 'A source or historical observation timestamp is invalid.'
}

function Get-ConsumptionQuantity {
    param([object]$Value, [string]$Unit)
    if ($null -eq $Value) { return @{ kind = 'missing'; value = $null } }
    if ($Value -is [Collections.IDictionary] -and $Value.unit -ceq $Unit) {
        if ($Value['@odata.type'] -eq '#microsoft.graph.unlimitedCostManagementQuantity') {
            return @{ kind = 'unlimited'; value = $null }
        }
        if ($Value['@odata.type'] -eq '#microsoft.graph.limitedCostManagementQuantity' -and (Test-ConsumptionNumber $Value.value)) {
            return @{ kind = 'limited'; value = $Value.value }
        }
    }
    @{ kind = 'invalid'; value = $null }
}

function New-ConsumptionObservation {
    param([object]$Row, [object]$Member, [string]$GroupId, [object]$Policy, [object]$Manifest)
    $valid = $Row -is [Collections.IDictionary]
    if (-not $valid) { $rowFields = @{} } else { $rowFields = $Row }
    $remaining = Get-ConsumptionQuantity $rowFields.remainingQuantity $rowFields.unit
    $sourceTime = $null
    if ($rowFields.dataStatus -is [Collections.IDictionary] -and $rowFields.dataStatus.lastModifiedDateTime) {
        try { $sourceTime = ConvertTo-ConsumptionUtcTime $rowFields.dataStatus.lastModifiedDateTime }
        catch { $valid = $false }
    }
    $valid = $valid -and $rowFields.serviceId -is [string] -and
        -not [string]::IsNullOrWhiteSpace($rowFields.serviceId) -and
        (Test-ConsumptionId $rowFields.spendingPolicyId) -and
        (Test-ConsumptionNumber $rowFields.consumedQuantity) -and
        $rowFields.unit -is [string] -and -not [string]::IsNullOrWhiteSpace($rowFields.unit) -and
        $remaining.kind -ne 'invalid'
    $match = if (-not (Test-ConsumptionId $rowFields.spendingPolicyId)) { 'unknown' } elseif ($rowFields.spendingPolicyId -eq $Policy.id) { 'matched' } else { 'mismatch' }
    [ordered]@{
        schemaVersion = 1; runId = $Manifest.runId; tenantId = $Manifest.tenantId
        collectionStartedAtUtc = $Manifest.startedAtUtc; observedAtUtc = (Get-ConsumptionTime)
        discoveryPolicyId = $Policy.id; discoveryPolicyName = $Policy.displayName
        sourceAssignedGroupId = $GroupId; userId = $Member.id
        userDisplayName = $Member.displayName; userPrincipalName = $Member.userPrincipalName
        serviceId = $rowFields.serviceId; returnedSpendingPolicyId = $rowFields.spendingPolicyId
        policyMatch = $match; validationStatus = $(if ($valid) { 'valid' } else { 'invalid' })
        consumedQuantity = $(if (Test-ConsumptionNumber $rowFields.consumedQuantity) { $rowFields.consumedQuantity } else { $null })
        unit = $rowFields.unit; remainingQuantityKind = $remaining.kind; remainingQuantityValue = $remaining.value
        isConsumptionAllowed = $rowFields.isConsumptionAllowed; accessDecisionReason = $rowFields.accessDecisionReason
        sourceLastModifiedDateTime = $sourceTime; serviceBalance = $Row
    }
}

function ConvertTo-ConsumptionCsvCell {
    param([object]$Value)
    if ($null -eq $Value) { $text = '' }
    elseif ($Value -is [bool]) { $text = $Value.ToString().ToLowerInvariant() }
    elseif ($Value -is [ValueType] -and $Value -is [IFormattable]) {
        $text = $Value.ToString($null, [Globalization.CultureInfo]::InvariantCulture)
    } else {
        $text = [string]$Value
        if ($text -match '^[\s\x00-\x1f]*[=+\-@]') { $text = "'" + $text }
    }
    '"' + $text.Replace('"', '""') + '"'
}

function Write-ConsumptionCsv {
    param([string]$Path, [object[]]$Observations)
    $writer = [IO.StreamWriter]::new($Path, $false, [Text.UTF8Encoding]::new($true))
    try {
        $writer.WriteLine(($script:ConsumptionCsvColumns | ForEach-Object { ConvertTo-ConsumptionCsvCell $_ }) -join ',')
        foreach ($row in $Observations) {
            $writer.WriteLine(($script:ConsumptionCsvColumns | ForEach-Object { ConvertTo-ConsumptionCsvCell $row[$_] }) -join ',')
        }
    } finally { $writer.Dispose() }
}

function Get-ConsumptionPolicyIndex {
    param([object[]]$Observations)
    $index = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    foreach ($row in $Observations) {
        if ($row.validationStatus -ne 'valid') { continue }
        $key = "$($row.userId.ToLowerInvariant())|$($row.serviceId)"
        if (-not $index.ContainsKey($key)) {
            $index[$key] = @{
                userId = $row.userId; serviceId = $row.serviceId; observedAtUtc = $row.observedAtUtc
                policyIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                sourceTimes = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
                missingSourceTime = $false
            }
        }
        $entry = $index[$key]
        [void]$entry.policyIds.Add($row.returnedSpendingPolicyId)
        if ($row.sourceLastModifiedDateTime) { [void]$entry.sourceTimes.Add($row.sourceLastModifiedDateTime) }
        else { $entry.missingSourceTime = $true }
        if ($row.observedAtUtc -gt $entry.observedAtUtc) { $entry.observedAtUtc = $row.observedAtUtc }
    }
    @($index.Values | ForEach-Object {
        @{ userId = $_.userId; serviceId = $_.serviceId; observedAtUtc = $_.observedAtUtc
            policyIds = @($_.policyIds); sourceTimes = @($_.sourceTimes); missingSourceTime = $_.missingSourceTime }
    })
}

function Get-ConsumptionPreviousPolicies {
    param([string]$TenantDirectory, [string]$CurrentRunId)
    $index = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    foreach ($month in @(Get-ChildItem -LiteralPath $TenantDirectory -Directory | Where-Object Name -Match '^\d{4}-\d{2}$')) {
        if ($month.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            Stop-ConsumptionOperation 'HistoryReadFailed' 'History must not traverse symbolic links or junctions.'
        }
        foreach ($run in @(Get-ChildItem -LiteralPath $month.FullName -Directory)) {
            if ($run.Name -eq $CurrentRunId -or ($run.Attributes -band [IO.FileAttributes]::ReparsePoint)) { continue }
            $path = Join-Path $run.FullName 'manifest.json'
            if (-not (Test-Path -LiteralPath $path)) { continue }
            $previous = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -AsHashtable -ErrorAction Stop
            if ($previous.schemaVersion -ne 1) { Stop-ConsumptionOperation 'HistoryReadFailed' 'An earlier manifest has an unsupported schema.' }
            foreach ($entry in @($previous.policyObservations)) {
                $entry.observedAtUtc = ConvertTo-ConsumptionUtcTime $entry.observedAtUtc
                $entry.sourceTimes = @($entry.sourceTimes | ForEach-Object { ConvertTo-ConsumptionUtcTime $_ })
                $key = "$($entry.userId.ToLowerInvariant())|$($entry.serviceId)"
                if (-not $index.ContainsKey($key) -or $entry.observedAtUtc -gt $index[$key].observedAtUtc) { $index[$key] = $entry }
            }
        }
    }
    return ,$index
}

function Compare-ConsumptionPolicies {
    param([object[]]$Current, [Collections.Generic.Dictionary[string, object]]$Previous)
    foreach ($entry in $Current) {
        $key = "$($entry.userId.ToLowerInvariant())|$($entry.serviceId)"
        $before = if ($Previous.ContainsKey($key)) { $Previous[$key] } else { $null }
        if ($entry.policyIds.Count -gt 1 -or ($before -and $before.policyIds.Count -gt 1)) {
            @{ userId = $entry.userId; serviceId = $entry.serviceId; state = 'ambiguous'; reason = 'ConflictingPolicyObservations' }
        } elseif ($before -and $entry.policyIds[0] -ne $before.policyIds[0]) {
            $reliable = -not $entry.missingSourceTime -and -not $before.missingSourceTime -and
                $entry.sourceTimes.Count -gt 0 -and $before.sourceTimes.Count -gt 0 -and
                (@($entry.sourceTimes | Sort-Object)[0] -gt @($before.sourceTimes | Sort-Object)[-1])
            @{ userId = $entry.userId; serviceId = $entry.serviceId
                state = $(if ($reliable) { 'observedChange' } else { 'ambiguous' })
                reason = $(if ($reliable) { 'DifferentReportedPolicy' } else { 'InsufficientOrNonIncreasingSourceFreshness' })
                previousPolicyId = $before.policyIds[0]; currentPolicyId = $entry.policyIds[0]
                previousObservedAtUtc = $before.observedAtUtc; currentObservedAtUtc = $entry.observedAtUtc }
        }
    }
}

function Add-ConsumptionError {
    param([object]$Manifest, [object]$Policy, [object]$Failure, [string]$Stage, [string]$ResourceId)
    $errorEntry = @{ stage = $Stage; resourceId = $ResourceId; failure = $Failure }
    $Manifest.errors.Add($errorEntry)
    if ($Policy) { $Policy.errors.Add($errorEntry); $Policy.collectionStatus = 'partial' }
}

function Invoke-ConsumptionCollectionRun {
    param([hashtable]$Config, [string]$RunDirectory, [object]$Manifest)
    $allRows = [Collections.Generic.List[object]]::new()
    $tenantDirectory = Split-Path (Split-Path $RunDirectory -Parent) -Parent
    $previous = Get-ConsumptionPreviousPolicies $tenantDirectory $Manifest.runId
    $inventory = Get-ConsumptionCollection '/beta/copilot/costManagement/spendingPolicies' $Config
    if (-not $inventory.complete) { Add-ConsumptionError $Manifest $null $inventory.error 'policyInventory' '' }
    foreach ($policy in $inventory.items) {
        if ($policy -isnot [Collections.IDictionary] -or -not (Test-ConsumptionId $policy.id)) {
            Add-ConsumptionError $Manifest $null @{ category = 'InvalidPolicy'; message = 'Policy inventory contained an invalid ID.' } 'policyInventory' ''
            continue
        }
        $policyFile = 'policy-' + ([guid]$policy.id).ToString() + '.json'
        if ($policyFile -in $Manifest.policyFiles) {
            Add-ConsumptionError $Manifest $null @{ category = 'RepeatedPolicy'; message = 'Policy inventory repeated a policy ID.' } 'policyInventory' $policy.id
            continue
        }
        $Manifest.policyFiles.Add($policyFile)
        $document = [ordered]@{
            schemaVersion = 1; runId = $Manifest.runId; tenantId = $Config.tenantId
            startedAtUtc = $Manifest.startedAtUtc; completedAtUtc = $null; policy = $policy
            collectionStatus = 'collecting'; exclusion = $null
            groups = [Collections.Generic.List[object]]::new()
            observations = [Collections.Generic.List[object]]::new()
            errors = [Collections.Generic.List[object]]::new()
        }
        $policyPath = Join-Path $RunDirectory $policyFile
        Write-ConsumptionJson $policyPath $document
        if ($policy['@odata.type'] -ne '#microsoft.graph.selectedGroupsSpendingPolicy') {
            $document.collectionStatus = 'excluded'
            $document.exclusion = if ($policy['@odata.type'] -eq '#microsoft.graph.allUsersSpendingPolicy') { 'TenantWidePolicyNotCollected' } else { 'UnsupportedPolicyType' }
            $Manifest.exclusions.Add(@{ policyId = $policy.id; reason = $document.exclusion })
        } else {
            $document.collectionStatus = 'complete'
            $assignments = Get-ConsumptionCollection "/beta/copilot/costManagement/spendingPolicies/$($policy.id)/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups" $Config
            if (-not $assignments.complete) { Add-ConsumptionError $Manifest $document $assignments.error 'assignedGroups' $policy.id }
            foreach ($group in $assignments.items) {
                if ($group -isnot [Collections.IDictionary] -or -not (Test-ConsumptionId $group.id)) {
                    Add-ConsumptionError $Manifest $document @{ category = 'InvalidGroup'; message = 'Assigned group had an invalid ID.' } 'assignedGroups' $policy.id
                    continue
                }
                $members = Get-ConsumptionCollection "/v1.0/groups/$($group.id)/transitiveMembers" $Config
                $groupResult = @{ group = $group; complete = $members.complete; skippedNonUserObjects = 0; users = [Collections.Generic.List[object]]::new() }
                $document.groups.Add($groupResult)
                $Manifest.groupOccurrences++
                if (-not $members.complete) { Add-ConsumptionError $Manifest $document $members.error 'transitiveMembers' $group.id }
                foreach ($member in $members.items) {
                    if ($member -isnot [Collections.IDictionary] -or -not (Test-ConsumptionId $member.id) -or -not $member['@odata.type']) {
                        Add-ConsumptionError $Manifest $document @{ category = 'InvalidMember'; message = 'Membership contained an invalid directory object.' } 'transitiveMembers' $group.id
                        continue
                    }
                    if ($member['@odata.type'] -ne '#microsoft.graph.user') { $groupResult.skippedNonUserObjects++; continue }
                    $Manifest.discoveredUserOccurrences++
                    $balances = Get-ConsumptionCollection "/beta/copilot/costManagement/userBalances/$($member.id)/serviceBalances" $Config
                    $groupResult.users.Add(@{ user = $member; complete = $balances.complete; serviceRowCount = $balances.items.Count })
                    if ($balances.complete) { $Manifest.successfulUserOccurrences++ }
                    else { Add-ConsumptionError $Manifest $document $balances.error 'serviceBalances' $member.id }
                    $services = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
                    foreach ($balance in $balances.items) {
                        $row = New-ConsumptionObservation $balance $member $group.id $policy $Manifest
                        if ($row.serviceId -and -not $services.Add($row.serviceId)) { $row.validationStatus = 'invalid' }
                        if ($row.validationStatus -ne 'valid') {
                            Add-ConsumptionError $Manifest $document @{ category = 'InvalidServiceBalance'; message = 'A service balance had missing/invalid fields, freshness or quantities, or repeated a service.' } 'serviceBalances' $member.id
                        }
                        $document.observations.Add($row)
                        $allRows.Add($row)
                        if ($row.policyMatch -eq 'mismatch') { $Manifest.policyMismatchRows++ }
                    }
                }
                Write-ConsumptionJson $policyPath $document
            }
        }
        $document.completedAtUtc = Get-ConsumptionTime
        Write-ConsumptionJson $policyPath $document
        Write-ConsumptionJson (Join-Path $RunDirectory 'manifest.json') $Manifest
    }
    $Manifest.policyObservations = @(Get-ConsumptionPolicyIndex $allRows.ToArray())
    $Manifest.policyChanges = @(Compare-ConsumptionPolicies $Manifest.policyObservations $previous)
    $Manifest.csvRowCount = $allRows.Count
    Write-ConsumptionCsv (Join-Path $RunDirectory 'consumption.csv') $allRows.ToArray()
    $Manifest.completedAtUtc = Get-ConsumptionTime
    if ($Manifest.startedAtUtc.Substring(0, 7) -ne $Manifest.completedAtUtc.Substring(0, 7)) {
        Add-ConsumptionError $Manifest $null @{ category = 'MixedReportingPeriod'; message = 'Collection crossed a UTC month boundary.' } 'publication' ''
    }
    $Manifest.status = if ($Manifest.errors.Count) { 'partial' } else { 'complete' }
    Write-ConsumptionJson (Join-Path $RunDirectory 'manifest.json') $Manifest
    $Manifest
}

function Publish-ConsumptionRun {
    param([string]$TenantDirectory, [string]$RunDirectory, [object]$Manifest)
    if ($Manifest.status -ne 'complete' -or -not [IO.File]::Exists((Join-Path $RunDirectory 'consumption.csv'))) {
        Stop-ConsumptionOperation 'PublicationFailed' 'Only a complete JSON/CSV run may be published.'
    }
    foreach ($file in $Manifest.policyFiles) {
        if ($file -notmatch '^policy-[a-fA-F0-9-]{36}\.json$' -or -not [IO.File]::Exists((Join-Path $RunDirectory $file))) {
            Stop-ConsumptionOperation 'PublicationFailed' 'A required policy snapshot is missing.'
        }
    }
    Write-ConsumptionJson (Join-Path $TenantDirectory 'latest-complete.json') @{
        schemaVersion = 1; runId = $Manifest.runId; tenantId = $Manifest.tenantId
        collectionScope = 'groupScopedPolicies'; completedAtUtc = $Manifest.completedAtUtc
        manifestPath = [IO.Path]::GetRelativePath($TenantDirectory, (Join-Path $RunDirectory 'manifest.json'))
        csvPath = [IO.Path]::GetRelativePath($TenantDirectory, (Join-Path $RunDirectory 'consumption.csv'))
    }
}

function New-ConsumptionManifest {
    param([hashtable]$Config, [string]$RunId)
    [ordered]@{
        schemaVersion = 1; csvSchemaVersion = 1; tenantId = $Config.tenantId; runId = $RunId
        startedAtUtc = (Get-ConsumptionTime); completedAtUtc = $null; collectionScope = 'groupScopedPolicies'
        status = 'collecting'; policyFiles = [Collections.Generic.List[string]]::new()
        csvFile = 'consumption.csv'; csvRowCount = 0; groupOccurrences = 0
        discoveredUserOccurrences = 0; successfulUserOccurrences = 0; policyMismatchRows = 0
        errors = [Collections.Generic.List[object]]::new(); exclusions = [Collections.Generic.List[object]]::new()
        policyObservations = @(); policyChanges = @()
    }
}

function Invoke-ConsumptionWorker {
    param([hashtable]$Config)
    $runDirectory = $env:COPILOT_EXPORT_RUN
    $manifest = New-ConsumptionManifest $Config (Split-Path $runDirectory -Leaf)
    Write-ConsumptionJson (Join-Path $runDirectory 'manifest.json') $manifest
    try {
        Connect-ConsumptionGraph $Config
        $manifest = Invoke-ConsumptionCollectionRun $Config $runDirectory $manifest
        Write-ConsumptionEvent @{ type = 'result'; status = $manifest.status; runId = $manifest.runId }
        if ($manifest.status -eq 'complete') { return 0 } else { return 1 }
    } catch {
        $manifest.status = 'failed'
        $manifest.completedAtUtc = Get-ConsumptionTime
        $failure = Get-ConsumptionFailure $_
        Add-ConsumptionError $manifest $null $failure 'worker' ''
        Write-ConsumptionJson (Join-Path $runDirectory 'manifest.json') $manifest
        Write-ConsumptionEvent @{ type = 'result'; status = 'failed'; runId = $manifest.runId; failure = $failure }
        return 1
    }
}

function Wait-ConsumptionWorker {
    param([Diagnostics.Process]$Process, [hashtable]$Config)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $phase = 'authentication'
    $result = $null
    $lineTask = $Process.StandardOutput.ReadLineAsync()
    $stderrTask = $Process.StandardError.ReadToEndAsync()
    try {
        while ($true) {
            if ($lineTask.IsCompleted) {
                $line = $lineTask.GetAwaiter().GetResult()
                if ($null -eq $line) { break }
                if (-not $line.StartsWith('COPILOT_EXPORT:')) {
                    Stop-ConsumptionOperation 'ReauthenticationRequired' 'The SDK requested interaction or produced unexpected console output. Export was stopped without displaying the sign-in code. Run Setup interactively.'
                }
                $event = $line.Substring(15) | ConvertFrom-Json -AsHashtable -ErrorAction Stop
                if ($event.type -eq 'phase') { $phase = $event.phase; $clock.Restart() }
                elseif ($event.type -eq 'result') { $result = $event }
                else { Stop-ConsumptionOperation 'WorkerProtocolFailed' 'The export worker returned an unknown event.' }
                $lineTask = $Process.StandardOutput.ReadLineAsync()
            }
            $limit = if ($phase -eq 'authentication') { $Config.authenticationTimeoutSeconds } elseif ($phase -eq 'request') { $Config.requestTimeoutSeconds + 15 } else { 300 }
            if ($clock.Elapsed.TotalSeconds -gt $limit) {
                Stop-ConsumptionOperation 'WorkerTimeout' "The export worker timed out during $phase. Check connectivity and run Setup if authentication is required."
            }
            Start-Sleep -Milliseconds 25
        }
        if (-not $Process.WaitForExit(5000)) { Stop-ConsumptionOperation 'WorkerTimeout' 'The export worker did not exit after closing its output.' }
        $stderr = $stderrTask.GetAwaiter().GetResult()
        if (-not $result -or $stderr -or ($Process.ExitCode -ne 0 -and $result.status -eq 'complete')) {
            Stop-ConsumptionOperation 'WorkerFailed' 'The export worker failed outside its structured result. Check the PowerShell/module installation.'
        }
        $result
    } finally {
        if (-not $Process.HasExited) { $Process.Kill($true); $Process.WaitForExit() }
    }
}

function Invoke-ConsumptionExport {
    param([hashtable]$Config)
    Assert-ConsumptionConfiguration $Config
    $tenantDirectory = Join-Path ([IO.Path]::GetFullPath($Config.outputPath)) $Config.tenantId
    $tenantConfig = $Config.Clone()
    $tenantConfig.outputPath = $tenantDirectory
    Assert-ConsumptionConfiguration $tenantConfig
    [void][IO.Directory]::CreateDirectory($tenantDirectory)
    $lock = $null
    try {
        try { $lock = [IO.File]::Open((Join-Path $tenantDirectory 'export.lock'), 'OpenOrCreate', 'ReadWrite', 'None') }
        catch { Stop-ConsumptionOperation 'ExportAlreadyRunning' 'Could not acquire the tenant export lock. Check for an active export or inaccessible output.' }
        $runId = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHHmmssfffffffZ') + '-' + [guid]::NewGuid().ToString('N')
        $runDirectory = Join-Path (Join-Path $tenantDirectory $runId.Substring(0, 7)) $runId
        $runConfig = $Config.Clone()
        $runConfig.outputPath = $runDirectory
        Assert-ConsumptionConfiguration $runConfig
        [void][IO.Directory]::CreateDirectory($runDirectory)
        $manifestPath = Join-Path $runDirectory 'manifest.json'
        Write-ConsumptionJson $manifestPath (New-ConsumptionManifest $Config $runId)
        $process = [Diagnostics.Process]::new()
        try {
            $process.StartInfo.FileName = (Get-Process -Id $PID).Path
            $process.StartInfo.UseShellExecute = $false
            $process.StartInfo.RedirectStandardOutput = $true
            $process.StartInfo.RedirectStandardError = $true
            foreach ($argument in @('-NoProfile', '-NonInteractive', '-File', $script:ConsumptionScriptPath, '-Mode', 'Worker',
                '-TenantId', $Config.tenantId, '-ClientId', $Config.clientId, '-Account', $Config.account,
                '-OutputPath', $Config.outputPath, '-RequestTimeoutSeconds', $Config.requestTimeoutSeconds.ToString(),
                '-MaxRetryCount', $Config.maxRetryCount.ToString())) { $process.StartInfo.ArgumentList.Add($argument) }
            $process.StartInfo.Environment['COPILOT_EXPORT_RUN'] = $runDirectory
            [void]$process.Start()
            $result = Wait-ConsumptionWorker $process $Config
            if ($result.status -eq 'complete') {
                $completedManifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -AsHashtable
                if ($completedManifest.tenantId -ne $Config.tenantId -or $completedManifest.runId -ne $runId) {
                    Stop-ConsumptionOperation 'PublicationFailed' 'The completed manifest does not identify the expected tenant and run.'
                }
                Publish-ConsumptionRun $tenantDirectory $runDirectory $completedManifest
            }
        } catch {
            $failure = Get-ConsumptionFailure $_
            $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json -AsHashtable
            $manifest.status = 'failed'
            $manifest.completedAtUtc = Get-ConsumptionTime
            $manifest.errors = @($manifest.errors) + @{ stage = 'supervisor'; resourceId = ''; failure = $failure }
            Write-ConsumptionJson $manifestPath $manifest
            $result = @{ status = 'failed'; runId = $runId; failure = $failure }
        } finally { $process.Dispose() }
        $result['schemaVersion'] = 1
        $result['tenantId'] = $Config.tenantId
        $result['recordedAtUtc'] = Get-ConsumptionTime
        $result['manifestPath'] = $manifestPath
        Write-ConsumptionJson (Join-Path $tenantDirectory 'latest-status.json') $result
        $result
    } finally { if ($lock) { $lock.Dispose() } }
}

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    $config = @{
        tenantId = $TenantId.ToString(); clientId = $ClientId.ToString(); account = $Account
        outputPath = $OutputPath; authenticationTimeoutSeconds = $AuthenticationTimeoutSeconds
        requestTimeoutSeconds = $RequestTimeoutSeconds; maxRetryCount = $MaxRetryCount
    }
    try {
        Assert-ConsumptionConfiguration $config
        if ($Mode -eq 'Worker') {
            if (-not $env:COPILOT_EXPORT_RUN) { Stop-ConsumptionOperation 'InvalidConfiguration' 'Worker is an internal mode; use Export.' }
            $script:ConsumptionProtocol = $true
            exit (Invoke-ConsumptionWorker $config)
        }
        if (-not $IsWindows) { Stop-ConsumptionOperation 'UnsupportedPlatform' 'Live setup and export require Windows OS-protected Graph authentication.' }
        if ($Mode -eq 'Setup') {
            Connect-ConsumptionGraph $config
            Write-Output 'Persistent delegated authentication is ready. Test Export in a fresh process and in the actual scheduled-task logon context.'
        } else {
            $result = Invoke-ConsumptionExport $config
            $result | ConvertTo-Json -Depth 10
            if ($result.status -ne 'complete') { exit 1 }
        }
    } catch {
        Get-ConsumptionFailure $_ | ConvertTo-Json -Depth 10 | Write-Error
        exit 1
    }
}
