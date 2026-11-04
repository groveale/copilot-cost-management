#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$SiteUrl,
    [string]$RuntimeAppId,
    [ValidateRange(1, 3650)][int]$RetentionDays = 30,
    [ValidateRange(1, 8)][int]$DefaultConcurrency = 3,
    [ValidateRange(1, 8)][int]$MaxConcurrency = 8,
    [switch]$GrantSiteAccess,
    [switch]$LockDownReporting,
    [string[]]$AdminObjectIds = @(),
    [switch]$Apply
)

function New-SyncColumn {
    param([string]$Name, [string]$Kind = 'text', [switch]$Indexed, [switch]$Unique)
    $column = @{ name = $Name; indexed = [bool]($Indexed -or $Unique); enforceUniqueValues = [bool]$Unique }
    switch ($Kind) {
        'text' { $column.text = @{ allowMultipleLines = $false; maxLength = 255 } }
        'json' { $column.text = @{ allowMultipleLines = $true; appendChangesToExistingText = $false; textType = 'plain' } }
        'number' { $column.number = @{ decimalPlaces = 'automatic' } }
        'dateTime' { $column.dateTime = @{ displayAs = 'default'; format = 'dateTime' } }
        'boolean' { $column.boolean = @{} }
        default { throw "Unknown column kind: $Kind." }
    }
    return $column
}

function Get-SyncListDefinitions {
    $policy = @(
        New-SyncColumn PolicyKey -Unique
        New-SyncColumn TenantId -Indexed
        New-SyncColumn PolicyId
        New-SyncColumn DisplayName
        New-SyncColumn PolicyType
        New-SyncColumn ActiveSnapshotId
        New-SyncColumn JobId
        New-SyncColumn FenceId
        New-SyncColumn Presence
        New-SyncColumn Configuration json
        New-SyncColumn Totals json
        New-SyncColumn Coverage json
        New-SyncColumn StartedAt dateTime
        New-SyncColumn ObservedAt dateTime
        New-SyncColumn LastInventoryAt dateTime
        New-SyncColumn SchemaVersion number
    )
    $balances = @(
        New-SyncColumn ObservationKey -Unique
        New-SyncColumn SnapshotKey -Indexed
        New-SyncColumn PolicyKey -Indexed
        New-SyncColumn TenantId -Indexed
        New-SyncColumn PolicyId
        New-SyncColumn SnapshotId
        New-SyncColumn JobId
        New-SyncColumn UserId
        New-SyncColumn ServiceId
        New-SyncColumn DisplayName
        New-SyncColumn UserPrincipalName
        New-SyncColumn Groups json
        New-SyncColumn ReportedPolicyId
        New-SyncColumn PolicyMatch
        New-SyncColumn ConsumedQuantity number
        New-SyncColumn Unit
        New-SyncColumn RemainingKind
        New-SyncColumn RemainingValue number
        New-SyncColumn ConsumptionAllowed
        New-SyncColumn AccessReason
        New-SyncColumn SourceTime dateTime
        New-SyncColumn ObservedAt dateTime
        New-SyncColumn SchemaVersion number
    )
    $jobs = @(
        New-SyncColumn JobId -Unique
        New-SyncColumn TenantId -Indexed
        New-SyncColumn State
        New-SyncColumn Revision number
        New-SyncColumn ScopePolicyId
        New-SyncColumn Concurrency number
        New-SyncColumn InitiatorId
        New-SyncColumn ExecutorId
        New-SyncColumn Actors json
        New-SyncColumn CreatedAt dateTime
        New-SyncColumn UpdatedAt dateTime
        New-SyncColumn FinishedAt dateTime
        New-SyncColumn Phase
        New-SyncColumn PolicyCount number
        New-SyncColumn CompletedCount number
        New-SyncColumn FailedCount number
        New-SyncColumn Outcomes json
        New-SyncColumn SchemaVersion number
    )
    @(
        @{ key = 'policyListId'; displayName = 'CopilotSpendingPolicies'; columns = $policy }
        @{ key = 'balanceListId'; displayName = 'CopilotUserBalances'; columns = $balances }
        @{ key = 'jobListId'; displayName = 'CopilotSyncJobs'; columns = $jobs }
    )
}

function Invoke-SyncGraph {
    param([string]$Method = 'GET', [string]$Path, [object]$Body)
    $uri = [uri]::new([uri]'https://graph.microsoft.com', $Path)
    if ($uri.Scheme -ne 'https' -or $uri.Host -ne 'graph.microsoft.com' -or $uri.Port -ne 443 -or
        $uri.UserInfo -or $uri.Fragment -or $uri.AbsolutePath -notmatch '^/v1\.0/') {
        throw 'Refusing a request outside the expected Microsoft Graph endpoint.'
    }
    $arguments = @{ Method = $Method; Uri = $uri.AbsoluteUri; OutputType = 'Hashtable'; ErrorAction = 'Stop' }
    if ($null -ne $Body) { $arguments.Body = ConvertTo-Json -InputObject $Body -Depth 30; $arguments.ContentType = 'application/json' }
    Invoke-MgGraphRequest @arguments
}

function Get-SyncPages {
    param([string]$Path)
    $visited = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    while ($Path) {
        if (-not $visited.Add($Path)) { throw 'Graph returned a repeated continuation link.' }
        $page = Invoke-SyncGraph -Path $Path
        if (-not $page.Contains('value') -or $page.value -isnot [array]) { throw 'Graph did not return a collection.' }
        foreach ($row in $page.value) { $row }
        $Path = $page['@odata.nextLink']
    }
}

function Test-SyncColumn {
    param([object]$Actual, [object]$Expected)
    $kind = @('text', 'number', 'dateTime', 'boolean') | Where-Object { $Expected.ContainsKey($_) } | Select-Object -First 1
    if (-not $Actual.ContainsKey($kind) -or [bool]$Actual.indexed -ne $Expected.indexed -or
        [bool]$Actual.enforceUniqueValues -ne $Expected.enforceUniqueValues) { return $false }
    if ($kind -eq 'text' -and [bool]$Actual.text.allowMultipleLines -ne $Expected.text.allowMultipleLines) { return $false }
    return $true
}

function Protect-SyncLists {
    [CmdletBinding()]
    param([string]$TenantId, [string]$ClientId, [string]$SiteUrl, [hashtable]$Destination,
        [string[]]$AdminObjectIds, [switch]$Apply)
    if (-not $AdminObjectIds.Count) { throw 'LockDownReporting requires explicit AdminObjectIds matching the private API administrator configuration.' }
    $directoryAdmins = @($AdminObjectIds | ForEach-Object {
        $id = [guid]::Empty
        if (-not [guid]::TryParse($_, [ref]$id) -or $id -eq [guid]::Empty) { throw 'Administrator IDs must be non-zero GUIDs.' }
        $user = Invoke-SyncGraph -Path "/v1.0/users/$id`?`$select=id,userPrincipalName"
        if ($user.id -ne $id.ToString() -or -not $user.userPrincipalName) { throw 'An administrator could not be resolved as a directory user.' }
        [pscustomobject]@{ id = $id.ToString(); upn = $user.userPrincipalName }
    } | Sort-Object -Property id -Unique)
    $admins = @($directoryAdmins.upn)
    Write-Host 'SECURITY MIGRATION: remove all non-administrator list/item grants on the three app-managed lists, including stale/staging rows. No data is deleted. Owners lose access until server reconciliation succeeds.'
    if (-not $Apply) { Write-Host 'Preview only: no SharePoint permission changes. Existing broad access is NOT considered secure.'; return }
    if ($Destination.Count -ne 4) { throw 'Lockdown requires all three provisioned lists.' }
    Import-Module PnP.PowerShell -MinimumVersion 3.1.0 -ErrorAction Stop
    $connection = Connect-PnPOnline -Url $SiteUrl -Tenant $TenantId -ClientId $ClientId -Interactive -ReturnConnection -ErrorAction Stop
    $context = Get-PnPContext -Connection $connection
    $currentUser = Get-PnPProperty -ClientObject $context.Web -Property CurrentUser -Connection $connection
    $adminUsers = @($directoryAdmins | ForEach-Object {
        $user = New-PnPUser -LoginName $_.upn -Connection $connection
        $aad = Get-PnPProperty -ClientObject $user -Property AadObjectId -Connection $connection
        if ($aad.NameId -ne $_.id) { throw 'SharePoint administrator identity does not match the current directory object ID. Repair stale site-user mappings first.' }
        $user
    })
    if ($currentUser.Id -notin $adminUsers.Id) { throw 'Run lockdown as one of the explicitly configured administrators, to avoid losing repair access.' }
    $title = 'CopilotCostSync Administrators'
    $group = @(Get-PnPGroup -Connection $connection | Where-Object Title -CEQ $title)
    if ($group.Count -gt 1) { throw 'Ambiguous reporting administrator group.' }
    if ($group.Count -and $group[0].Description -ne 'CopilotCostSync/v1/administrators') { throw 'Refusing to reuse an unmanaged administrator group.' }
    $group = if ($group.Count) { $group[0] } else {
        New-PnPGroup -Title $title -Description 'CopilotCostSync/v1/administrators' -Owner $admins[0] -Connection $connection
    }
    Set-PnPGroup -Identity $group.Id -Owner $admins[0] -AllowMembersEditMembership $false `
        -AllowRequestToJoinLeave $false -AutoAcceptRequestToJoinLeave $false -OnlyAllowMembersViewMembership $true -Connection $connection
    $verifiedGroup = Get-PnPGroup -Identity $group.Id -Includes Owner,AllowMembersEditMembership,AllowRequestToJoinLeave,AutoAcceptRequestToJoinLeave -Connection $connection
    if ($verifiedGroup.Owner.Id -notin $adminUsers.Id -or $verifiedGroup.AllowMembersEditMembership -ne $false -or
        $verifiedGroup.AllowRequestToJoinLeave -ne $false -or $verifiedGroup.AutoAcceptRequestToJoinLeave -ne $false) {
        throw 'Administrator group ownership/self-join readback failed.'
    }
    foreach ($member in @(Get-PnPGroupMember -Group $group.Id -Connection $connection)) {
        if ($member.Id -notin $adminUsers.Id) { Remove-PnPGroupMember -Group $group.Id -LoginName $member.LoginName -Connection $connection }
    }
    foreach ($admin in $adminUsers) { Add-PnPGroupMember -Group $group.Id -LoginName $admin.LoginName -Connection $connection | Out-Null }
    $members = @(Get-PnPGroupMember -Group $group.Id -Connection $connection)
    if ($members.Count -ne $adminUsers.Count -or @($members | Where-Object Id -NotIn $adminUsers.Id).Count) { throw 'Administrator group readback failed.' }
    $fullControl = @(Get-PnPRoleDefinition -Connection $connection | Where-Object RoleTypeKind -EQ Administrator)
    if ($fullControl.Count -ne 1) { throw 'Full Control role could not be uniquely resolved.' }
    foreach ($definition in Get-SyncListDefinitions) {
        $id = $Destination[$definition.key]
        $list = Get-PnPList -Identity $id -Includes Description,BaseTemplate -Connection $connection
        if ($list.Description -ne "CopilotCostSync/v1/$($definition.key)" -or [int]$list.BaseTemplate -ne 100) { throw 'Refusing to change an unmanaged list.' }
        Set-PnPList -Identity $id -BreakRoleInheritance -CopyRoleAssignments:$false -EnableFolderCreation $false -Connection $connection | Out-Null
        Set-PnPListPermission -Identity $id -Group $group.Id -AddRole $fullControl[0].Name -Connection $connection
        $assignments = @(Get-PnPProperty -ClientObject $list -Property RoleAssignments -Connection $connection)
        foreach ($assignment in $assignments) {
            $member = Get-PnPProperty -ClientObject $assignment -Property Member -Connection $connection
            if ($member.Id -ne $group.Id) { $assignment.DeleteObject() }
        }
        Invoke-PnPQuery -Connection $connection
        $pending = 0
        foreach ($item in @(Get-PnPListItem -List $id -PageSize 500 -Connection $connection)) {
            $item.ResetRoleInheritance()
            $pending++
            if ($pending -eq 100) { Invoke-PnPQuery -Connection $connection; $pending = 0 }
        }
        Invoke-PnPQuery -Connection $connection
        $verified = Get-PnPList -Identity $id -Includes HasUniqueRoleAssignments,RoleAssignments -Connection $connection
        if (-not $verified.HasUniqueRoleAssignments -or $verified.RoleAssignments.Count -ne 1) { throw 'List permission readback failed. Do not enable owner reporting.' }
        $principal = Get-PnPProperty -ClientObject $verified.RoleAssignments[0] -Property Member -Connection $connection
        $bindings = @(Get-PnPProperty -ClientObject $verified.RoleAssignments[0] -Property RoleDefinitionBindings -Connection $connection)
        if ($principal.Id -ne $group.Id -or $bindings.Count -ne 1 -or $bindings[0].Id -ne $fullControl[0].Id) { throw 'List role readback failed.' }
        foreach ($item in @(Get-PnPListItem -List $id -PageSize 500 -Includes HasUniqueRoleAssignments -Connection $connection)) {
            if ($item.HasUniqueRoleAssignments) { throw 'Item permission revocation readback failed. Owner reporting must remain disabled.' }
        }
    }
    Write-Output "REPORTING_SITE_URL=$SiteUrl"
    Write-Output "REPORTING_ADMIN_GROUP_ID=$($group.Id)"
    Write-Host 'Lists are administrator-only. Configure certificate-backed OBO and run the API Verify SharePoint permissions action before enabling owner reporting. Setup never marks security ready.'
}

function Initialize-SyncSharePoint {
    [CmdletBinding()]
    param(
        [string]$TenantId, [string]$ClientId, [string]$SiteUrl, [string]$RuntimeAppId,
        [ValidateRange(1, 3650)][int]$RetentionDays = 30,
        [ValidateRange(1, 8)][int]$DefaultConcurrency = 3,
        [ValidateRange(1, 8)][int]$MaxConcurrency = 8,
        [switch]$GrantSiteAccess, [switch]$LockDownReporting, [string[]]$AdminObjectIds = @(), [switch]$Apply
    )
    if ($DefaultConcurrency -gt $MaxConcurrency) { throw 'DefaultConcurrency cannot exceed MaxConcurrency.' }
    if ($LockDownReporting -and -not $AdminObjectIds.Count) { throw 'LockDownReporting requires explicit AdminObjectIds.' }
    foreach ($value in @($TenantId, $ClientId)) {
        $parsed = [guid]::Empty
        if (-not [guid]::TryParse($value, [ref]$parsed) -or $parsed -eq [guid]::Empty) { throw 'Supply non-zero TenantId and ClientId GUIDs.' }
    }
    $site = [uri]$SiteUrl
    if (-not $site.IsAbsoluteUri -or $site.Scheme -ne 'https' -or $site.Host -notmatch '^[a-z0-9-]+\.sharepoint\.com$' -or
        $site.UserInfo -or $site.Query -or $site.Fragment -or $site.Port -ne 443) { throw 'SiteUrl must be an existing HTTPS SharePoint site URL.' }
    if ($GrantSiteAccess) {
        $parsed = [guid]::Empty
        if (-not [guid]::TryParse($RuntimeAppId, [ref]$parsed) -or $parsed -eq [guid]::Empty) { throw 'GrantSiteAccess requires the runtime SPA application GUID.' }
    }
    Import-Module Microsoft.Graph.Authentication -MinimumVersion 2.37.0 -ErrorAction Stop
    $scopes = if ($GrantSiteAccess) { @('Sites.FullControl.All') } else { @('Sites.Manage.All') }
    if ($LockDownReporting) { $scopes += 'User.ReadBasic.All' }
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Scopes $scopes -ContextScope Process -NoWelcome -ErrorAction Stop | Out-Null
    $context = Get-MgContext
    if ($context.AuthType -ne 'Delegated' -or $context.TenantId -ne $TenantId -or $context.ClientId -ne $ClientId) {
        throw 'Setup requires the expected delegated tenant and operator application.'
    }
    $sitePath = $site.AbsolutePath.TrimEnd('/')
    $lookup = if ($sitePath) { "$($site.Host):$sitePath" } else { $site.Host }
    $resolved = Invoke-SyncGraph -Path "/v1.0/sites/$lookup"
    if (-not $resolved.id) { throw 'The target SharePoint site could not be resolved.' }
    $root = "/v1.0/sites/$($resolved.id)"
    $lists = @(Get-SyncPages "$root/lists")
    $destination = @{ siteId = $resolved.id }
    foreach ($definition in Get-SyncListDefinitions) {
        $matching = @($lists | Where-Object displayName -CEQ $definition.displayName)
        if ($matching.Count -gt 1) { throw "Multiple lists named $($definition.displayName)." }
        $existing = if ($matching.Count) { $matching[0] } else { $null }
        $description = "CopilotCostSync/v1/$($definition.key)"
        if ($existing -and $existing.description -cne $description) {
            throw "Existing list $($definition.displayName) is not a compatible sync list. No data was changed."
        }
        if (-not $existing) {
            Write-Host "Create list $($definition.displayName) with $($definition.columns.Count) columns."
            if ($Apply) {
                $existing = Invoke-SyncGraph -Method POST -Path "$root/lists" -Body @{
                    displayName = $definition.displayName; description = $description
                    list = @{ template = 'genericList' }; columns = $definition.columns
                }
            }
        }
        if ($existing) {
            $columns = @(Get-SyncPages "$root/lists/$($existing.id)/columns")
            foreach ($column in $definition.columns) {
                $found = @($columns | Where-Object name -CEQ $column.name)
                if ($found.Count -gt 1 -or ($found.Count -eq 1 -and -not (Test-SyncColumn $found[0] $column))) {
                    throw "Incompatible column $($column.name) in $($definition.displayName). Repair it explicitly; setup does not recreate lists."
                }
                if (-not $found.Count) {
                    Write-Host "Add column $($column.name) to $($definition.displayName)."
                    if ($Apply) { Invoke-SyncGraph -Method POST -Path "$root/lists/$($existing.id)/columns" -Body $column | Out-Null }
                }
            }
            $destination[$definition.key] = $existing.id
        }
    }
    if ($GrantSiteAccess) {
        $permissions = @(Get-SyncPages "$root/permissions")
        $grant = @($permissions | Where-Object {
            $ids = @($_.grantedToIdentitiesV2) + @($_.grantedToIdentities)
            @($ids | Where-Object { $_.application.id -eq $RuntimeAppId }).Count -gt 0
        })
        if ($grant.Count -gt 1) { throw 'Multiple runtime site grants exist; reconcile them explicitly.' }
        if ($grant.Count -eq 1 -and @($grant[0].roles | Where-Object { $_ -notin @('read', 'write') }).Count) {
            throw 'Runtime app has an elevated site grant. Review it explicitly rather than preserving excessive access.'
        }
        if (-not $grant.Count -or 'write' -notin $grant[0].roles) {
            Write-Host "Grant delegated runtime app $RuntimeAppId write access to this site."
            if ($Apply) {
                if ($grant.Count) {
                    Invoke-SyncGraph -Method PATCH -Path "$root/permissions/$($grant[0].id)" -Body @{ roles = @('write') } | Out-Null
                } else {
                    Invoke-SyncGraph -Method POST -Path "$root/permissions" -Body @{
                        roles = @('write'); grantedToIdentities = @(@{ application = @{ id = $RuntimeAppId; displayName = 'Copilot sync runtime' } })
                    } | Out-Null
                }
            }
        }
    }
    if ($LockDownReporting) {
        Protect-SyncLists -TenantId $TenantId -ClientId $ClientId -SiteUrl $SiteUrl -Destination $destination -AdminObjectIds $AdminObjectIds -Apply:$Apply
    } else { Write-Host 'WARNING: lists may inherit site access. Owner reporting remains disabled until explicit -LockDownReporting migration and API permission readback. Use a restricted site.' }
    Write-Host 'Configure list version history/retention explicitly in SharePoint.'
    if (-not $Apply) { Write-Host 'Preview only. Rerun with -Apply to provision; no Graph writes were made.' }
    if ($destination.Count -eq 4) {
        Write-Output ('SYNC_DESTINATION_JSON=' + (ConvertTo-Json -InputObject $destination -Compress))
        Write-Output "SYNC_DEFAULT_CONCURRENCY=$DefaultConcurrency"
        Write-Output "SYNC_MAX_CONCURRENCY=$MaxConcurrency"
        Write-Output "SYNC_JOB_RETENTION_DAYS=$RetentionDays"
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    $ErrorActionPreference = 'Stop'
    Initialize-SyncSharePoint @PSBoundParameters
}
