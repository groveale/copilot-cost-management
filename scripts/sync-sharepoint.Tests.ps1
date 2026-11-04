#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.7.1' }

BeforeAll {
    . (Join-Path $PSScriptRoot 'setup-sync-sharepoint.ps1') `
        -TenantId '11111111-1111-4111-8111-111111111111' `
        -ClientId '22222222-2222-4222-8222-222222222222' `
        -SiteUrl 'https://example.sharepoint.com/sites/reporting'
    $connection = @{
        TenantId = '11111111-1111-4111-8111-111111111111'
        ClientId = '22222222-2222-4222-8222-222222222222'
        SiteUrl = 'https://example.sharepoint.com/sites/reporting'
    }
    function Connect-MgGraph { param($TenantId, $ClientId, [string[]]$Scopes, $ContextScope, [switch]$NoWelcome, $ErrorAction) }
    function Get-MgContext { throw 'Unexpected context lookup.' }
    function Invoke-MgGraphRequest { param($Method, $Uri, $OutputType, $ErrorAction, $Body, $ContentType); throw 'Unexpected request.' }
    function Connect-PnPOnline { param($Url, $Tenant, $ClientId, [switch]$Interactive, [switch]$ReturnConnection, $ErrorAction) }
    function Get-PnPContext { param($Connection) }
    function Get-PnPProperty { param($ClientObject, $Property, $Connection) }
    function Get-PnPGroup { param($Identity, $Includes, $Connection) }
    function New-PnPGroup { param($Title, $Description, $Owner, $Connection) }
    function Set-PnPGroup { param($Identity, $Owner, $AllowMembersEditMembership, $AllowRequestToJoinLeave, $AutoAcceptRequestToJoinLeave, $OnlyAllowMembersViewMembership, $Connection) }
    function Get-PnPGroupMember { param($Group, $Connection) }
    function New-PnPUser { param($LoginName, $Connection) }
    function Add-PnPGroupMember { param($Group, $LoginName, $Connection) }
    function Remove-PnPGroupMember { param($Group, $LoginName, $Connection) }
    function Get-PnPRoleDefinition { param($Connection) }
    function Get-PnPList { param($Identity, $Includes, $Connection) }
    function Set-PnPList { param($Identity, [switch]$BreakRoleInheritance, [switch]$CopyRoleAssignments, $EnableFolderCreation, $Connection) }
    function Set-PnPListPermission { param($Identity, $Group, $AddRole, $Connection) }
    function Get-PnPListItem { param($List, $PageSize, $Includes, $Connection) }
    function Invoke-PnPQuery { param($Connection) }
}

Describe 'SharePoint sync setup schema' {
    It 'defines three lists with unique keys, indexed partition fields and explicit quantity/job columns' {
        $definitions = @(Get-SyncListDefinitions)
        $definitions.Count | Should -Be 3
        foreach ($definition in $definitions) {
            @($definition.columns | Where-Object enforceUniqueValues -EQ $true).Count | Should -Be 1
            @($definition.columns | Where-Object { $_.name -eq 'TenantId' -and $_.indexed }).Count | Should -Be 1
        }
        ($definitions[1].columns.name) | Should -Contain 'RemainingKind'
        ($definitions[1].columns.name) | Should -Contain 'SnapshotId'
        ($definitions[2].columns.name) | Should -Contain 'Revision'
        ($definitions[2].columns.name) | Should -Not -Contain 'AccessToken'
    }
    It 'detects incompatible column types and uniqueness without silently modifying them' {
        $expected = New-SyncColumn PolicyKey -Unique
        Test-SyncColumn @{ name = 'PolicyKey'; text = @{ allowMultipleLines = $false }; indexed = $true; enforceUniqueValues = $true } $expected | Should -BeTrue
        Test-SyncColumn @{ name = 'PolicyKey'; text = @{ allowMultipleLines = $true }; indexed = $true; enforceUniqueValues = $true } $expected | Should -BeFalse
        Test-SyncColumn @{ name = 'PolicyKey'; number = @{}; indexed = $true; enforceUniqueValues = $true } $expected | Should -BeFalse
    }
}
Describe 'SharePoint sync setup preview and idempotence' {
    BeforeEach {
        $script:testLists = [Collections.Generic.List[object]]::new()
        $script:testGrants = [Collections.Generic.List[object]]::new()
        $script:requestedScopes = [Collections.Generic.List[object]]::new()
        $script:writes = 0
        Mock Import-Module {}
        Mock Connect-MgGraph { param($Scopes); $script:requestedScopes.Add(@($Scopes)) }
        Mock Get-MgContext { @{ AuthType = 'Delegated'; TenantId = '11111111-1111-4111-8111-111111111111'; ClientId = '22222222-2222-4222-8222-222222222222' } }
        Mock Invoke-SyncGraph {
            param($Method, $Path, $Body)
            if ($Method -eq 'POST' -and $Path.EndsWith('/lists')) {
                $script:writes++
                $row = @{ id = [guid]::NewGuid().ToString(); displayName = $Body.displayName; description = $Body.description; columns = $Body.columns }
                $script:testLists.Add($row)
                return $row
            }
            if ($Path.EndsWith('/lists')) { return @{ value = @($script:testLists.ToArray()) } }
            if ($Path.EndsWith('/permissions')) {
                if ($Method -eq 'POST') {
                    $Body.roles | Should -Contain 'write'
                    $row = @{ id = 'grant'; roles = $Body.roles; grantedToIdentities = $Body.grantedToIdentities }
                    $script:testGrants.Add($row)
                    return $row
                }
                return @{ value = @($script:testGrants.ToArray()) }
            }
            if ($Path.EndsWith('/columns')) {
                $list = $script:testLists | Where-Object { $Path.Contains($_.id) }
                return @{ value = @($list.columns) }
            }
            if ($Path -match '/sites/example\.sharepoint\.com:') { return @{ id = 'example.sharepoint.com,11111111-1111-4111-8111-111111111111,33333333-3333-4333-8333-333333333333' } }
            throw "Unexpected setup call: $Method $Path"
        }
    }
    It 'makes no provisioning writes in preview' {
        Initialize-SyncSharePoint @connection
        $script:writes | Should -Be 0
        Should -Invoke Connect-MgGraph -Times 1 -Exactly
        ($script:requestedScopes[0] -join ',') | Should -Be 'Sites.Selected'
        Should -Invoke Invoke-SyncGraph -Times 0 -ParameterFilter { $Path -like '*/permissions*' }
    }
    It 'requests selected site access and basic user reads for permission lockdown without tenant-wide site scopes' {
        Mock Protect-SyncLists {}
        Initialize-SyncSharePoint @connection -LockDownReporting -AdminObjectIds @('44444444-4444-4444-8444-444444444444')
        Should -Invoke Connect-MgGraph -Times 1 -Exactly
        ($script:requestedScopes[0] -join ',') | Should -Be 'Sites.Selected,User.ReadBasic.All'
        Should -Invoke Invoke-SyncGraph -Times 0 -ParameterFilter { $Path -like '*/permissions*' }
        Should -Invoke Protect-SyncLists -Times 1 -Exactly
    }
    It 'fails visibly when selected-site access is missing, without provisioning or requesting broader scopes' {
        Mock Invoke-SyncGraph { throw 'Access denied: app selected-site grant is missing.' } -ParameterFilter {
            $Path -match '/sites/example\.sharepoint\.com:'
        }
        { Initialize-SyncSharePoint @connection -Apply } | Should -Throw '*selected-site grant is missing*'
        $script:writes | Should -Be 0
        Should -Invoke Connect-MgGraph -Times 1 -Exactly
        ($script:requestedScopes[0] -join ',') | Should -Be 'Sites.Selected'
        Should -Invoke Invoke-SyncGraph -Times 0 -ParameterFilter { $Method -in @('POST', 'PATCH', 'DELETE') }
    }
    It 'creates once and verifies existing lists on rerun' {
        $first = @(Initialize-SyncSharePoint @connection -Apply)
        $first | Should -Contain 'SYNC_DEFAULT_CONCURRENCY=3'
        $script:writes | Should -Be 3
        Initialize-SyncSharePoint @connection -Apply | Out-Null
        $script:writes | Should -Be 3
    }
    It 'provisions explicit prototype lists without PnP, ACL changes or extra scopes and is idempotent' {
        Mock Protect-SyncLists { throw 'Prototype must not run lockdown.' }
        Mock Connect-PnPOnline { throw 'Prototype must not connect to SharePoint CSOM.' }
        $output = @(Initialize-SyncSharePoint @connection -PrototypeReporting -Apply)
        $output | Should -Contain 'REPORTING_ACCESS_MODE=prototype'
        $output | Should -Contain 'REPORTING_SECURITY_ENABLED=false'
        $output | Should -Not -Contain 'REPORTING_ADMIN_GROUP_ID=20'
        $script:writes | Should -Be 3
        Initialize-SyncSharePoint @connection -PrototypeReporting -Apply | Out-Null
        $script:writes | Should -Be 3
        ($script:requestedScopes[0] -join ',') | Should -Be 'Sites.Selected'
        Should -Invoke Protect-SyncLists -Times 0
        Should -Invoke Connect-PnPOnline -Times 0
        Should -Invoke Invoke-SyncGraph -Times 0 -ParameterFilter { $Path -like '*/permissions*' }
    }
    It 'previews prototype provisioning without writes and rejects mixing access modes before connecting' {
        Initialize-SyncSharePoint @connection -PrototypeReporting
        $script:writes | Should -Be 0
        Should -Invoke Connect-MgGraph -Times 1 -Exactly
        { Initialize-SyncSharePoint @connection -PrototypeReporting -LockDownReporting -Apply } | Should -Throw '*not both*'
        Should -Invoke Connect-MgGraph -Times 1 -Exactly
    }
    It 'rejects incompatible lists and invalid concurrency' {
        $script:testLists.Add(@{ id = 'bad'; displayName = 'CopilotSpendingPolicies'; description = 'unrelated' })
        { Initialize-SyncSharePoint @connection } | Should -Throw '*not a compatible sync list*'
        { Initialize-SyncSharePoint @connection -DefaultConcurrency 8 -MaxConcurrency 2 } | Should -Throw '*cannot exceed*'
    }
    It 'adds only the explicit runtime site write grant and preserves it on rerun' {
        Initialize-SyncSharePoint @connection -RuntimeAppId '99999999-9999-4999-8999-999999999999' -GrantSiteAccess -Apply | Out-Null
        $script:testGrants.Count | Should -Be 1
        Initialize-SyncSharePoint @connection -RuntimeAppId '99999999-9999-4999-8999-999999999999' -GrantSiteAccess -Apply | Out-Null
        $script:testGrants.Count | Should -Be 1
        Should -Invoke Connect-MgGraph -Times 2 -Exactly
        ($script:requestedScopes[0] -join ',') | Should -Be 'Sites.FullControl.All'
        ($script:requestedScopes[1] -join ',') | Should -Be 'Sites.FullControl.All'
    }
}
Describe 'SharePoint setup request boundaries' {
    It 'rejects external pagination destinations before acquiring or sending a token' {
        Mock Invoke-MgGraphRequest { throw 'Request should not run.' }
        { Invoke-SyncGraph -Path 'https://untrusted.example/v1.0/sites' } | Should -Throw '*outside*'
        { Invoke-SyncGraph -Path 'https://graph.microsoft.com:8443/v1.0/sites' } | Should -Throw '*outside*'
        Should -Invoke Invoke-MgGraphRequest -Times 0
    }
    Describe 'Explicit administrator-only permission migration' {
        BeforeEach {
            $script:resetCount = 0
            $script:groupWrites = 0
            $script:pnpConnection = 'synthetic-connection'
            $script:destination = @{ siteId = 'site'; policyListId = 'policy'; balanceListId = 'balance'; jobListId = 'jobs' }
            Mock Import-Module {}
            Mock Invoke-SyncGraph { @{ id = '11111111-1111-4111-8111-111111111111'; userPrincipalName = 'admin@example.test' } }
            Mock Connect-PnPOnline { $script:pnpConnection }
            Mock Get-PnPContext { [pscustomobject]@{ Web = [pscustomobject]@{ CurrentUser = [pscustomobject]@{ Id = 101 } } } }
            Mock Get-PnPProperty { param($ClientObject, $Property); $ClientObject.$Property }
            Mock New-PnPUser { [pscustomobject]@{ Id = 101; LoginName = 'admin@example.test'; AadObjectId = @{ NameId = '11111111-1111-4111-8111-111111111111' } } }
            Mock Get-PnPGroup { [pscustomobject]@{
                Id = 20; Title = 'CopilotCostSync Administrators'; Description = 'CopilotCostSync/v1/administrators'
                Owner = @{ Id = 101 }; AllowMembersEditMembership = $false; AllowRequestToJoinLeave = $false; AutoAcceptRequestToJoinLeave = $false
            } }
            Mock Set-PnPGroup {
                param($AllowMembersEditMembership, $AllowRequestToJoinLeave, $AutoAcceptRequestToJoinLeave)
                $AllowMembersEditMembership | Should -BeFalse
                $AllowRequestToJoinLeave | Should -BeFalse
                $AutoAcceptRequestToJoinLeave | Should -BeFalse
                $script:groupWrites++
            }
            Mock Get-PnPGroupMember { [pscustomobject]@{ Id = 101; LoginName = 'admin@example.test' } }
            Mock Add-PnPGroupMember {}
            Mock Remove-PnPGroupMember {}
            Mock Get-PnPRoleDefinition { [pscustomobject]@{ Id = 1073741829; Name = 'Full Control'; RoleTypeKind = 'Administrator' } }
            Mock Get-PnPList {
                param($Identity)
                $key = switch ($Identity) { policy { 'policyListId' }; balance { 'balanceListId' }; jobs { 'jobListId' } }
                [pscustomobject]@{
                    Description = "CopilotCostSync/v1/$key"; BaseTemplate = 100; HasUniqueRoleAssignments = $true
                    RoleAssignments = @([pscustomobject]@{ Member = [pscustomobject]@{ Id = 20 }; RoleDefinitionBindings = @([pscustomobject]@{ Id = 1073741829 }) })
                }
            }
            Mock Set-PnPList {
                param($BreakRoleInheritance, $CopyRoleAssignments, $EnableFolderCreation)
                $BreakRoleInheritance | Should -BeTrue
                $CopyRoleAssignments | Should -BeFalse
                $EnableFolderCreation | Should -BeFalse
            }
            Mock Set-PnPListPermission {}
            Mock Get-PnPListItem {
                $item = [pscustomobject]@{ HasUniqueRoleAssignments = $false }
                $item | Add-Member -MemberType ScriptMethod -Name ResetRoleInheritance -Value { $script:resetCount++ }
                $item
            }
            Mock Invoke-PnPQuery {}
        }
        It 'previews revocation without connecting to SharePoint or changing access' {
            Protect-SyncLists @connection -Destination $script:destination -AdminObjectIds @('11111111-1111-4111-8111-111111111111')
            Should -Invoke Connect-PnPOnline -Times 0
            $script:resetCount | Should -Be 0
        }
        It 'requires explicit administrator IDs before provisioning writes' {
            { Initialize-SyncSharePoint @connection -LockDownReporting -Apply } | Should -Throw '*requires explicit AdminObjectIds*'
            { Protect-SyncLists @connection -Destination $script:destination -AdminObjectIds @('not-guid') -Apply } | Should -Throw '*non-zero GUIDs*'
            Should -Invoke Connect-PnPOnline -Times 0
        }
        It 'locks all three lists and resets every existing generation before declaring administrator-only' {
            $output = @(Protect-SyncLists @connection -Destination $script:destination -AdminObjectIds @('11111111-1111-4111-8111-111111111111') -Apply)
            $output | Should -Contain 'REPORTING_ADMIN_GROUP_ID=20'
            $output | Should -Not -Contain 'REPORTING_SECURITY_ENABLED=true'
            $script:resetCount | Should -Be 3
            $script:groupWrites | Should -Be 1
            Should -Invoke Set-PnPListPermission -Times 3
        }
        It 'fails closed when even one item retains unique access after reset' {
            Mock Get-PnPListItem {
                $item = [pscustomobject]@{ HasUniqueRoleAssignments = $true }
                $item | Add-Member -MemberType ScriptMethod -Name ResetRoleInheritance -Value {}
                $item
            }
            { Protect-SyncLists @connection -Destination $script:destination -AdminObjectIds @('11111111-1111-4111-8111-111111111111') -Apply } | Should -Throw '*revocation readback failed*'
        }
        It 'refuses a stale SharePoint administrator identity even when its UPN resolves' {
            Mock New-PnPUser { [pscustomobject]@{ Id = 101; LoginName = 'admin@example.test'; AadObjectId = @{ NameId = '99999999-9999-4999-8999-999999999999' } } }
            { Protect-SyncLists @connection -Destination $script:destination -AdminObjectIds @('11111111-1111-4111-8111-111111111111') -Apply } | Should -Throw '*current directory object ID*'
            Should -Invoke Set-PnPListPermission -Times 0
        }
    }
}
