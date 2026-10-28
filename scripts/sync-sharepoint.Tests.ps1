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
    function Connect-MgGraph { param($TenantId, $ClientId, $Scopes, $ContextScope, [switch]$NoWelcome, $ErrorAction) }
    function Get-MgContext { throw 'Unexpected context lookup.' }
    function Invoke-MgGraphRequest { param($Method, $Uri, $OutputType, $ErrorAction, $Body, $ContentType); throw 'Unexpected request.' }
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
        $script:writes = 0
        Mock Import-Module {}
        Mock Connect-MgGraph {}
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
    }
    It 'creates once and verifies existing lists on rerun' {
        $first = @(Initialize-SyncSharePoint @connection -Apply)
        $first | Should -Contain 'SYNC_DEFAULT_CONCURRENCY=3'
        $script:writes | Should -Be 3
        Initialize-SyncSharePoint @connection -Apply | Out-Null
        $script:writes | Should -Be 3
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
    }
}
Describe 'SharePoint setup request boundaries' {
    It 'rejects external pagination destinations before acquiring or sending a token' {
        Mock Invoke-MgGraphRequest { throw 'Request should not run.' }
        { Invoke-SyncGraph -Path 'https://untrusted.example/v1.0/sites' } | Should -Throw '*outside*'
        { Invoke-SyncGraph -Path 'https://graph.microsoft.com:8443/v1.0/sites' } | Should -Throw '*outside*'
        Should -Invoke Invoke-MgGraphRequest -Times 0
    }
}
