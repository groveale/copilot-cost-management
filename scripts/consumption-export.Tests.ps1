#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.7.1' }

BeforeAll {
    . (Join-Path $PSScriptRoot '..\Export-CopilotConsumption.ps1')
    $script:policyA = '33333333-3333-3333-3333-333333333333'
    $script:policyB = '77777777-7777-7777-7777-777777777777'
    $script:groupA = '55555555-5555-5555-5555-555555555555'
    $script:groupB = '66666666-6666-6666-6666-666666666666'
    $script:userA = '44444444-4444-4444-4444-444444444444'
    $script:base = '/beta/copilot/costManagement'

    function Connect-MgGraph {
        [CmdletBinding()]
        param($TenantId, $ClientId, $Scopes, $ContextScope, [switch]$UseDeviceCode, $ClientTimeout, [switch]$NoWelcome)
        throw 'Unexpected unmocked authentication.'
    }
    function Get-MgContext { throw 'Unexpected unmocked context access.' }
    function Invoke-MgGraphRequest {
        [CmdletBinding()]
        param($Method, $Uri, $OutputType)
        throw 'Unexpected unmocked Graph request.'
    }

    function New-TestConfiguration {
        @{
            tenantId = '11111111-1111-1111-1111-111111111111'
            clientId = '22222222-2222-2222-2222-222222222222'
            account = 'reporter@example.test'; outputPath = (Join-Path $TestDrive ([guid]::NewGuid().ToString('N')))
            requestTimeoutSeconds = 5; authenticationTimeoutSeconds = 10; maxRetryCount = 0
        }
    }

    function New-TestPages {
        $name = "Ren$([char]0xe9)e, `"Example`"`nSecond line"
        @{
            "$script:base/spendingPolicies" = @{ value = @(
                @{ id = $script:policyA; displayName = 'Policy A'; '@odata.type' = '#microsoft.graph.selectedGroupsSpendingPolicy' }
                @{ id = $script:policyB; displayName = 'Tenant-wide'; '@odata.type' = '#microsoft.graph.allUsersSpendingPolicy' }
            ) }
            "$script:base/spendingPolicies/$script:policyA/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups" = @{
                value = @(@{ id = $script:groupA }, @{ id = $script:groupB })
            }
            "/v1.0/groups/$script:groupA/transitiveMembers" = @{ value = @(
                @{ id = $script:userA; displayName = $name; userPrincipalName = 'user@example.test'; '@odata.type' = '#microsoft.graph.user' }
                @{ id = '88888888-8888-8888-8888-888888888888'; '@odata.type' = '#microsoft.graph.group' }
            ) }
            "/v1.0/groups/$script:groupB/transitiveMembers" = @{ value = @(
                @{ id = $script:userA; displayName = $name; '@odata.type' = '#microsoft.graph.user' }
            ) }
            "$script:base/userBalances/$script:userA/serviceBalances" = @{ value = @(
                @{ serviceId = 'cowork'; spendingPolicyId = $script:policyA; consumedQuantity = 0; unit = 'copilotCredit'
                    remainingQuantity = @{ '@odata.type' = '#microsoft.graph.unlimitedCostManagementQuantity'; unit = 'copilotCredit' }
                    dataStatus = @{ lastModifiedDateTime = '2026-10-20T01:00:00Z' }; isConsumptionAllowed = $true }
                @{ serviceId = 'workIQ'; spendingPolicyId = $script:policyB; consumedQuantity = 12.5; unit = 'copilotCredit'
                    remainingQuantity = @{ '@odata.type' = '#microsoft.graph.limitedCostManagementQuantity'; value = 87.5; unit = 'copilotCredit' }
                    dataStatus = @{ lastModifiedDateTime = '2026-10-20T01:00:00Z' }; isConsumptionAllowed = $false; accessDecisionReason = 'notEntitled' }
            ) }
        }
    }

    function Invoke-TestCollection {
        param([hashtable]$Config, [string]$StartedAt)
        $id = [DateTimeOffset]::UtcNow.ToString('yyyy-MM-ddTHHmmssfffffffZ') + '-' + [guid]::NewGuid().ToString('N')
        $tenantDirectory = Join-Path $Config.outputPath $Config.tenantId
        $directory = Join-Path (Join-Path $tenantDirectory $id.Substring(0, 7)) $id
        [void][IO.Directory]::CreateDirectory($directory)
        $manifest = New-ConsumptionManifest $Config $id
        if ($StartedAt) { $manifest.startedAtUtc = $StartedAt }
        $manifest = Invoke-ConsumptionCollectionRun $Config $directory $manifest
        @{ manifest = $manifest; directory = $directory; tenantDirectory = $tenantDirectory }
    }

    function Start-TestWorker {
        param([string]$Code)
        $process = [Diagnostics.Process]::new()
        $process.StartInfo.FileName = (Get-Process -Id $PID).Path
        $process.StartInfo.UseShellExecute = $false
        $process.StartInfo.RedirectStandardOutput = $true
        $process.StartInfo.RedirectStandardError = $true
        foreach ($argument in @('-NoProfile', '-NonInteractive', '-EncodedCommand', [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Code)))) {
            $process.StartInfo.ArgumentList.Add($argument)
        }
        [void]$process.Start()
        $process
    }
}

Describe 'Consumption collection and immutable output' {
    BeforeEach {
        $script:pages = New-TestPages
        $script:config = New-TestConfiguration
        Mock Invoke-ConsumptionGraphGet {
            param($Uri)
            $path = ([uri]$Uri).PathAndQuery
            if (-not $script:pages.ContainsKey($path)) { throw "Unexpected API call $path" }
            $script:pages[$path]
        }
    }

    It 'expands membership, retains duplicate observations and validates reported policy per service' {
        $run = Invoke-TestCollection $script:config
        $run.manifest.status | Should -Be 'complete'
        $run.manifest.csvRowCount | Should -Be 4
        $run.manifest.discoveredUserOccurrences | Should -Be 2
        $run.manifest.policyMismatchRows | Should -Be 2
        $run.manifest.exclusions[0].reason | Should -Be 'TenantWidePolicyNotCollected'
        Should -Invoke Invoke-ConsumptionGraphGet -Times 2 -Exactly -ParameterFilter { $Uri -like '*/serviceBalances' }
        Should -Invoke Invoke-ConsumptionGraphGet -Times 0 -ParameterFilter { $Uri -like "*/$script:policyB/*" }
        $policy = Get-Content (Join-Path $run.directory "policy-$script:policyA.json") -Raw | ConvertFrom-Json -AsHashtable
        $policy.groups[0].skippedNonUserObjects | Should -Be 1
        $policy.observations.Count | Should -Be 4
        $policy.observations[1].returnedSpendingPolicyId | Should -Be $script:policyB
        $policy.observations[1].discoveryPolicyId | Should -Be $script:policyA
        $csv = @(Import-Csv (Join-Path $run.directory 'consumption.csv'))
        $csv.Count | Should -Be 4
        $csv[0].consumedQuantity | Should -Be '0'
        $csv[0].remainingQuantityKind | Should -Be 'unlimited'
        $csv[0].remainingQuantityValue | Should -BeNullOrEmpty
        $csv[0].userDisplayName | Should -Be $policy.observations[0].userDisplayName
        $csv[1].consumedQuantity | Should -Be '12.5'
        $csv[1].policyMatch | Should -Be 'mismatch'
        @($csv[0].PSObject.Properties.Name) | Should -Be $script:ConsumptionCsvColumns
        $bytes = [IO.File]::ReadAllBytes((Join-Path $run.directory 'consumption.csv'))
        $bytes[0..2] | Should -Be @(239, 187, 191)
        $run.manifest.policyObservations.Count | Should -Be 2
        $run.manifest.policyChanges.Count | Should -Be 0
    }

    It 'does not advance latest complete until both formats have been published' {
        $first = Invoke-TestCollection $script:config
        Test-Path (Join-Path $first.tenantDirectory 'latest-complete.json') | Should -BeFalse
        Publish-ConsumptionRun $first.tenantDirectory $first.directory $first.manifest
        $pointer = Get-Content (Join-Path $first.tenantDirectory 'latest-complete.json') -Raw
        $second = Invoke-TestCollection $script:config
        Remove-Item -LiteralPath (Join-Path $second.directory 'consumption.csv')
        { Publish-ConsumptionRun $second.tenantDirectory $second.directory $second.manifest } | Should -Throw '*complete JSON/CSV*'
        Get-Content (Join-Path $first.tenantDirectory 'latest-complete.json') -Raw | Should -BeExactly $pointer
        Test-Path (Join-Path $first.directory 'consumption.csv') | Should -BeTrue
        $first.directory | Should -Not -Be $second.directory
    }

    It 'preserves successful rows from a failed later page and marks the run partial' {
        $path = "$script:base/userBalances/$script:userA/serviceBalances"
        $script:pages[$path]['@odata.nextLink'] = "https://graph.microsoft.com$path`?page=2"
        Mock Invoke-ConsumptionGraphGet { Stop-ConsumptionOperation 'AuthorizationFailed' 'HTTP 403 on continuation' } -ParameterFilter { $Uri -like '*page=2' }
        $run = Invoke-TestCollection $script:config
        $run.manifest.status | Should -Be 'partial'
        $run.manifest.csvRowCount | Should -Be 4
        $run.manifest.successfulUserOccurrences | Should -Be 0
        $run.manifest.errors[0].failure.category | Should -Be 'AuthorizationFailed'
    }

    It 'follows continuation links for every collection type' {
        foreach ($path in @($script:pages.Keys)) {
            $items = $script:pages[$path].value
            $script:pages[$path] = @{ value = @(); '@odata.nextLink' = "https://graph.microsoft.com$path`?page=2" }
            $script:pages["$path`?page=2"] = @{ value = $items }
        }
        $run = Invoke-TestCollection $script:config
        $run.manifest.status | Should -Be 'complete'
        $run.manifest.csvRowCount | Should -Be 4
        Should -Invoke Invoke-ConsumptionGraphGet -Times 6 -Exactly -ParameterFilter { $Uri -like '*page=2' }
    }

    It 'rejects foreign and repeated continuation links without leaking a request' {
        $path = "$script:base/spendingPolicies"
        $script:pages[$path]['@odata.nextLink'] = 'https://example.test/collect'
        $collection = Get-ConsumptionCollection $path $script:config
        $collection.complete | Should -BeFalse
        $collection.error.category | Should -Be 'InvalidContinuation'
        Should -Invoke Invoke-ConsumptionGraphGet -Times 1 -Exactly
        $script:pages[$path]['@odata.nextLink'] = "https://graph.microsoft.com$path"
        (Get-ConsumptionCollection $path $script:config).error.category | Should -Be 'InvalidContinuation'
    }

    It 'records empty service responses without inventing zero-consumption rows' {
        $script:pages["$script:base/userBalances/$script:userA/serviceBalances"].value = @()
        $run = Invoke-TestCollection $script:config
        $run.manifest.status | Should -Be 'complete'
        $run.manifest.csvRowCount | Should -Be 0
        $csv = Get-Content (Join-Path $run.directory 'consumption.csv')
        @($csv).Count | Should -Be 1
        $csv | Should -Match '"schemaVersion","runId"'
    }

    It 'keeps malformed quantities and missing policy IDs explicit rather than zero' {
        $row = $script:pages["$script:base/userBalances/$script:userA/serviceBalances"].value[0]
        $row.consumedQuantity = '12'
        $row.Remove('spendingPolicyId')
        $run = Invoke-TestCollection $script:config
        $run.manifest.status | Should -Be 'partial'
        $csv = @(Import-Csv (Join-Path $run.directory 'consumption.csv'))
        $csv[0].consumedQuantity | Should -BeNullOrEmpty
        $csv[0].policyMatch | Should -Be 'unknown'
        $csv[0].validationStatus | Should -Be 'invalid'
    }

    It 'treats inaccessible membership as an error, not an empty group' {
        Mock Invoke-ConsumptionGraphGet { Stop-ConsumptionOperation 'AuthorizationFailed' 'Hidden membership not authorized.' } -ParameterFilter { $Uri -like '*/transitiveMembers' }
        $run = Invoke-TestCollection $script:config
        $run.manifest.status | Should -Be 'partial'
        $run.manifest.errors.Count | Should -Be 2
        $run.manifest.discoveredUserOccurrences | Should -Be 0
    }

    It 'does not publish a mixed UTC month snapshot' {
        $run = Invoke-TestCollection $script:config '2001-01-31T23:59:00Z'
        $run.manifest.status | Should -Be 'partial'
        $run.manifest.errors[-1].failure.category | Should -Be 'MixedReportingPeriod'
    }

    It 'does not convert a CSV write failure into a complete run' {
        Mock Write-ConsumptionCsv { throw 'CSV write failed' }
        { Invoke-TestCollection $script:config } | Should -Throw '*CSV write failed*'
        @(Get-ChildItem $script:config.outputPath -Filter 'latest-complete.json' -Recurse).Count | Should -Be 0
    }

    It 'retains originals in JSON when neutralizing spreadsheet formula text in CSV' {
        $script:pages["/v1.0/groups/$script:groupA/transitiveMembers"].value[0].displayName = '=UNTRUSTED()'
        $run = Invoke-TestCollection $script:config
        $policy = Get-Content (Join-Path $run.directory "policy-$script:policyA.json") -Raw | ConvertFrom-Json -AsHashtable
        $csv = @(Import-Csv (Join-Path $run.directory 'consumption.csv'))
        $policy.observations[0].userDisplayName | Should -Be '=UNTRUSTED()'
        $csv[0].userDisplayName | Should -Be "'=UNTRUSTED()"
    }

    It 'reads last successful observations across month directories and rejects corrupt history' {
        $first = Invoke-TestCollection $script:config
        $historicDirectory = Join-Path (Join-Path $first.tenantDirectory '2001-01') 'historic'
        [void][IO.Directory]::CreateDirectory($historicDirectory)
        Move-Item -LiteralPath (Join-Path $first.directory 'manifest.json') -Destination (Join-Path $historicDirectory 'manifest.json')
        $index = Get-ConsumptionPreviousPolicies $first.tenantDirectory 'new'
        $index.Count | Should -Be 2
        $index["$script:userA|cowork"].sourceTimes[0] | Should -BeOfType [string]
        [IO.File]::WriteAllText((Join-Path $historicDirectory 'manifest.json'), '{broken')
        { Get-ConsumptionPreviousPolicies $first.tenantDirectory 'new' } | Should -Throw
    }
}

Describe 'Public Graph PowerShell authentication and retry boundaries' {
    BeforeEach {
        $script:config = New-TestConfiguration
        Mock Import-Module {} -ParameterFilter { $Name -eq 'Microsoft.Graph.Authentication' }
        Mock Connect-MgGraph {}
        Mock Get-MgContext {
            @{ AuthType = 'Delegated'; ContextScope = 'CurrentUser'; TenantId = $script:config.tenantId
                ClientId = $script:config.clientId; Account = $script:config.account; Scopes = $script:ConsumptionScopes }
        }
    }

    It 'uses persistent device-code authentication and never requests a browser' {
        Connect-ConsumptionGraph $script:config
        Should -Invoke Connect-MgGraph -Times 1 -Exactly -ParameterFilter { $UseDeviceCode -and $ContextScope -eq 'CurrentUser' -and $NoWelcome }
    }

    It 'rejects a different account and an in-memory fallback' {
        Mock Get-MgContext {
            @{ AuthType='Delegated'; ContextScope='CurrentUser'; TenantId=$script:config.tenantId
                ClientId=$script:config.clientId; Account='other@example.test'; Scopes=$script:ConsumptionScopes }
        }
        { Connect-ConsumptionGraph $script:config } | Should -Throw '*configured tenant, client and account*'
        Mock Get-MgContext {
            @{ AuthType='Delegated'; ContextScope='Process'; TenantId=$script:config.tenantId
                ClientId=$script:config.clientId; Account=$script:config.account; Scopes=$script:ConsumptionScopes }
        }
        { Connect-ConsumptionGraph $script:config } | Should -Throw '*persistent delegated connection*'
    }

    It 'honors Retry-After and stops after the configured retry count' {
        $script:config.maxRetryCount = 1
        $script:response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]::TooManyRequests)
        $script:response.Headers.RetryAfter = [Net.Http.Headers.RetryConditionHeaderValue]::new([TimeSpan]::FromSeconds(7))
        Mock Start-Sleep {}
        Mock Invoke-MgGraphRequest { throw [Microsoft.PowerShell.Commands.HttpResponseException]::new('Throttled', $script:response) }
        { Invoke-ConsumptionGraphGet 'https://graph.microsoft.com/beta/copilot/costManagement/spendingPolicies' $script:config } | Should -Throw '*HTTP 429*'
        Should -Invoke Start-Sleep -Times 1 -Exactly -ParameterFilter { $Seconds -eq 7 }
        Should -Invoke Invoke-MgGraphRequest -Times 2 -Exactly
        $script:response.Dispose()
    }

    It 'fails explicitly for an excessive Retry-After instead of retrying too soon' {
        $script:config.maxRetryCount = 1
        $script:response = [Net.Http.HttpResponseMessage]::new([Net.HttpStatusCode]::TooManyRequests)
        $script:response.Headers.RetryAfter = [Net.Http.Headers.RetryConditionHeaderValue]::new([TimeSpan]::FromSeconds(120))
        Mock Start-Sleep {}
        Mock Invoke-MgGraphRequest { throw [Microsoft.PowerShell.Commands.HttpResponseException]::new('Throttled', $script:response) }
        { Invoke-ConsumptionGraphGet 'https://graph.microsoft.com/beta/copilot/costManagement/spendingPolicies' $script:config } | Should -Throw '*120 seconds*'
        Should -Invoke Start-Sleep -Times 0
        $script:response.Dispose()
    }
}

Describe 'CSV projection and observed policy changes' {
    It 'uses invariant numbers and spreadsheet-safe text without changing raw observations' {
        $culture = [Globalization.CultureInfo]::CurrentCulture
        try {
            [Globalization.CultureInfo]::CurrentCulture = 'de-DE'
            ConvertTo-ConsumptionCsvCell 12.5 | Should -Be '"12.5"'
            ConvertTo-ConsumptionCsvCell "`t=SUM(1,2)" | Should -Be ('"' + "'`t=SUM(1,2)" + '"')
            ConvertTo-ConsumptionCsvCell 'a"b,c' | Should -Be '"a""b,c"'
            ConvertTo-ConsumptionCsvCell $null | Should -Be '""'
        } finally { [Globalization.CultureInfo]::CurrentCulture = $culture }
    }

    It 'does not accept negative, nonfinite or string consumption values' {
        foreach ($value in @(-1, [double]::NaN, [double]::PositiveInfinity, '1', $null)) {
            Test-ConsumptionNumber $value | Should -BeFalse
        }
        Test-ConsumptionNumber 0 | Should -BeTrue
    }

    It 'normalizes JSON date objects without culture-dependent string conversions' {
        $row = New-TestPages
        $balance = $row["$script:base/userBalances/$script:userA/serviceBalances"].value[0] |
            ConvertTo-Json -Depth 10 | ConvertFrom-Json -AsHashtable
        $manifest = New-ConsumptionManifest (New-TestConfiguration) 'date-roundtrip'
        $observation = New-ConsumptionObservation $balance @{id=$script:userA} $script:groupA @{id=$script:policyA} $manifest
        $observation.validationStatus | Should -Be 'valid'
        $observation.sourceLastModifiedDateTime | Should -Be '2026-10-20T01:00:00.0000000+00:00'
    }

    It 'reports one observed policy change for repeated rows and handles uncertain freshness' {
        $previous = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
        $previous["$script:userA|cowork"] = @{
            userId = $script:userA; serviceId = 'cowork'; policyIds = @($script:policyA)
            observedAtUtc = '2026-10-20T02:00:00Z'; sourceTimes = @('2026-10-20T01:00:00Z'); missingSourceTime = $false
        }
        $row = @{ userId = $script:userA; serviceId = 'cowork'; returnedSpendingPolicyId = $script:policyB
            validationStatus = 'valid'; observedAtUtc = '2026-10-21T02:00:00Z'; sourceLastModifiedDateTime = '2026-10-21T01:00:00Z' }
        $index = @(Get-ConsumptionPolicyIndex @($row, $row))
        $findings = @(Compare-ConsumptionPolicies $index $previous)
        $findings.Count | Should -Be 1
        $findings[0].state | Should -Be 'observedChange'
        $row.sourceLastModifiedDateTime = $null
        @(Compare-ConsumptionPolicies @(Get-ConsumptionPolicyIndex @($row)) $previous)[0].state | Should -Be 'ambiguous'
        @(Compare-ConsumptionPolicies @() $previous).Count | Should -Be 0
    }

    It 'flags conflicting repeated observations instead of inventing a cutover' {
        $rows = @(
            @{userId=$script:userA;serviceId='cowork';returnedSpendingPolicyId=$script:policyA;validationStatus='valid';observedAtUtc='2026-10-20T01:00:00Z'}
            @{userId=$script:userA;serviceId='cowork';returnedSpendingPolicyId=$script:policyB;validationStatus='valid';observedAtUtc='2026-10-20T01:00:01Z'}
        )
        $index = @(Get-ConsumptionPolicyIndex $rows)
        $previous = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
        @(Compare-ConsumptionPolicies $index $previous)[0].reason | Should -Be 'ConflictingPolicyObservations'
    }
}

Describe 'Worker supervision and configuration' {
    It 'accepts structured results from a separate PowerShell process' {
        $process = Start-TestWorker '[Console]::WriteLine(''COPILOT_EXPORT:{"type":"result","status":"complete"}'')'
        try { (Wait-ConsumptionWorker $process (New-TestConfiguration)).status | Should -Be 'complete' }
        finally { $process.Dispose() }
    }

    Describe 'Separate-process exporter with a synthetic Graph module' {
        BeforeAll {
            $script:originalModulePath = $env:PSModulePath
            $script:moduleRoot = Join-Path $TestDrive 'modules'
            $moduleDirectory = Join-Path $script:moduleRoot 'Microsoft.Graph.Authentication\2.37.0'
            [void][IO.Directory]::CreateDirectory($moduleDirectory)
            $syntheticModule = {
    function Connect-MgGraph {
        [CmdletBinding()]
        param($TenantId, $ClientId, $Scopes, $ContextScope, [switch]$UseDeviceCode, $ClientTimeout, [switch]$NoWelcome)
        if ($env:COPILOT_TEST_PROMPT -eq '1') {
            [Console]::WriteLine('Device sign-in requested: DO-NOT-LOG-THIS-CODE')
            Start-Sleep -Seconds 30
        }
        $script:context = @{ AuthType='Delegated'; ContextScope=$ContextScope; TenantId=$TenantId
            ClientId=$ClientId; Scopes=$Scopes; Account='reporter@example.test' }
    }
    function Get-MgContext { $script:context }
    function Invoke-MgGraphRequest {
        [CmdletBinding()]
        param($Method, $Uri, $OutputType)
        $pages = Get-Content -LiteralPath $env:COPILOT_TEST_PAGES -Raw | ConvertFrom-Json -AsHashtable
        $path = ([uri]$Uri).PathAndQuery
        if (-not $pages.ContainsKey($path)) { throw "Unexpected synthetic API route: $path" }
        $pages[$path]
    }
    Export-ModuleMember -Function Connect-MgGraph, Get-MgContext, Invoke-MgGraphRequest
            }
            $syntheticModule.ToString() | Set-Content -LiteralPath (Join-Path $moduleDirectory 'Microsoft.Graph.Authentication.psm1')
            New-ModuleManifest -Path (Join-Path $moduleDirectory 'Microsoft.Graph.Authentication.psd1') `
                -RootModule 'Microsoft.Graph.Authentication.psm1' -ModuleVersion '2.37.0' `
                -FunctionsToExport @('Connect-MgGraph', 'Get-MgContext', 'Invoke-MgGraphRequest')
        }
        BeforeEach {
            $env:PSModulePath = $script:moduleRoot + [IO.Path]::PathSeparator + $script:originalModulePath
            $env:COPILOT_TEST_PAGES = Join-Path $TestDrive 'synthetic-pages.json'
            $env:COPILOT_TEST_PROMPT = '0'
            New-TestPages | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $env:COPILOT_TEST_PAGES
            $script:config = New-TestConfiguration
        }
        AfterEach {
            $env:PSModulePath = $script:originalModulePath
            Remove-Item Env:COPILOT_TEST_PAGES, Env:COPILOT_TEST_PROMPT -ErrorAction SilentlyContinue
        }

        It 'writes both formats, a manifest and operational pointers through the real worker process' {
            $result = Invoke-ConsumptionExport $script:config
            if ($result.status -ne 'complete') {
                $failedManifest = Get-Content -LiteralPath $result.manifestPath -Raw | ConvertFrom-Json -AsHashtable
                throw (ConvertTo-Json -InputObject $failedManifest.errors -Depth 10 -Compress)
            }
            $result.status | Should -Be 'complete'
            $tenantDirectory = Join-Path $script:config.outputPath $script:config.tenantId
            $pointer = Get-Content (Join-Path $tenantDirectory 'latest-complete.json') -Raw | ConvertFrom-Json -AsHashtable
            @(Import-Csv (Join-Path $tenantDirectory $pointer.csvPath)).Count | Should -Be 4
            $manifest = Get-Content -LiteralPath $result.manifestPath -Raw | ConvertFrom-Json -AsHashtable
            $manifest.policyFiles.Count | Should -Be 2
            $manifest.csvRowCount | Should -Be 4
            (Get-Content (Join-Path $tenantDirectory 'latest-status.json') -Raw | ConvertFrom-Json).status | Should -Be 'complete'
        }

        It 'records reauthentication without exposing the device code or replacing the previous complete run' {
            $complete = Invoke-ConsumptionExport $script:config
            $tenantDirectory = Join-Path $script:config.outputPath $script:config.tenantId
            $pointer = Get-Content (Join-Path $tenantDirectory 'latest-complete.json') -Raw
            $env:COPILOT_TEST_PROMPT = '1'
            $failed = Invoke-ConsumptionExport $script:config
            $failed.status | Should -Be 'failed'
            $failed.failure.category | Should -Be 'ReauthenticationRequired'
            Get-Content (Join-Path $tenantDirectory 'latest-complete.json') -Raw | Should -BeExactly $pointer
            $manifest = Get-Content -LiteralPath $failed.manifestPath -Raw
            $manifest | Should -Not -Match 'DO-NOT-LOG-THIS-CODE'
            (Get-Content -LiteralPath $complete.manifestPath -Raw | ConvertFrom-Json).status | Should -Be 'complete'
        }
    }

    It 'cancels a device-code prompt immediately without recording its code' {
        $process = Start-TestWorker '[Console]::WriteLine("To sign in use SECRET-CODE"); Start-Sleep -Seconds 30'
        try {
            { Wait-ConsumptionWorker $process (New-TestConfiguration) } | Should -Throw '*without displaying the sign-in code*'
            $process.HasExited | Should -BeTrue
        } finally { $process.Dispose() }
    }

    It 'bounds stalled authentication and terminates only its worker' {
        $process = Start-TestWorker 'Start-Sleep -Seconds 30'
        $config = New-TestConfiguration
        $config.authenticationTimeoutSeconds = 0.1
        try {
            { Wait-ConsumptionWorker $process $config } | Should -Throw '*timed out during authentication*'
            $process.HasExited | Should -BeTrue
        } finally { $process.Dispose() }
    }

    It 'rejects outputs inside the repository and overlapping exports' {
        $config = New-TestConfiguration
        $config.outputPath = $script:ConsumptionRepositoryPath
        { Assert-ConsumptionConfiguration $config } | Should -Throw '*outside the repository*'
        $config = New-TestConfiguration
        $directory = Join-Path $config.outputPath $config.tenantId
        [void][IO.Directory]::CreateDirectory($directory)
        $lock = [IO.File]::Open((Join-Path $directory 'export.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
        try { { Invoke-ConsumptionExport $config } | Should -Throw '*export lock*' }
        finally { $lock.Dispose() }
    }
}
