# Standalone consumption export

[Export-CopilotConsumption.ps1](../Export-CopilotConsumption.ps1) is a Windows
PowerShell script. It does not require a .NET project, build step, application
deployment, Key Vault, or manually managed refresh tokens.

It retains one JSON snapshot per spending policy and one combined reporting CSV
per run. Collection is read-only and independent of the portal.

> Cost Management Graph APIs are beta/preview and are not supported for production
> applications. Consumption is month-to-date in UTC, measured in the reported
> unit, not invoiced currency. This is not a historical accounting ledger.

## Prerequisites and consent

- Windows and PowerShell 7.4 or later (`pwsh.exe`).
- Microsoft.Graph.Authentication 2.37.0 or later, installed for the Windows account
  that will run the task. Pester is required only for development/tests.
- A tenant enabled for the preview and a public-client app registration configured
  for delegated device-code authentication. Use the organization's approved
  registration and consent process; this script does not create or alter consent.
- The user must have the endpoint's required runtime authorization and any required
  enterprise-application assignment. Consent alone is insufficient.

```powershell
Install-Module Microsoft.Graph.Authentication -MinimumVersion 2.37.0 -Scope CurrentUser
```

Grant these delegated read scopes:

| Scope | Purpose |
| --- | --- |
| `CopilotCostManagement-Policy.Read.All` | Policy inventory |
| `CopilotCostManagement-Assignment.Read.All` | Assigned groups |
| `CopilotCostManagement-UserData.Read.All` | User/service balances |
| `GroupMember.Read.All` | Direct and nested membership |
| `User.ReadBasic.All` | Basic user identity fields where available |

No write scopes or tenant-credit-balance permission are requested. The published
membership API mentions a narrower `GroupMember.ReadBasic.All`, but its availability
is not consistently reflected in the permissions reference; the script explicitly
uses the established `GroupMember.Read.All`, without substituting scopes silently.
Hidden membership requires additional authorization, including `Member.Read.Hidden`.
It is not requested automatically. Unreadable membership is an error, not an
empty group. Graph visibility constraints can limit discovery; the exporter
cannot prove that invisible users or groups do not exist.

## Initial sign-in and export

Run setup interactively under the Windows account that will execute the task:

```powershell
$connection = @{
    TenantId = '<tenant-guid>'
    ClientId = '<public-client-app-guid>'
    Account = 'reporting-user@example.com'
    OutputPath = 'C:\PrivateReports\CopilotConsumption'
}
.\Export-CopilotConsumption.ps1 -Mode Setup @connection
```

Follow the device-code instructions in the terminal. Graph PowerShell owns the
protected `CurrentUser` cache and refresh-token renewal. The script verifies the
tenant, client, expected user, delegated access, requested scopes and persistent
cache context. It does not inspect or copy tokens. Do not call `Disconnect-MgGraph`
after setup or daily runs: doing so clears the sign-in the next run needs.

Then test export in a **new PowerShell process**, using the same arguments:

```powershell
.\Export-CopilotConsumption.ps1 -Mode Export @connection
```

The output directory must be outside the repository, must not traverse junctions
or symbolic links, and should have restrictive NTFS permissions. The script uses
the existing directory's access controls; it does not grant access to other users.
The default is beneath the current Windows account's local application-data
directory. Do not place output in a shared/public web directory or commit it.

If Graph's shared cached identity belongs to another account, the script fails
before retrieving data. Prefer a dedicated Windows profile. If you deliberately
sign out with `Disconnect-MgGraph` to correct the identity, be aware this affects
other scripts using that profile's Graph sign-in, then run Setup again.

### Scheduled authentication safeguards

Graph PowerShell has no public silent-only connection switch. Export therefore
runs the same script in a supervised child PowerShell process:

- It uses `-UseDeviceCode`, so authentication cannot launch a browser.
- Only structured worker events are accepted on stdout. If the SDK prints a
  device-code challenge or other unexpected output, the supervisor stops its
  worker, discards that text and writes a reauthentication-required failure.
  Unexpected SDK output may conservatively produce the same failure category.
- Authentication and Graph requests have bounded watchdogs. The default initial
  authentication watchdog is 90 seconds; the HTTP timeout is 30 seconds, with a
  request watchdog allowing an additional 15 seconds. Local processing/retry
  phases have a 300-second inactivity limit.
- The same guard applies if a refresh needs interaction midway through a scan.
- The supervisor terminates only its own worker process tree and preserves the
  previous complete pointer. A killed run's finished policy files remain evidence,
  but a partially collected group might not yet have been checkpointed.

This is supervised cached authentication, **not** an SDK promise that interaction
will never be requested. A refresh token's nominal lifetime is not a guaranteed
unattended window; Conditional Access, revocation and directory changes can require
reauthentication. Do not disable tenant protections to accommodate this script.

## Collection and coverage

All collection routes are paginated; continuation URLs are validated and repeated
links are rejected.

1. `GET /beta/copilot/costManagement/spendingPolicies`
2. For each selected-groups policy:
   `GET /beta/copilot/costManagement/spendingPolicies/{id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups`
3. For each assigned group: `GET /v1.0/groups/{id}/transitiveMembers`.
   Graph flattens nested membership; the script selects user objects locally.
4. For every user occurrence:
   `GET /beta/copilot/costManagement/userBalances/{id}/serviceBalances`.
5. Compare each returned `spendingPolicyId` with the discovery policy ID.

The script intentionally retains users repeated across groups or policies,
including repeated balance requests. A mismatch is flagged, not reassigned.
Two matching observations can still be duplicates: **do not sum rows as unique
policy consumption**. There are no aggregate consumption totals.

Tenant-wide policies and unknown policy types receive JSON inventory/exclusion
records. Their consumption is **not collected**, not zero. A complete run means
complete for its declared group-scoped collection, not the whole tenant.

GET requests are sequential. Transient 429/503/504 failures retry up to three
times, honoring Retry-After; waits longer than 60 seconds are surfaced rather than
retrying prematurely. Graph's own HTTP handling may also retry, subject to the
supervisor watchdog. Failed pages preserve earlier returned rows, mark coverage
partial and prevent complete publication.

## Output contract

```text
<output>\
  <tenant-id>\
    <UTC-collection-month>\
      <unique-run-id>\
        policy-<id>.json
        consumption.csv
        manifest.json
    latest-complete.json
    latest-status.json
    export.lock
```

No completed snapshot is overwritten, including on manual reruns. Working files
are checkpointed until their manifest is finalized. There is no automatic deletion:
set an organizational retention policy before accumulating long-lived user data.

- **Policy JSON:** full policy object, group/member outcomes, normalized
  observations, original service-balance objects, exclusions and explicit errors.
  This is the canonical evidence.
- **Combined CSV:** the same user/service observations across collected policies.
  No additional Graph requests are made to create it.
- **Manifest:** schema/run IDs, UTC times, scope, complete/partial/failed state,
  exclusions, occurrence counts, JSON paths, CSV schema/count, errors, per-user
  observed-policy index and policy-change findings.
- **Latest complete:** relative paths identifying the newest complete run's
  manifest and CSV. Publication requires both output formats and a successful
  supervised worker exit.
- **Latest status:** the latest attempt, even when authentication or collection
  failed. Always check this as well as the older complete pointer to avoid
  presenting stale data as a successful refresh.
- **Lock:** an OS-held exclusive file handle prevents overlapping exports for one
  tenant/output destination. The empty file may remain after exit; its presence
  alone does not mean the task is running.

CSV schema version 1 has these columns:

```text
schemaVersion,runId,tenantId,collectionStartedAtUtc,observedAtUtc,
discoveryPolicyId,discoveryPolicyName,sourceAssignedGroupId,userId,
userDisplayName,userPrincipalName,serviceId,returnedSpendingPolicyId,
policyMatch,validationStatus,consumedQuantity,unit,remainingQuantityKind,
remainingQuantityValue,isConsumptionAllowed,accessDecisionReason,
sourceLastModifiedDateTime
```

`policyMatch` is `matched`, `mismatch` or `unknown`; `validationStatus` is `valid`
or `invalid`. A genuine zero quantity is `0`; unavailable/invalid quantities are
empty, not zero. Remaining quantity kind is `limited`, `unlimited`, `missing` or
`invalid`; unlimited has no invented numeric value. Read configured limits and
other nested quantities from the original JSON.

CSV uses UTF-8 with BOM, quoted fields, invariant decimal notation and UTC ISO
timestamps. Quotes, commas, newlines and Unicode names round-trip. Untrusted text
beginning with a spreadsheet formula indicator (`=`, `+`, `-`, `@`, including after
whitespace/control characters) receives an apostrophe prefix. Original text
remains unchanged in JSON. Numeric quantities do not receive this text escaping.
A CSV with no service observations still has its headers; failed/no-balance users,
empty policies and excluded policies remain described in JSON/manifest rather
than being represented as fabricated service rows.

Select **one run** for a current report. Each reading is month-to-date, so summing
daily snapshots double-counts consumption. Source freshness can lag collection;
collection is not an atomic tenant snapshot. A scan crossing a UTC month boundary
is partial and is not promoted.

Policy-change findings compare the last usable observation per user/service
across retained run manifests. Increasing source freshness and a different returned
policy support `observedChange`; conflicting or insufficient freshness produces
`ambiguous`. New or absent users do not prove a move. Users leaving all discovered
groups stop being scanned; moves away and back between runs can be missed. No
exact cutover time, historical recovery or consumption-delta attribution is claimed.

## Scheduling, health and recovery

Configure a daily Windows Task Scheduler task:

- **Program:** the full path to `pwsh.exe`.
- **Arguments:** `-NoProfile -NonInteractive -File "C:\Path\Export-CopilotConsumption.ps1" -Mode Export -TenantId "<guid>" -ClientId "<guid>" -Account "reporting-user@example.com" -OutputPath "C:\PrivateReports\CopilotConsumption"`.
- Run under the same Windows identity/profile used for Setup.
- Select **Do not start a new instance**. The script also has its own destination
  lock. Ensure the account/profile and encrypted cache are available for the
  selected task logon type; do not assume an interactive run proves this.

Exit code `0` means a complete export; `1` means partial or failed. Inspect
`latest-status.json`, then the referenced manifest and its errors. Failures before
an output lock can be acquired (including invalid/unwritable destinations) are
reported on the error stream and by exit code; an active export's status is not
overwritten by a competing attempt. External alert delivery is not included.

For `ReauthenticationRequired`, run Setup interactively again and retry Export.
For an account mismatch, deliberately correct the cached identity. For HTTP 403,
review consent, user authorization, group visibility and preview access. Longer
Retry-After failures should be retried later, not bypassed.

**Operator acceptance is required:** test a new-process export, the actual task
logon context, token renewal when needed, and recovery after reauthentication is
required. Automated tests use synthetic Graph data and do not prove live Entra
renewal or tenant-specific scheduling behavior.

## Development validation

```powershell
Install-PSResource Pester -Version 5.7.1 -Scope CurrentUser -TrustRepository
Import-Module Pester -RequiredVersion 5.7.1
Invoke-Pester -Path .\scripts\consumption-export.Tests.ps1 -Output Detailed -CI
```

Tests use mocked requests and a synthetic Graph module in a private test directory.
They make no tenant requests or consent changes.

## Microsoft references

- [Cost Management API overview](https://learn.microsoft.com/microsoft-365/copilot/extensibility/api/admin-settings/cost-management/overview)
- [Transitive membership](https://learn.microsoft.com/graph/api/group-list-transitivemembers?view=graph-rest-1.0)
- [Service-balance contract](https://learn.microsoft.com/microsoft-365/copilot/extensibility/api/admin-settings/cost-management/resources/costmanagementuserservicebalance)
- [Graph PowerShell authentication](https://learn.microsoft.com/powershell/microsoftgraph/authentication-commands)
- [Refresh tokens](https://learn.microsoft.com/entra/identity-platform/refresh-tokens)
