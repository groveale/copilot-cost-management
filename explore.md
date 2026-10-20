

```powershell
Connect-MgGraph -TenantId d8464949-bb1d-476f-bbe8-a46293833f6d -ClientId e4b3e7ed-9827-4b77-bd8a-dc50e4447528

$baseUri = 'https://graph.microsoft.com/beta/copilot/costManagement'
```

## User balance

Looks at everything

```powershell
$userBalance = Invoke-MgGraphRequest -Method GET -Uri "$baseUri/userBalances/fa936341-b3df-4ea2-98db-66ce0f3fbdcd/serviceBalances"

$balanceJSON = $userBalance | ConvertTo-Json -Depth 5
```

Just cowork 


```powershell
$userCoworkBalance = Invoke-MgGraphRequest -Method GET -Uri "$baseUri/userBalances/fa936341-b3df-4ea2-98db-66ce0f3fbdcd/serviceBalances"

$coworkBalanceJSON = $userCoworkBalance | ConvertTo-Json -Depth 5
```

This retrn a payload similer to

```json
{
    "@odata.context": "https://graph.microsoft.com/beta/$metadata#copilot/costManagement/userBalances('fa936341-b3df-4ea2-98db-66ce0f3fbdcd')/serviceBalances/$entity",
    "@microsoft.graph.tips": "Use $select to choose only the properties your app needs, as this can lead to performance improvements. For example: GET copilot/costManagement/userBalances('<guid>')/serviceBalances('<key>')?$select=accessDecisionReason,configuredPeriodLimit",
    "serviceId": "cowork",
    "spendingPolicyId": "ee6ec031-acb2-4f7e-872d-923cf3ffd5ba",
    "consumedQuantity": 2000,
    "unit": "copilotCredit",
    "isConsumptionAllowed": false,
    "accessDecisionReason": "limitExhausted",
    "configuredPeriodLimit": {
        "@odata.type": "#microsoft.graph.limitedCostManagementQuantity",
        "value": 10000,
        "unit": "copilotCredit"
    },
    "configuredPerUserLimit": {
        "@odata.type": "#microsoft.graph.limitedCostManagementQuantity",
        "value": 2000,
        "unit": "copilotCredit"
    },
    "effectiveCapacityQuantity": {
        "@odata.type": "#microsoft.graph.limitedCostManagementQuantity",
        "value": 2000,
        "unit": "copilotCredit"
    },
    "remainingQuantity": {
        "@odata.type": "#microsoft.graph.limitedCostManagementQuantity",
        "value": 0,
        "unit": "copilotCredit"
    },
    "dataStatus": {
        "lastModifiedDateTime": "2026-10-19T11:06:38.8216976Z"
    }
}
```

## Spend policies


### List All

```powershell
$policies = Invoke-MgGraphRequest -Method GET -Uri "$baseUri/spendingPolicies"

$policiesJson = $policies | ConvertTo-Json -Depth 5
```


### List a policy


```powershell
$policy = Invoke-MgGraphRequest -Method GET -Uri "$baseUri/spendingPolicies/ee6ec031-acb2-4f7e-872d-923cf3ffd5ba"

$policyJson = $policy | ConvertTo-Json -Depth 5
```

Example spend policy

```json
{
  "createdDateTime": "2026-10-19T09:55:22Z",
  "perUserLimit": {
    "unit": "copilotCredit",
    "@odata.type": "#microsoft.graph.unlimitedCostManagementQuantity"
  },
  "createdBy": {
    "user": {
      "displayName": null,
      "id": "fa936341-b3df-4ea2-98db-66ce0f3fbdcd",
      "additionalData": []
    },
    "additionalData": [],
    "application": null,
    "device": null
  },
  "id": "52e309c0-8e7d-40e6-99ad-c8e14cb26700",
  "billingMethods": "payAsYouGo",
  "displayName": "TheAPIPolicy",
  "@odata.type": "#microsoft.graph.selectedGroupsSpendingPolicy",
  "@odata.etag": "\"ab023eb3-0000-0200-0000-6ad5e90e0000\"",
  "lastModifiedDateTime": "2026-10-19T09:55:22Z",
  "periodLimit": {
    "value": 1000.0,
    "unit": "copilotCredit",
    "@odata.type": "#microsoft.graph.limitedCostManagementQuantity"
  },
  "lastModifiedBy": {
    "user": {
      "displayName": null,
      "id": "fa936341-b3df-4ea2-98db-66ce0f3fbdcd",
      "additionalData": []
    },
    "additionalData": [],
    "application": null,
    "device": null
  },
  "@odata.context": "https://graph.microsoft.com/beta/$metadata#copilot/costManagement/spendingPolicies/$entity",
  "status": "enabled"
}
```

## Update a spend policy

Test updating the period limit on the policy returned by the single-policy GET.
This is a candidate payload based on that response; running it will test whether
the endpoint and current delegated permissions support the update.

Running the following changes the live policy limit to 1100 Copilot credits.
It leaves `perUserLimit` unchanged. Run the restore step in the same PowerShell
session after checking the result.

```powershell
$testPolicyId = '52e309c0-8e7d-40e6-99ad-c8e14cb26700'
$policyUri = "$baseUri/spendingPolicies/$testPolicyId"
$before = Invoke-MgGraphRequest -Method GET -Uri $policyUri -ErrorAction Stop

$newLimit = 1100
if ($before.periodLimit.value -eq $newLimit) {
    throw 'Choose a different test value so the read-back can prove a change.'
}

$restoreBody = @{
    periodLimit = @{
        '@odata.type' = '#microsoft.graph.limitedCostManagementQuantity'
        unit = 'copilotCredit'
        value = $before.periodLimit.value
    }
} | ConvertTo-Json -Depth 5

$body = @{
    periodLimit = @{
        '@odata.type' = '#microsoft.graph.limitedCostManagementQuantity'
        unit = 'copilotCredit'
        value = $newLimit
    }
} | ConvertTo-Json -Depth 5

Invoke-MgGraphRequest -Method PATCH -Uri $policyUri `
    -Headers @{ 'If-Match' = $before.'@odata.etag' } `
    -Body $body -ContentType 'application/json' -ErrorAction Stop

$after = Invoke-MgGraphRequest -Method GET -Uri $policyUri -ErrorAction Stop
if ($after.periodLimit.value -ne $newLimit) {
    throw 'The GET did not confirm the new limit. Inspect the policy before retrying.'
}

$after | Select-Object id, displayName, periodLimit, perUserLimit, lastModifiedDateTime |
    ConvertTo-Json -Depth 5
```


## Create a spend policy

Candidate endpoint:

```http
POST https://graph.microsoft.com/beta/copilot/costManagement/spendingPolicies
Content-Type: application/json
```

The [preview guide](./Copilot%20Usage-based%20Billing%20Graph%20API%20Preview%20Guide.pdf)
lists creation as planned Phase 1 functionality, but does not document the
request-body contract or write permissions. Confirm availability and write
authorization with Microsoft before testing. These beta APIs are for evaluation
only, not production use.

### Prepare the candidate payload

This proposes a group-targeted policy with a policy-level period limit of 20000
Copilot credits and a per-user limit of 2000, with only Cowork included in
`serviceSettings`. The spending period is not assumed to be monthly.

The initial POST returned `400 badRequest` because the required `serviceSettings`
property was missing. A GET of an existing policy's service settings returned
`serviceId = 'cowork'` and `state = 'enabled'`. Include those fields in the
candidate body, but omit the returned parent metadata (`parentEtag` and
`parentStatus`). The create-time shape still needs validation. Do not assume that
services omitted from the array are disabled; their default behavior is unverified.

The policy subtype is referenced by the guide's assigned-group endpoints.
`periodLimit` follows the existing update example; using the same quantity
structure for `perUserLimit` is an unverified assumption. Required fields,
create-time support, valid ranges, and enforcement/default behavior still need
confirmation. Do not copy server-generated IDs, timestamps, or ETags from a GET
response into the create body.

The API also confirmed that `displayName` must be 10-64 characters and contain
only letters, digits, and spaces. The example below uses a 42-character name with
no punctuation.

The following only prepares and displays JSON; it does not create a policy.

```powershell
$createBody = @{
    '@odata.type' = '#microsoft.graph.selectedGroupsSpendingPolicy'
    displayName = 'Finance Copilot spending policy evaluation'
    periodLimit = @{
        '@odata.type' = '#microsoft.graph.limitedCostManagementQuantity'
        unit = 'copilotCredit'
        value = 20000
    }
    perUserLimit = @{
        '@odata.type' = '#microsoft.graph.limitedCostManagementQuantity'
        unit = 'copilotCredit'
        value = 2000
    }
    serviceSettings = @(
        @{
            serviceId = 'cowork'
            state = 'enabled'
        }
    )
} | ConvertTo-Json -Depth 5

$createBody
```

### Submit and inspect

Running this sends a live create request using the delegated Graph connection
above. The existing read permissions do not establish write access. If successful,
this leaves a new policy in the tenant; there is no automatic cleanup.

Do not automatically retry a failed or ambiguous POST: check the policy list
first to avoid creating duplicates.

```powershell
$createdPolicy = Invoke-MgGraphRequest -Method POST -Uri "$baseUri/spendingPolicies" `
    -Body $createBody -ContentType 'application/json' -ErrorAction Stop

if (-not $createdPolicy.id) {
    throw 'The POST returned no policy ID. Inspect the response and policy list before retrying.'
}

$createdPolicyId = $createdPolicy.id
$createdPolicy | ConvertTo-Json -Depth 5

$readBack = Invoke-MgGraphRequest -Method GET `
    -Uri "$baseUri/spendingPolicies/$createdPolicyId" -ErrorAction Stop

$readBack | Select-Object id, displayName, periodLimit, perUserLimit, status |
    ConvertTo-Json -Depth 5

$createdServiceSettings = Invoke-MgGraphRequest -Method GET `
    -Uri "$baseUri/spendingPolicies/$createdPolicyId/serviceSettings" -ErrorAction Stop

$createdServiceSettings | ConvertTo-Json -Depth 10
```

Inspect the read-back for the requested name and limits, as well as any
server-applied defaults. Inspect the returned service settings to confirm Cowork
is enabled and determine how omitted services were configured. A successful create
alone does not prove that spending enforcement is configured as intended.

### Assigning a backing group

Run this in the same PowerShell session as the create step. Set `$groupId` to
the existing Entra group's object ID, not its name.
The block first fetches that group and validates the returned ID before attempting
assignment. The lookup requires permission to read groups in the signed-in tenant;
cost-management access alone does not establish that permission. A failed lookup
stops the block rather than attempting the assignment.
Running the block changes the live policy's group assignments; confirm the policy,
group, and service settings first. It does not change the group's membership.

The endpoint is listed in the guide; the conventional Graph `$ref` body below
still needs validation against the preview API.

The first assignment attempt, without `If-Match`, returned `412 Precondition
Failed`. The next test fetches the policy immediately before the POST and supplies
its current `@odata.etag` as `If-Match`, following the existing policy-update
pattern. This is the next candidate fix, not yet a confirmed assignment contract.

With `If-Match`, the attempt for group
`c7d70f70-3272-4a85-8564-97fd232e6a06` returned `404 Not Found`. A subsequent GET
of the assigned-groups collection succeeded with `"value": []`. This confirms the
read route resolves, but not the `$ref` write contract.

The second group, `babef59f-0102-4e8c-9223-c0b87de3313e`, resolved as a Microsoft
365 group (`Unified`, `securityEnabled: false`), but assignment also returned 404.
The UI reportedly accepts Teams-backed groups, so this does not establish a
security-group-only restriction.

The block below now tests the supplied security-group ID,
`71058377-197f-419b-bf27-bf14728a59b2`, individually. Only the group ID changes;
the endpoint, reference format, and ETag handling remain unchanged. This group's
GET and assignment results are not yet confirmed.

```powershell
$createdPolicyId = $createdPolicy.id
$groupId = '71058377-197f-419b-bf27-bf14728a59b2'

if (-not $createdPolicyId) {
    throw 'No created policy ID is available. Use the response from the successful create step.'
}

$group = Invoke-MgGraphRequest -Method GET `
    -Uri "https://graph.microsoft.com/beta/groups/$groupId" -ErrorAction Stop
if (-not $group.id -or $group.id -ne $groupId) {
    throw 'The group GET did not return the requested group. Check $groupId and the signed-in tenant before assigning it.'
}

$group | Select-Object id, displayName, securityEnabled, mailEnabled, groupTypes

$assignedGroupsUri = "$baseUri/spendingPolicies/$createdPolicyId/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups"
$groupReferenceBody = @{
    '@odata.id' = "https://graph.microsoft.com/beta/groups/$groupId"
} | ConvertTo-Json -Depth 5

$policyBeforeAssignment = Invoke-MgGraphRequest -Method GET `
    -Uri "$baseUri/spendingPolicies/$createdPolicyId" -ErrorAction Stop
$assignmentEtag = $policyBeforeAssignment.'@odata.etag'
if ([string]::IsNullOrWhiteSpace($assignmentEtag)) {
    throw 'The policy GET returned no ETag. Inspect the response before attempting the assignment.'
}

Invoke-MgGraphRequest -Method POST -Uri ($assignedGroupsUri + '/$ref') `
    -Headers @{ 'If-Match' = $assignmentEtag } `
    -Body $groupReferenceBody -ContentType 'application/json' -ErrorAction Stop

$nextGroupsUri = $assignedGroupsUri
$groupFound = $false
while ($nextGroupsUri) {
    $assignedGroups = Invoke-MgGraphRequest -Method GET -Uri $nextGroupsUri -ErrorAction Stop
    if ($groupId -in $assignedGroups.value.id) {
        $groupFound = $true
    }
    $assignedGroups | ConvertTo-Json -Depth 5
    $nextGroupsUri = $assignedGroups.'@odata.nextLink'
}

if (-not $groupFound) {
    throw 'The GET did not confirm the group assignment. Inspect the policy before retrying the POST.'
}
```

The single-quoted `/$ref` suffix is intentional: it prevents PowerShell from
interpreting `$ref` as a variable. The read-back follows any returned
`@odata.nextLink` unchanged and checks for the assigned group ID. If the POST
fails or its outcome is ambiguous, inspect the assigned groups before retrying.

Pass the ETag exactly as returned, including its quotes. Do not reuse the create
response's ETag or use `If-Match: *` to bypass concurrency checks. If a fresh ETag
still produces 412, inspect the current policy and group assignments before
another attempt; another write may have changed the policy, or the preview API
may require a different precondition.

Service settings have separate management endpoints, but the observed create
validation also requires `serviceSettings` in the initial body. Email notifications
have separate endpoints.

This exploration is an administrator-side evaluation, not a change to the
[owner portal specification](./SPEC.md), which prohibits policy creation by owners.

## Budgets 

Get budget for a user in a policy

This will happen as we get a policy, get thr groups then list the users.

The we can use `/userBalances/{userId}/serviceBalances` to get the balance.

Need to validate what policy takes presedent
