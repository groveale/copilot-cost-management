# Granting Copilot Cost Management permissions to Graph Explorer

Use this guide to grant Microsoft Graph **delegated permissions** to Graph
Explorer for administrator-side testing of the Copilot Cost Management preview
APIs.

Graph Explorer is a Microsoft-owned application. Its local **Enterprise
Application (service principal)** holds the consent grant in your tenant; you do
not need to create or modify an App Registration for Graph Explorer.

> The PowerShell fallback below changes live, tenant-wide consent. Review the
> requested permissions and tenant before running it. It does not create spending
> policies, change budgets, or assign groups.

## 1. Choose the permissions

Grant only what your testing requires. These are delegated work or school account
permissions, not application permissions:

| Permission | Purpose |
| --- | --- |
| `CopilotCostManagement-Policy.Read.All` | Read spending policies, service settings, alerts, and recipients. |
| `CopilotCostManagement-Policy.ReadWrite.All` | Create, update, and delete policies and manage their settings. |
| `CopilotCostManagement-Assignment.Read.All` | Read a policy's assigned groups. |
| `CopilotCostManagement-Assignment.ReadWrite.All` | Add and remove a policy's assigned groups. |
| `CopilotCostManagement-UserData.Read.All` | Read a user's balance diagnostics and service consumption. |

The script below includes all five permissions to match the full exploration
workflow. For read-only testing, remove the two `ReadWrite.All` entries. For a
single operation, select the least-privileged permission documented for that
operation rather than granting the entire set.

The separate `CopilotCostManagement.Read.All` permission is needed for the tenant
credit balance operation, `GET /beta/copilot/costManagement/getTenantCreditBalance`.
Add it to the script's `$newScopes` list only if needed. That operation also
supports application permissions, but this guide grants **delegated access only**.

Consent does not assign directory roles or enable preview access. The signed-in
user must also meet the endpoint's runtime authorization requirements. For
example, Microsoft documents Billing Administrator or Global Administrator as
required when creating a spending policy.

These APIs use **`/beta`**. Microsoft states that beta APIs are subject to change
and are not supported for production applications. See the
[Cost Management API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/cost-management/overview)
for current permissions, roles, and availability.

## 2. Try Graph Explorer's consent UI first

1. Open [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer).
2. Sign in with a work or school account and verify the displayed tenant.
   An unsigned sample-data session is not sufficient.
3. Enter a read-only request such as:

   ```http
   GET https://graph.microsoft.com/beta/copilot/costManagement/spendingPolicies
   ```

4. Open **Modify permissions** and consent to the required permission, using an
   authorized administrator when admin consent is required.
5. If it is not listed there, open the profile avatar's **Consent to permissions**
   list and search for the exact permission name.

Microsoft notes that **Modify permissions** can be missing permissions for some
queries. If the preview scopes are not exposed in the UI, use the administrator
PowerShell flow below. A missing UI entry does not prove the scope is available
in your tenant; the script checks Microsoft Graph's advertised delegated scopes.

## 3. Identify the right Enterprise Application

In the Microsoft Entra admin center, open **Enterprise applications**, find Graph
Explorer, and check its identifiers:

| Identifier | Value or meaning |
| --- | --- |
| Graph Explorer Application (client) ID | `de8bc8b5-d9f9-48b1-a8ad-b748da725064` |
| Graph Explorer Object ID | Tenant-specific service principal ID; this becomes the grant's `clientId`. |
| Microsoft Graph Application ID | `00000003-0000-0000-c000-000000000000` |
| Microsoft Graph Object ID | Tenant-specific resource service principal ID; this becomes the grant's `resourceId`. |

**Application ID and Object ID are not interchangeable.** The script looks up
both service principals by their known Application IDs and uses their returned
Object IDs in the grant.

If Graph Explorer is not present, first complete its approved sign-in/consent
onboarding in your tenant, then retry. The script intentionally does not create
service principals.

This grant applies to **Graph Explorer only**. It does not authorize the custom
client used in [the exploration notes](./explore.md) or another application's
PowerShell session.

## 4. Administrator prerequisites

Use PowerShell 7 and the following Microsoft Graph PowerShell modules. If they
are missing, install them for the current user:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Install-Module Microsoft.Graph.Applications -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.SignIns -Scope CurrentUser
```

The administrator connection requests:

- `Application.Read.All` to find service principals and inspect Graph's published
  delegated permission scopes.
- `DelegatedPermissionGrant.ReadWrite.All` to read, create, and update delegated
  consent grants.

These are **permissions for the administering PowerShell client**, not permissions
being added to Graph Explorer. `DelegatedPermissionGrant.ReadWrite.All` is highly
privileged; use it only in an approved administrator session.

The signed-in administrator also needs a supported Entra role and must satisfy
tenant consent policies. Microsoft lists roles including **Cloud Application
Administrator**, **Application Administrator**, and **Privileged Role
Administrator** for the grant create/update APIs. Activate an eligible role first
if your tenant uses Privileged Identity Management.

## 5. Create or update the tenant-wide grant

The relationship is:

```text
Graph Explorer service principal
    -> OAuth2PermissionGrant
    -> Microsoft Graph service principal
```

The safe sequence is **find the matching grant -> merge and update if present ->
create only if absent**.

This example deliberately uses `consentType = AllPrincipals`: administrator
consent for all users in the tenant, subject to their own authorization and the
application's access restrictions. It does not grant app-only access or give
every user an administrator role.

User-specific grants have `consentType = Principal` and a `principalId`. They can
coexist with the tenant-wide grant. The script neither modifies them nor mistakes
them for the tenant-wide grant. If you only want consent for one user, do not run
this tenant-wide example.

Replace `<TENANT-ID>`, review `$newScopes`, and run the following as one block.
It reads all grant pages, preserves existing scopes, and skips the update when
the requested scopes are already present.

```powershell
$tenantId = '<TENANT-ID>'
$newScopes = @(
    'CopilotCostManagement-Assignment.Read.All'
    'CopilotCostManagement-Assignment.ReadWrite.All'
    'CopilotCostManagement-Policy.Read.All'
    'CopilotCostManagement-Policy.ReadWrite.All'
    'CopilotCostManagement-UserData.Read.All'
)

Connect-MgGraph -TenantId $tenantId -ContextScope Process `
    -Scopes 'Application.Read.All', 'DelegatedPermissionGrant.ReadWrite.All' `
    -ErrorAction Stop

$clients = @(
    Get-MgServicePrincipal -Filter "appId eq 'de8bc8b5-d9f9-48b1-a8ad-b748da725064'" `
        -Property Id, AppId, DisplayName -All -ErrorAction Stop
)
if ($clients.Count -ne 1) {
    throw 'Expected one Graph Explorer service principal. Check the tenant and Graph Explorer onboarding.'
}
$clientSp = $clients[0]

$resources = @(
    Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" `
        -Property Id, AppId, DisplayName, Oauth2PermissionScopes -All -ErrorAction Stop
)
if ($resources.Count -ne 1) {
    throw 'Expected one Microsoft Graph service principal. Check the signed-in tenant.'
}
$graphSp = $resources[0]

$availableScopes = @(
    $graphSp.Oauth2PermissionScopes |
        Where-Object { $_.IsEnabled } |
        ForEach-Object { $_.Value }
)
$unavailableScopes = @($newScopes | Where-Object { $_ -cnotin $availableScopes })
if ($unavailableScopes.Count -gt 0) {
    throw "Graph does not advertise these enabled delegated scopes in this tenant: $($unavailableScopes -join ', '). Confirm preview availability and spelling."
}

$clientSp | Format-List DisplayName, AppId, Id
$graphSp | Format-List DisplayName, AppId, Id

$grants = @(
    Get-MgOauth2PermissionGrant -Filter "clientId eq '$($clientSp.Id)'" `
        -All -ErrorAction Stop |
        Where-Object {
            $_.ResourceId -eq $graphSp.Id -and
            $_.ConsentType -eq 'AllPrincipals' -and
            -not $_.PrincipalId
        }
)
if ($grants.Count -gt 1) {
    throw 'Multiple tenant-wide Graph grants matched. Inspect the grants before making changes.'
}

$grant = $null
$existingScopes = @()
if ($grants.Count -eq 1) {
    $grant = $grants[0]
    $existingScopes = @($grant.Scope -split '\s+' | Where-Object { $_ })
}
$expectedScopes = @($existingScopes + $newScopes | Sort-Object -Unique -CaseSensitive)
$combinedScopes = $expectedScopes -join ' '
if ($combinedScopes.Length -gt 3850) {
    throw 'The combined scope string exceeds the documented 3850-character limit. No grant was changed.'
}

if ($grant) {
    $missingScopes = @($newScopes | Where-Object { $_ -cnotin $existingScopes })
    if ($missingScopes.Count -gt 0) {
        Write-Host 'Updating the existing tenant-wide Graph grant, preserving its current scopes.'
        Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id `
            -BodyParameter @{ scope = $combinedScopes } -ErrorAction Stop
    }
    else {
        Write-Host 'The tenant-wide Graph grant already contains the requested scopes.'
    }
}
else {
    Write-Host 'Creating the tenant-wide Graph grant for Graph Explorer.'
    $grant = New-MgOauth2PermissionGrant -BodyParameter @{
        clientId = $clientSp.Id
        resourceId = $graphSp.Id
        consentType = 'AllPrincipals'
        principalId = $null
        scope = $combinedScopes
    } -ErrorAction Stop
}

if (-not $grant.Id) {
    throw 'No grant ID is available. Inspect the grants before retrying.'
}
$verified = Get-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id -ErrorAction Stop
if ($verified.ClientId -ne $clientSp.Id -or $verified.ResourceId -ne $graphSp.Id -or
    $verified.ConsentType -ne 'AllPrincipals' -or $verified.PrincipalId) {
    throw 'The read-back did not match the expected tenant-wide Graph Explorer grant.'
}
$verifiedScopes = @($verified.Scope -split '\s+' | Where-Object { $_ })
$missingAfter = @($expectedScopes | Where-Object { $_ -cnotin $verifiedScopes })
if ($missingAfter.Count -gt 0) {
    throw "Read-back is missing expected scopes: $($missingAfter -join ', '). Inspect the grant before retrying."
}
$verified | Format-List Id, ClientId, ResourceId, ConsentType, PrincipalId, Scope
```

The grant-management cmdlets use the Graph **v1.0** consent APIs; the Cost
Management requests themselves use **beta**.

### Why update instead of creating another grant?

Creating a duplicate grant can return:

```text
409 Conflict
Request_MultipleObjectsWithSameKeyValue
Permission entry already exists.
```

Do not make a duplicate POST a required step. Find the existing matching grant
and update it with `PATCH /v1.0/oauth2PermissionGrants/{id}`.

The `scope` property is a **space-separated replacement value**, not an append
operation. Sending only the five new scopes would remove other scopes from that
grant. The script merges the original and requested scopes before updating.

This is safe to rerun sequentially, but it is not an atomic concurrent merge.
Coordinate consent changes with other administrators. If creation races with
another administrator and returns 409, re-read the grants and rerun the
create-or-update flow; do not blindly repeat the POST. For an ambiguous failure,
inspect the current grant first.

## 6. Refresh Graph Explorer and verify actual access

1. Confirm the script's read-back contains the expected original and new scopes.
2. Sign out of Graph Explorer and sign back in to the **same tenant** to acquire
   a fresh token. An already-issued token is not updated by changing consent.
3. Request/consent to the required scopes in Graph Explorer if they are selectable.
   Consent stored on a service principal and scopes requested for a token are
   separate concerns; a successful grant update alone does not prove that Graph
   Explorer's token includes those scopes.
4. Use Graph Explorer's **Access token** tab to inspect the token locally. Check
   its `tid`, its Graph Explorer client identifier (`appid` or `azp`), its
   Microsoft Graph audience (`aud`), and that `scp` includes the delegated scope
   required for the request. Treat access tokens as secrets; do not paste them
   into third-party decoders, tickets, or documentation.
5. Run the read-only spending-policies GET from section 2. A successful response
   verifies that request's access, not every write or diagnostic operation.

If Graph Explorer still cannot request a preview scope or its fresh token lacks
it, stop and investigate the client/preview limitation. Do not keep adding
consent or elevate directory roles to compensate for a missing token scope.

### Troubleshooting

| Symptom | What to check |
| --- | --- |
| Scope missing from Graph Explorer's UI | Check both permission lists; use the PowerShell fallback only if Graph advertises the enabled scope in the tenant. |
| Script reports an unavailable scope | Confirm spelling and preview availability with Microsoft; do not bypass the check or substitute an application permission. |
| `403` during service principal lookup or grant management | Check the administering PowerShell client's scopes, signed-in tenant, active Entra role, and consent policies. |
| `409` / permission entry already exists | Locate the matching `AllPrincipals` grant, merge scopes, and PATCH rather than creating a duplicate. |
| Grant looks correct but Cost Management returns `403` | Check the fresh token's scopes, the user's endpoint-specific runtime role, and preview access. |
| Cost Management returns `404` | Check the route, resource ID, tenant, preview availability, and authorization; the API can hide out-of-scope resources with 404. |

After completing the administrator work, run `Disconnect-MgGraph` to end that
PowerShell connection. This does **not** revoke the persisted consent grant.

## References

- [Work with Graph Explorer](https://learn.microsoft.com/en-us/graph/graph-explorer/graph-explorer-features)
- [Create a delegated permission grant](https://learn.microsoft.com/en-us/graph/api/oauth2permissiongrant-post?view=graph-rest-1.0)
- [Update a delegated permission grant](https://learn.microsoft.com/en-us/graph/api/oauth2permissiongrant-update?view=graph-rest-1.0)
- [Microsoft 365 Copilot Cost Management API overview](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/cost-management/overview)
