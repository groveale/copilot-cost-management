# Administrator SharePoint data sync

The admin **Sync data** section collects policy configuration, assigned-group
users and user/service balances into three SharePoint lists. It is an
evaluation feature over the **beta Copilot Cost Management APIs**, not a
production-support guarantee, reconciled billing report or historical ledger.

## Execution and identity

Work runs in the signed-in administrator's browser using delegated Graph access.
Keep the browser active/open. Closing it, suspension of a background tab,
interaction-required authentication or an expired execution lease stops work.
There is **no unattended Azure queue worker**.

The Function API stores authoritative job state in the private `SyncJobs` Table
and immutable checkpoint blobs in `sync-checkpoints`. It uses the existing
system-assigned identity and private Blob/Table endpoints. Collection tokens stay
managed by browser MSAL. The narrowly scoped, explicitly enabled permission
service uses **delegated OBO** to SharePoint REST and Graph user-ID resolution;
it is not a collection proxy or unattended worker. There is no persisted
downstream token cache, client secret, copied refresh token, or app-only Graph/
SharePoint permission grant.

Only IDs in `ONBOARDING_ADMIN_OBJECT_IDS` can use the sync API. Any currently
authorized administrator can resume after fresh validation and lease acquisition.
One tenant execution gate prevents overlapping all-policy/single-policy work.
A live lease cannot be taken over; pause the original browser, explicitly cancel,
or wait for expiry. Leases last 90 seconds and active execution renews them.

## Setup

1. Choose an **existing, restricted SharePoint site** for reporting. Do not use
   an employee-wide readable/writable site for identifiable balance data.
2. Use an approved operator public-client registration with delegated
   `Sites.Manage.All` for provisioning. The operator must also have sufficient
   SharePoint rights. Explicit site-grant administration requires
   `Sites.FullControl.All` and an appropriately authorized operator.
3. Configure the runtime SPA's additional delegated permissions:

   ```powershell
   .\scripts\configure-entra.ps1 -SpaClientId '<spa-app-id>' -EnableSync -EnableReportingSecurity
   .\scripts\configure-entra.ps1 -SpaClientId '<spa-app-id>' -EnableSync -EnableReportingSecurity -Apply
   ```

   This preserves existing permissions/callbacks. Arrange admin consent through
   the organization's approval process. Adding required permissions is not itself
   consent or PIM activation.

4. Preview, then apply list setup using PowerShell 7.4+ and the existing
   `Microsoft.Graph.Authentication` module, version 2.37.0 or later:

   ```powershell
   $setup = @{
       TenantId = '<tenant-guid>'
       ClientId = '<approved-operator-public-client-guid>'
       SiteUrl = 'https://contoso.sharepoint.com/sites/copilot-reporting'
       RuntimeAppId = '<runtime-spa-app-guid>'
       GrantSiteAccess = $true
       LockDownReporting = $true
       AdminObjectIds = @('<same-admin-user-object-ids-as-the-private-API-configuration>')
       RetentionDays = 30
       DefaultConcurrency = 3
       MaxConcurrency = 8
   }
   .\scripts\setup-sync-sharepoint.ps1 @setup
   .\scripts\setup-sync-sharepoint.ps1 @setup -Apply
   ```

   `GrantSiteAccess` is optional if a site administrator manages the explicit
   grant separately. Runtime consent to `Sites.Selected` and the selected-site
   **write** grant are both necessary. Delegated effective access is the
   intersection of that app grant and the executing user's SharePoint rights.

   Lockdown additionally uses `PnP.PowerShell` 3.1.0+ with an interactive
   delegated connection. The approved operator registration needs SharePoint
   **AllSites.FullControl** (a SharePoint resource permission, not the similarly
   named Graph scope), and the executing user must be a configured administrator
   with sufficient site rights. Graph `User.ReadBasic.All` resolves the explicit
   administrator object IDs. Preview reports the destructive-to-access change
   without changing ACLs. Apply creates/reuses an app-marked SharePoint
   administrator group, sets its owner to a configured administrator, disallows
   self-join/member edits, reconciles members and removes non-administrator
   list grants. It resets **all existing item scopes**, including stale/staging
   generations, to the administrator-only list ACL and reads back the result.
   No rows or lists are deleted. Owners remain blocked until API verification.

   The script creates/verifies the lists, types and indexes; it does not delete
   or recreate incompatible lists. Rerunning is supported. Review partial setup
   failures and rerun after correcting authorization or schema issues.

5. Store the script's `SYNC_DESTINATION_JSON` output as the GitHub environment/
   repository secret of the same name. Set the non-secret configuration
   variables `SYNC_DEFAULT_CONCURRENCY`, `SYNC_MAX_CONCURRENCY` and
   `SYNC_JOB_RETENTION_DAYS` to the desired output values.
6. Review/apply infrastructure and the separately privileged access template.
   The Function needs Table Data Contributor on **SyncJobs** and Blob Data
   Contributor on **sync-checkpoints**, individually scoped. Existing deployment
   steps still apply; the workflow does not acquire RBAC authority.
7. Configure the certificate-backed permission service below, then deploy
   application code and use **Verify SharePoint permissions**. The admin settings surface reports disabled until
   all four destination settings are supplied. Without GitHub, pass the secure
   `syncDestination` Bicep object and concurrency/retention parameters directly
   through the normal reviewed deployment process.

Runtime app settings are `SYNC_SITE_ID`, `SYNC_POLICY_LIST_ID`,
`SYNC_BALANCE_LIST_ID`, `SYNC_JOB_LIST_ID`, `SYNC_DEFAULT_CONCURRENCY`,
`SYNC_MAX_CONCURRENCY`, and `SYNC_JOB_RETENTION_DAYS`. Defaults are 3 requests,
a maximum of 8, and 30 days of operational job history. All destination IDs must
be supplied together. Only the authenticated reporting settings endpoint returns
the browser-safe site and two reporting-list IDs to administrators or verified
current owners; task-list IDs and administrator mappings are never public.

Review **List settings > Versioning settings** and site retention policies.
Disable unnecessary reporting-list version history where organizational policy
permits; Graph list provisioning does not configure this setting. Deletion is
subject to SharePoint recycle-bin/retention holds and is not guaranteed physical
erasure.

## Permission-service opt-in and upgrade

The confidential **API registration**, not the SPA or managed identity, requires
SharePoint delegated `AllSites.FullControl` and Graph delegated
`User.ReadBasic.All`. `-EnableReportingSecurity` adds those required permissions
without granting consent or creating credentials. Arrange tenant administrator
consent separately. SharePoint effective rights still depend on the signed-in
administrator (including PIM/Conditional Access); API administration alone does
not make someone a SharePoint/site-collection administrator. Snapshot readers
use only browser Graph `Sites.Selected`, not these powerful service scopes.
Do not use Graph listItem permission creation: it creates application grants,
not the required real SharePoint user ACLs.

`infra/main.bicep` accepts an explicit secure `reportingSecurity` object:

```json
{
  "enabled": true,
  "siteUrl": "https://contoso.sharepoint.com/sites/copilot-reporting",
  "adminGroupId": 42,
  "certificateId": "https://dedicated-vault.vault.azure.net/certificates/reporting/<32-hex-version>"
}
```

Use the lockdown script's group ID and exact site URL. Store the complete object
as `REPORTING_SECURITY_JSON` for the reviewed infrastructure workflow, which
hash-binds it between what-if and apply without storing the private object in
the artifact. Without GitHub, pass the secure Bicep parameter directly.
Empty/default configuration does **not** make permissions ready; security POSTs
return `reporting-not-configured` rather than pretending to succeed.

Enabling this object provisions a **dedicated** Key Vault with private endpoint/
DNS, RBAC, purge protection and public networking disabled. First provision it
through a reviewed direct deployment with `{ "enabled": true }` (the runtime is
deliberately incomplete), then, from an authorized network-connected operator,
create a nonexportable RSA signing certificate and register **only its public
certificate** on the API Entra application. No certificate is created or
registered automatically. Apply the separately privileged `infra/access.bicep`
with `reportingVaultName` from deployment outputs. Its custom runtime role has
only public-certificate read and cryptographic sign data actions on this
dedicated vault: no secret/private-key retrieval, certificate/key management or
RBAC administration. Keep unrelated keys out of this vault. Finally supply the
complete, **versioned** certificate URI in the normal reviewed deployment.

Runtime settings are `REPORTING_SECURITY_ENABLED=true`, `REPORTING_SITE_URL`,
`REPORTING_ADMIN_GROUP_ID`, and `REPORTING_CERTIFICATE_ID`. The API reads the
public certificate and sends a SHA-256 digest to its Key Vault key's sign
operation, then exchanges the validated incoming API assertion through OBO.
It never exports the private key. Tokens exist only in the current request.
Certificate expiry/sign failures fail closed. For rotation, create a new
nonexportable certificate version, add its public certificate to the API app,
update the pinned version URI, verify permissions, then retire the prior app
certificate after validation/overlap. Renewal alone does not update a pinned URI.

The administrator **Verify SharePoint permissions** action is both the upgrade
migration and manual revocation-recovery entry point. Sync invokes it before
collection and again after staging each policy but **before** publishing its
active pointer. The server accepts only optional policy/job/lease identifiers,
uses configured destinations, and derives every owner from current private
budget metadata. A supplied job/lease must match the current tenant job, current
executor and an unexpired running lease (checked again before readiness).
The reconciler intentionally verifies all lists/all rows even for a policy call,
including superseded, abandoned and staged generations. Malformed/foreign and
unmapped rows and every SyncJobs row are administrator-only. Master and balance
rows for mapped policies grant their current owners **Read**, not edit; no
employee self-read grant is added. New rows inherit administrator-only lists.
Expected item ACLs are applied immediately before publication; an owner may see
same-policy staged rows directly at that point, but never another policy's rows.

Lists break inheritance without retaining broad site-member/visitor grants.
The only broad data grant is the verified administrator group. SharePoint's
implicit **Limited Access** navigation grants on parents are allowed; these are
not list-data read grants. Site collection administrators remain privileged.
The reconciler also checks the underlying Read/Limited Access permission masks,
rejecting customized levels that would add writes or broad data-read access.
Every role change is read back, and administrator group membership/owner/
self-join settings are verified. No policy fields, snapshots or publication
fences are changed by reconciliation. Check item ETag behavior in the live tenant;
permission updates are separate from field updates, not a transaction with them.

Readiness is stored in private `metadata/reporting-security.json`, tied to both
the authoritative metadata ETag/ownership hash and complete service configuration.
No environment flag or browser claim can mark it ready. The service uses
renewable Blob leases on that state and existing metadata. Metadata writes from
onboarding completion/bootstrap invalidate readiness and revoke/read back **all**
existing owner grants before committing a changed ownership authority. A partial
revocation prevents that metadata commit and leaves readiness blocked. Retry
verification after fixing access; do not remove the marker to bypass a failure.
Direct `seed.js` replacement is blocked after security initialization. Retire
direct seed-operator metadata-write RBAC after initial setup; privileged manual
Blob edits bypass application controls and are not an ownership-change workflow.
Disabling a setting does not revoke grants or allow bypassing an existing marker.
Before removing the feature/configuration, rerun explicit list lockdown to revoke
all owners, then disable it. SharePoint propagation/caching can delay revocation;
consider tenant session revocation for incident response rather than claiming
instant removal from previously downloaded data.

Isolation consumes one unique scope per list/item. Runtime refuses expansion
when **existing rows + list scope + reserved staging headroom** exceed
`REPORTING_MAX_SCOPES` (default/hard maximum **5,000**) with
`REPORTING_SCOPE_RESERVE` (default **250**, must be positive). Stale/staged rows
count too. At the limit no pointer should be published and no isolation is
silently dropped. Administratively remove obsolete generations using the existing
maintenance process or adopt a separately approved policy-folder/partitioning
design. Microsoft's supported 50,000 limit is not this design's safe operating
target. Large full migrations may require repeated repair after request/lease
timeouts; each failure leaves readiness blocked, never simulated ready.

`reporting-interaction-required` includes an allowlisted `WWW-Authenticate`
claims challenge (exposed through CORS). Use the explicit refresh API sign-in/
PIM control, then retry; the service never launches an interactive prompt.
Invalid/untrusted continuation links and certificate/key endpoints are rejected
before forwarding a bearer token. No downstream token or raw error description
is returned to the browser or written to logs.
UPNs are used only to resolve SharePoint principals: their returned
`AadObjectId.NameId` must match the server's directory object ID. A reused/stale
site-user login fails closed rather than granting an old identity access.
When changing configured administrators, run explicit lockdown with the new
administrator IDs before updating deployment configuration, then verify again;
changing an app setting alone is not a SharePoint revocation.

## Choosing a reporting source

Signed-in **Settings** (cog) remembers **Real time** (default) or **Sync** in this
browser, and **Light / Dark / System** appearance. Failure to save a preference
is visible and retains the previous choice. Diagnostics show only allowlisted
decoded ID-token claims and in-memory MSAL scopes/acquisition/expiry information,
never raw bearer/refresh/ID-token strings; diagnostics are not proof of PIM.

Real time retains live Graph reporting. Sync reads delegated Graph
`Sites.Selected` snapshots for the policy's `ActiveSnapshotId`, validates keys
and generations, and shows dated normalized observations/freshness/exclusions.
It makes no directory-name lookups and never falls back to live Cost Management
APIs after missing, denied, stale or inconsistent snapshot data. Switching
source/account discards in-flight responses. Owners are blocked until current
server ownership and persisted permission verification agree; administrators
can inspect/recover configuration. **Personal Sync is administrator-only in v1.**

Requests, approvals, onboarding and owner actions remain live. Owner actions
require explicit fresh live consent/authorization, even when viewing Sync.
A successful live mutation does not change a snapshot; only a subsequent sync
does. Synced membership is discovered/transitive, not a claim of current direct
membership.

## Live security acceptance (separate authorization required)

Local mocks/builds do not establish tenant consent, SharePoint role propagation,
PIM behavior, Key Vault network access or real ACL isolation. Before owner rollout,
an independently authorized operator must verify admin access to all rows; owner
direct Graph/REST denial on another policy, unmapped data and SyncJobs; stale/
staged-row coverage; new-owner/old-owner revocation; mid-revocation failure;
conditional-access challenge refresh; certificate rotation; scope-headroom refusal;
and staging/pointer ETag behavior. No cloud/site/consent changes or live acceptance
are implied by deploying source changes.

## Lists and reporting semantics

| List | Shape |
| --- | --- |
| CopilotSpendingPolicies | One tenant/policy master row: configuration/service settings, groups, optional portal mappings, active snapshot ID, coverage and per-service/unit totals. |
| CopilotUserBalances | One policy/snapshot/user/service observation, with group provenance, reported policy, quantities, units, availability and source/collection timestamps. |
| CopilotSyncJobs | Operational job history: scope, actors, status, concurrency, counts and compact outcomes; no historical balances or authentication tokens. |

`PolicyKey`, `ObservationKey` and `JobId` are unique indexed keys. Partition/
snapshot fields are indexed for filtered retrieval. Large variable configuration
and service totals use bounded JSON columns; oversized values fail explicitly
rather than being truncated. Job/checkpoint bodies are bounded to 1 MiB; use
single-policy syncs if a tenant inventory exceeds that limit.

**Join balances using the policy's `ActiveSnapshotId`**, with the same tenant and
policy IDs. New observations are staged before that pointer and totals are
updated together on the master row using an ETag and execution fence. A raw
unfiltered list view can contain staging or superseded rows during publication/
cleanup; SharePoint has no transaction spanning the two lists.

Each successful policy is published independently. A failed/incomplete policy
keeps its previous successful snapshot. A lost publication acknowledgement is
reconciled before recollection. Superseded rows are deleted only after a new
generation is active; cleanup failures remain visible and retryable.

Collection follows:

1. All policy inventory pages, or one explicitly selected policy.
2. Supported policy configuration/service settings and every assigned-group page.
3. Each group's transitive membership, selecting user objects locally.
4. All service-balance pages for every unique user within that policy.

Nested/overlapping groups do not multiply a user's consumption. Returned balances
are attributed only when their `spendingPolicyId` matches the discovery policy.
Mismatches and users with no returned balance are visible exclusions. Duplicate/
malformed service rows, unreadable membership, failed pages and invalid quantities
are failures, not zero consumption. Limited, unlimited and missing remaining
quantities are represented separately.

Totals are separate for every **service and reported unit**. Do not add units
together or treat credit quantities as USD. Policy totals are group-discovered
observations, not a reconstruction of consumption by users who already left.
Collection is not instantaneous; inspect collection/source freshness timestamps.

Tenant-wide and unknown policy types get inventory/configuration only:
consumption is **not collected**, not zero. A complete all-policy inventory can
mark previously seen policies **stale** without deleting their last snapshot.
Single-policy or incomplete inventory cannot establish absence.

Hidden/limited directory visibility can limit discovery. The application does
not automatically request hidden-membership access or prove invisible members
do not exist.

## PIM and validation

Use **Validate access** before starting. The app obtains fresh delegated access,
checks the API admin capability, list keys/publication columns/indexes and real
policy/group/member/balance routes when representative resources exist. Missing
probe resources are explicitly **untested**, not proof of authorization.
The first SyncJobs metadata write establishes destination-write readiness;
per-list write failures still prevent affected snapshot publication.

An unexpired token or granted consent does not prove active directory roles.
If access is denied:

1. Check Graph consent, the target site's app grant and the user's site rights.
2. Activate the required eligible role in Microsoft Entra PIM, if applicable.
   The app neither guesses nor activates a role automatically.
3. Click **Refresh access / sign in**, which bypasses the cached access token.
   Complete an explicit sign-in/MFA prompt if MSAL requires interaction.
4. Revalidate and resume.

Claims challenges are passed to MSAL without following a supplied authority URL.
Ordinary 403 errors are not mislabeled as proof of PIM failure. Requests continue
to enforce access throughout the run; preflight cannot guarantee future
authorization. No tenant protections are relaxed to keep a sync running.

## Progress, takeover and recovery

- Concurrency is a global ceiling across sync Graph reads and SharePoint writes:
  default 3, selectable 1-8 within the configured maximum.
- GETs retry bounded transient failures and honor `Retry-After`. Long throttling
  pauses collection instead of ignoring the server's requested delay.
- Writes use deterministic keys, ETags and exact readback after uncertain
  responses; they are not blindly retried.
- Pause/resume retains completed policies. An interrupted unfinished policy
  restarts under a new generation, including on administrator takeover; the app
  does not promise per-user continuation after the browser is closed.
- The private API state is authoritative. A failed SharePoint job mirror is shown
  as out of date. **Reconcile this job** repairs a retained summary without
  rerunning successful collection.
- **Retry failed policies** preserves successful outcomes in the same job.
- **Reconcile reporting / cleanup** on the current finished task obtains an
  exclusive maintenance lease, cleans superseded/abandoned observations for that
  job's policies, and prunes terminal job history/checkpoint revisions older than
  the configured retention. The current task and referenced state are retained.
- Retention cleanup is operator-driven through this action, not scheduled. Run it
  routinely; no browser task or cleanup runs unattended.
- Direct SharePoint access remains governed by site/list permissions, not by the
  portal's admin check.

## Validation before live use

Run focused frontend/API sync tests, existing builds/type checks, deployment
contract/publication checks and `scripts\sync-sharepoint.Tests.ps1` with Pester.

In an explicitly authorized preview tenant/test site, verify setup idempotence,
all-policy and single-policy sync, exact deduplicated totals, unavailable group
handling, browser closure/resume, another admin's takeover, and real PIM
activation/expiry where available. Confirm that a failed source page or
destination write preserves the previous active snapshot. Mocked tests cannot
prove tenant-specific preview API availability, site consent or real PIM behavior.
