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
system-assigned identity and private Blob/Table endpoints. Graph tokens stay
managed by browser MSAL; there is no server-side Graph token cache, OBO client,
client secret, copied refresh token, or app-only Graph access.

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
   .\scripts\configure-entra.ps1 -SpaClientId '<spa-app-id>' -EnableSync
   .\scripts\configure-entra.ps1 -SpaClientId '<spa-app-id>' -EnableSync -Apply
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
7. Deploy application code. The admin settings surface reports disabled until
   all four destination settings are supplied. Without GitHub, pass the secure
   `syncDestination` Bicep object and concurrency/retention parameters directly
   through the normal reviewed deployment process.

Runtime app settings are `SYNC_SITE_ID`, `SYNC_POLICY_LIST_ID`,
`SYNC_BALANCE_LIST_ID`, `SYNC_JOB_LIST_ID`, `SYNC_DEFAULT_CONCURRENCY`,
`SYNC_MAX_CONCURRENCY`, and `SYNC_JOB_RETENTION_DAYS`. Defaults are 3 requests,
a maximum of 8, and 30 days of operational job history. All destination IDs must
be supplied together. None belong in public frontend configuration.

Review **List settings > Versioning settings** and site retention policies.
Disable unnecessary reporting-list version history where organizational policy
permits; Graph list provisioning does not configure this setting. Deletion is
subject to SharePoint recycle-bin/retention holds and is not guaranteed physical
erasure.

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
