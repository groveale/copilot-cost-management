# Deploy the single-tenant demo

## Architecture and limits

- Static Web Apps **Free**, West Europe resource location, globally served assets.
- Standalone Linux Functions **Flex Consumption / Node 24**, UK South.
- Separate Standard LRS host/package and business storage accounts, UK South.
- Dedicated UK South VNet, outbound Flex integration and private Blob/Table
  endpoints with private DNS. Both storage accounts disable public networking;
  private endpoints/DNS have ongoing charges even when the Function scales to zero.
- Private JSON blob for budget/owner metadata; `RequesterBudgetIndex`,
  `Requests` and `OnboardingRequests` Tables. No Graph proxy or application
  Graph permissions. New BU provisioning/allocation and approved credit-request
  membership moves use delegated Graph in the administrator/owner browser.
- UK South Application Insights/Log Analytics; 30-day retention, sampling and a
  small workspace ingestion cap. These are **not hard spending caps**.
- Function runtime uses its **system-assigned managed identity** for host state,
  deployment packages, Tables, private blobs and Entra-authenticated telemetry.
  Shared Key, anonymous blobs, local telemetry auth and SCM/FTP basic publishing
  are disabled. There are no key/SAS/client-secret fallback paths.
- Browser API calls use a separate delegated Entra token, not Function keys,
  Graph tokens, ID tokens or the Function's identity.
- GitHub has separate infrastructure/application federated identities. The only
  deployment-token exception is the site-scoped **frontend-only SWA token**.
- Allowance-increase approval records a decision and applies the mapped
  Standard-to-Enhanced group move in the owner's browser. Temporary moves require
  manual return to Standard; no expiry worker or automatic reversion is provided.
  BU onboarding approval is a separate real provisioning workflow. No notification,
  backups or recovery SLA.

## Before provisioning

Use an authorized operator with subscription deployment/RBAC permissions and an
Entra administrator for registrations/consent. Owners **may approve their own
requests for this MVP**, so the owner account can exercise the full flow.
An additional non-owner account is useful for live access-isolation testing. Recheck pricing,
Flex quota and Node 24 availability in UK South.

Review the public publication set. `*.private.json`, local settings, dependencies
and build artifacts are ignored. Confidential local reference files remain
excluded. Existing published history is not made private by these exclusions;
there is no automatic history rewrite.

Install locked dependencies and validate:

```powershell
npm ci --prefix demo
npm ci --prefix api
npm test --prefix demo
npm test --prefix api
npm run build --prefix demo
npm run check --prefix api
npm run build --prefix api
node scripts\check-public.mjs
node --test scripts\deployment.test.mjs
az bicep build --file infra\main.bicep --stdout
az bicep build --file infra\access.bicep --stdout
az bicep build --file infra\bootstrap.bicep --stdout
az bicep build --file infra\network.bicep --stdout
```

## Entra setup (separate from ARM)

1. Reuse the existing SPA registration and retain its localhost SPA callback.
2. Create a single-tenant API registration, set
   `api.requestedAccessTokenVersion` to `2`, and set its identifier URI to
   `api://<API-client-ID>`. Expose a delegated `access_as_user` scope. The API
   validates v2 tokens with **API client ID** as audience.
3. Add that permission to the SPA without replacing its existing Graph
   `requiredResourceAccess` entries. Create the API enterprise application and
   arrange consent/preauthorization with the authorized administrator.
4. After deployment outputs exist, add `https://<site>/auth.html` as a SPA
   redirect. Do not remove other valid callbacks or guess the hostname.

Request submission needs no Graph permissions. The Request credits balance panel
separately requests only the existing delegated
`CopilotCostManagement-UserData.Read.All` and
`CopilotCostManagement-Policy.Read.All` scopes, not the owner's eight-scope set.
For this demo, use an account already permitted to read its balance and policy,
including any required directory role. Ordinary-user self-read support is an
assumption, not verified by this application; do not grant all employees broad
roles just for the panel. Missing access is explicit and does not block request
submission. Reads use silent acquisition; only **Grant balance read access** may
open an interaction-required popup. Current owners acquire their separate
eight-scope Graph token for owner operations.
`scripts/configure-entra.ps1 -SpaClientId <id>` previews these operations;
`-Apply` creates/reuses the API, preauthorizes only this SPA and preserves its
existing Graph permissions. Rerun with `-SiteUrl <actual-origin> -Apply` after
provisioning to add callbacks. It does not add Graph application permissions.

## Initial infrastructure

Review each what-if before its apply. Use incremental deployments only:

For a completely new resource group, `infra/environment.bicep` supports
subscription-scoped what-if **before any Azure resource is created**, then
creates the group, identities and application resources. Supply `tenantId`,
`spaClientId` and the actual `apiClientId`; use `az deployment sub what-if`
and `az deployment sub create --location uksouth --template-file
infra/environment.bicep`. Apply `access.bicep` separately afterward.
The alternative operator sequence for an existing/manually-created group is:

```powershell
az account set --subscription <subscription-id>
az group create --name rg-copilot-cost-demo-uks --location uksouth
az deployment group what-if --resource-group rg-copilot-cost-demo-uks --template-file infra\bootstrap.bicep
az deployment group create --name bootstrap --resource-group rg-copilot-cost-demo-uks --template-file infra\bootstrap.bicep

$env:AZURE_TENANT_ID = '<tenant-id>'
$env:SPA_CLIENT_ID = '<spa-client-id>'
$env:API_CLIENT_ID = '<api-client-id>'
az deployment group what-if --resource-group rg-copilot-cost-demo-uks --parameters infra\demo.bicepparam
az deployment group create --name initial --resource-group rg-copilot-cost-demo-uks --parameters infra\demo.bicepparam --mode Incremental
```

Use bootstrap principal outputs and initial resource outputs to run the separate
privileged `access.bicep` template: `functionName`, **functionPrincipalId**,
`hostStorageName`, `businessStorageName`, `insightsName`, `infraPrincipalId`,
`appPrincipalId`, and optionally `seedOperatorId` (a user object ID).
Review its what-if first. No normal CI job has role-assignment authority.

Runtime scope: Blob Data Owner on host storage for this HTTP-only Functions host
and package read/write; Table Data Contributor on the **Requests**,
**OnboardingRequests** and **RequesterBudgetIndex** tables individually; Blob
Data Contributor on the **metadata container only**; Monitoring Metrics Publisher
on Insights. These narrowly scoped writes support ongoing onboarding, not only
one-off seeding. There are no triggers
requiring queue/durable/table host permissions. The seed operator gets write
access to the index and metadata, not the request store.

Role assignment names incorporate scope, role and principal ID. Creating the
Function establishes the identity before RBAC is applied; initial startup can
fail until access propagates. Retry readiness with bounded backoff (for example
5, 10, 20, 30, 60 seconds), then stop with an explicit error. Never enable keys to
work around propagation. If the Function is recreated, obtain its **new**
principal ID, review/apply access again, and explicitly review any orphaned old
assignments. Routine code deployment preserves the Function resource/identity.

`main.bicep` sets exact-origin CORS from the actual SWA output; localhost is not
allowed in the hosted API. Public bootstrap generation uses those real outputs.
This is a direct API URL, not SWA-linked `/api`; Free hosting is intentional.

## Private storage networking

Inherited governance enforces disabled public storage access. Do not request a
policy exemption, use exemption tags, enable public storage, or introduce keys
as a deployment workaround. `main.bicep` explicitly disables storage public
networking and firewall bypasses, and deploys [network.bicep](../infra/network.bicep):

- Isolated VNet `10.84.0.0/24`; no peering, VPN, public IP, NAT gateway or Dev Box.
  Review address overlap before connecting it to another network.
- `functions`: `10.84.0.0/27`, delegated to **Microsoft.App/environments** for
  Flex (not the Premium plan's Microsoft.Web/serverFarms delegation).
- `private-endpoints`: `10.84.0.32/27`, separate from the delegated subnet.
  Both subnets have NSGs with Azure's default rules, not custom internet ingress.
- Three private endpoints: host/package **Blob**, business **Blob**, business
  **Table**. This HTTP-only app has no Azure Files share, queue or durable
  triggers; do not add endpoints or permissions for unused services.
- `privatelink.blob.core.windows.net` and `privatelink.table.core.windows.net`
  private DNS zones linked to the VNet, with endpoint-managed A records.
- The Function integrates with the delegated subnet after endpoint/DNS
  provisioning, retaining its existing system-assigned principal and RBAC.

Keep ordinary `*.blob.core.windows.net` / `*.table.core.windows.net` application
URLs: DNS inside the VNet resolves them through Private Link. Do not hardcode
private IP addresses or replace SDK hostnames with `privatelink` URLs.
VNet integration controls **outbound** access, not browser/API authentication;
it does not create an inbound Function private endpoint. The existing public
Entra-protected API and exact-origin CORS remain unchanged.

Register `Microsoft.App` and `Microsoft.Network` before deployment. Run a fresh
what-if, review only expected additions/changes, then deploy incrementally.
Never recreate storage or the Function as a networking fix. The infrastructure
workflow packages and runs [verify-network.ps1](../scripts/verify-network.ps1)
after apply; operators can also run it with `-ResourceGroup` and
`-DeploymentName`. It verifies storage restrictions, retained identity,
integration/delegation, endpoint approvals, NIC IP/DNS agreement and VNet links.
These are **control-plane** checks, not proof of runtime data access.

GitHub-hosted runners can apply ARM/Bicep without direct storage access. Function
OneDeploy still requires a reachable deployment endpoint and a verified
Azure-side path to private package storage. Jobs that directly seed/read Tables
or blobs must execute inside an approved connected network. There is no implicit
VPN for this laptop, normal Cloud Shell or a standard GitHub-hosted runner.

## Private seeding

Copy [the synthetic seed](../config/demo-seed.example.json) to an ignored
`*.private.json`. Use **verified object IDs**, not UPNs, for owners and requesters.
One requester row maps to exactly one known budget. Do not upload private data
to GitHub. The function fails explicitly if metadata is unavailable/invalid.

```powershell
.\scripts\seed-demo.ps1 -SeedFile C:\private\legal.private.json -TableUrl https://<data>.table.core.windows.net -BlobUrl https://<data>.blob.core.windows.net -TenantId <tenant-id>
# After reviewing validation and the intended metadata replacement:
.\scripts\seed-demo.ps1 -SeedFile C:\private\legal.private.json -TableUrl https://<data>.table.core.windows.net -BlobUrl https://<data>.blob.core.windows.net -TenantId <tenant-id> -Apply
```

The script uses the explicitly signed-in Azure CLI operator. It conditionally
replaces the metadata blob and upserts only supplied index rows. It never clears
tables or changes requests. This cross-service seed operation is **not atomic**:
if an index write fails after metadata upload, fix permissions/data and rerun;
never claim the entire import rolled back. Removing mappings/owners requires
explicit, targeted operator changes; omission does not delete an index row.
Application deployments never seed or reset data.
Offline validation works locally; **Apply must run from a network with private
DNS and connectivity to these endpoints**, as well as the scoped Entra roles.
No network-connected seeding worker is created by this template. The admin-only
browser bootstrap described below instead executes writes inside the already
VNet-integrated Function. It still needs deployment, scoped roles and actual
private-network acceptance before the hosted form can be presented as ready.

## BU onboarding and browser bootstrap

The signed-in **Policy owner** setup area is independent of budget metadata, so
ordinary users can submit requests before their BU exists. Ownership comes from
the API-validated tenant/object ID, never a form field or email. Available token
name/contact claims are display hints only.

Before the first import, the ordinary workspace returns HTTP 409
`setup-required` when the metadata blob does not exist. Open **Policy owner**
to request a BU; a configured administrator can import the initial private JSON
under **Admin > Setup requests**, then refresh the workspace. Setup remains available without
that blob. Missing containers, denied access, invalid metadata and other storage
failures remain errors; they are not treated as an empty initialized workspace.
The setup endpoint sends literal JSON `null` before initialization, as does an
allocation read with no existing operation; an empty HTTP body is not the JSON
contract. Table lookups recognize the SDK's nested OData `ResourceNotFound` code
for absent entities. Missing tables and authorization failures still surface.

Configure a comma-separated `ONBOARDING_ADMIN_OBJECT_IDS` **repository secret**.
This is private authorization configuration, not a credential or Function key.
Use verified user object IDs, not UPNs. No other budget owner automatically
becomes an onboarding administrator. An empty allowlist disables administration.
Never put the real allowlist or seed file into public bootstrap JSON.

```powershell
'<verified-admin-object-id>' | gh secret set ONBOARDING_ADMIN_OBJECT_IDS --repo <owner>/<repo>
```

Deployment order for this feature:

Publish in two phases. First publish only the infrastructure templates,
Infrastructure workflow and deployment documentation. Hold application,
frontend and script changes until after steps 1-2: those paths trigger the
application workflow, which shares the infrastructure release concurrency
lock and can block Infrastructure while awaiting approval. Resolve any
obsolete waiting release before starting; do not approve older application
code as the onboarding release.

1. Run **Infrastructure** with `enable_bootstrap: true`. Review what-if in its
   logs, then approve apply. The secure administrator parameter is injected at
   execution, not stored in the artifact; a digest binds its value between
   what-if and apply. Restart if that secret changes during review. ARM what-if
   can omit Function app settings: after apply, verify the administrator
   allowlist is configured without printing it, the bootstrap flag is enabled
   and the onboarding table exists.
2. Separately review/apply the privileged `access.bicep` changes using the
   existing Function principal and current deployment outputs. Normal CI cannot
   grant these roles. Incremental deployment can leave the older narrow reader
   grants present; no role assignments are automatically deleted.
3. Publish the application, frontend, tests and script changes in the second
   phase. Deploy using the existing protected application workflow only after
   the infrastructure and access changes are ready.
4. Sign in as a configured administrator, open **Admin > Setup requests**, and use
   **Import private setup JSON**. Upload the ignored private seed from
   the local machine. The browser needs no direct storage connectivity.
5. Imports add budgets and create supplied requester mappings. Identical imports
   can be replayed; different existing budgets/mappings conflict instead of
   being overwritten. Requests are never cleared. Blob/Table writes are not
   atomic: a failed import can be partial; resolve the error and retry the same
   file. The older CLI seed tool intentionally retains its replacement/upsert
   semantics and is an in-network operator tool, not this additive endpoint.
6. Rerun Infrastructure with bootstrap disabled after confirming the import.
   This disables `POST /api/v1/onboarding/bootstrap`; ongoing onboarding keeps
   its narrowly scoped storage writes.

Fixed new-BU settings belong in private `metadata.onboardingDefaults`, not the
public frontend bootstrap:

```json
"onboardingDefaults": {
  "standard": { "perUserCredits": 10000, "billingMethods": "payAsYouGo" },
  "enhanced": { "perUserCredits": 25000, "billingMethods": "payAsYouGo" }
}
```

Deploy the updated API/frontend before importing this field. Reimporting the
updated private seed adds missing defaults without replacing existing budgets,
routing or requests, and does not modify existing Graph resources. Identical
defaults replay safely; different existing defaults conflict. A seed containing
only this configuration, the normal tenant/service/version fields, empty
`budgets` and empty `requesterIndex` is valid, so no existing BU is required.
If bootstrap was already disabled, temporarily enable it through the protected
infrastructure workflow, import/verify, then disable it again. New approvals
remain blocked until defaults are configured; do not substitute zero, unlimited
or a different billing mode for missing configuration.
After import, roll back only to an API version that understands
`onboardingDefaults`; older strict metadata schemas reject the new field.

For direct operator ARM deployment, supply the secure
`onboardingAdminObjectIds` parameter every time and set `bootstrapEnabled`
explicitly. Omitting the allowlist resets it to empty (fail closed). The
subscription wrapper forwards both settings. Neither deployment compiles the
private seed into the application.

### Approval, provisioning and allocation

- A request contains BU name, cost centre, monthly USD planning budget and a
  justification. It is persisted with a payload-bound operation ID and has no
  Graph side effects or ownership grant at submission.
- The user-selected demo convention is **USD 0.01 per Copilot Credit**, not an
  actual rated price or a purchase. Amounts are parsed as integer cents.
  **USD 20,000 = 2,000,000 credits**, initially **1,000,000 per policy**. A spare
  whole credit goes to Enhanced when the total is odd. The period is a UTC month.
- The administrator reviews fixed Standard/Enhanced per-user limits and billing
  methods from the private JSON; there is no BU-template selector or Graph
  template read. The API rechecks the reviewed values against configuration and
  freezes them into the approved operation. Later configuration changes do not
  rewrite already-approved/recovering operations. Unlimited remains explicit,
  not an error fallback. The administrator confirms manual
  cost-centre/cross-charge validation and enters an audit note before approval.
  There is no finance-system verification or automatic chargeback configuration.
- **Approve and provision** creates two non-mail-enabled security groups owned
  by the requester, without automatically enrolling the owner as a member. It
  creates two selected-groups policies with Cowork enabled, then assigns each
  group. Existing Legal resources are not edited.
- Graph policy creation supports delegated permissions only and currently
  requires Billing Administrator or Global Administrator at runtime. Use the
  approving admin's separate Graph connection and the existing consented scopes.
  The Function's system-assigned identity is for Azure storage, not Graph.
  Directory roles are assumed, never granted by this workflow.
- The browser verifies ownership, policy configuration, audiences and activation
  before attesting completion to the admin-only API. The Function validates
  authorization/state and publishes IDs and the immutable approved ceiling.
  It does not independently call Graph to verify that administrator attestation.
- ETags protect decisions/checkpoints; a five-minute operation lease coordinates
  sessions. Progress is durable. A lost create response is reconciled by its
  operation-specific resource marker/name; an unresolved create is never blindly
  repeated. Definitive Graph request rejection allows an explicit retry.
  Closing the browser pauses work; resume via the administrator queue after any
  outstanding lease expires. Partial resources are not automatically deleted.
- Owners see the new budget in their workspace, with Graph data loaded
  automatically using the cached/silently renewed delegated Graph token. Admin
  inventory and allocation reads behave the same way. Interactive authorization
  is offered only when silent acquisition requires it, never opened by an
  automatic read. Graph writes remain explicit. Permissions/roles are assumed, not inferred
  from group ownership. Request status refreshes while the onboarding page is
  open; there is no notification delivery.
- **Allocate the approved monthly budget** supports later tier changes, such as
  **1.5M Standard / 500K Enhanced** under the same 2M ceiling. Server-side plans
  recheck current ownership, resource mappings and the approved ceiling.
  Decreases happen before increases; actual Graph values and fresh policy ETags
  are checked. Failed allocations retain a resumable plan.
- This is an **application guardrail**, not a tenant-wide financial guarantee.
  Separately authorized admin-center/Graph changes can bypass it. API operations,
  metadata writes and Graph provisioning are not one atomic transaction.
- Onboarding does not change existing requester routing. New group membership
  and ordinary users' requester-index mappings remain explicit administration;
  BU ownership alone is sufficient to expose the owner workspace.

Before customer use, exercise ordinary-user submission, administrator approval,
actual Graph provisioning, owner refresh, allocation/reload and lost-response
recovery in the deployed environment. Unit tests use fake Graph responses; they
do not establish that tenant provisioning or private storage access works live.

## GitHub configuration

Create `demo-whatif`, `demo-infra` and `demo` environments. Restrict each to exact
`main` deployment branches, protect apply/release with required reviewers, and
disable untrusted PR deployments. OIDC subjects in `bootstrap.bicep` match these
exact names and the configured repository. Do not widen them to all branches.

The bootstrap template's `repositorySubjectPrefix` parameter must match GitHub's
actual `sub_claim_prefix`, not a guessed repository-name-only subject. Inspect it
with:

```powershell
gh api repos/<owner>/<repo>/actions/oidc/customization/sub
```

This repository enables immutable subjects, so its prefix includes the numeric
owner and repository IDs. The former bootstrap `repository` parameter is
replaced by the full `repositorySubjectPrefix`. Each existing federation record
appends only its exact environment. If GitHub's subject configuration changes,
review/apply the bootstrap template with an authorized operator before retrying
CI login; a workflow cannot repair its own rejected trust. Keep the issuer,
audience, identities and role assignments unchanged. Do not disable immutable
subjects, broaden trust or introduce a client secret to fix `AADSTS700213`.

Repository variables (public IDs/config used by the build):

| Variable | Source |
| --- | --- |
| `AZURE_TENANT_ID` | tenant |
| `SPA_CLIENT_ID` | existing SPA client ID |
| `API_CLIENT_ID` | API registration client ID |
| `API_BASE_URL` | `apiUrl`, no trailing slash |

Environment variables (in each relevant deployment environment):

| Variable | Environment |
| --- | --- |
| `AZURE_TENANT_ID`, `AZURE_SUBSCRIPTION_ID` | all three |
| `AZURE_RESOURCE_GROUP`, `AZURE_INFRA_CLIENT_ID` | whatif, infra |
| `SPA_CLIENT_ID`, `API_CLIENT_ID` | whatif |
| `AZURE_APP_CLIENT_ID`, `AZURE_FUNCTION_NAME`, `API_BASE_URL`, `SITE_URL` | demo |

In `demo`, set secret `SWA_DEPLOYMENT_TOKEN` from the selected site's deployment
token using a private operator session. Do not print it, store it in files or put
it in ARM outputs. Rotate/reset in SWA's deployment-token management, then replace
the environment secret. No Function keys, account keys, SAS, publish profiles or
Azure client-secret GitHub secrets.

Set the separate repository secret `ONBOARDING_ADMIN_OBJECT_IDS` as described
above before using the updated infrastructure workflow. The compiled templates,
parameter artifact and frontend never contain its actual value.

- **Validate:** credential-free tests/build, Bicep compilation and output checks.
- **Infrastructure:** manual main-only what-if; uploads the compiled template,
  parameters and private-configuration digest. Review what-if in the masked job
  logs; its raw output is not uploaded. A protected apply job consumes the **same**
  artifact. Review the summary before approval; restart the run if code/parameters
  or live state changed. ARM what-if can show service-default noise. After apply,
  actual private-network configuration is checked; policy-modified settings
  cannot silently be treated as the requested configuration.
- **Deploy application:** tests/builds once, configures only public bootstrap,
  packages production API dependencies, gates on `demo`, deploys API using OIDC
  then SWA's prebuilt static files. No SWA-managed API, previews or slots.
- All infrastructure/app runs share a non-cancelling concurrency group. Do not
  run operator deployments concurrently. Action versions are pinned to commit IDs.

Smoke checks are non-mutating: HTTPS, callback and missing-bearer rejection. A
minimal `/api/health` reports **process liveness only**, not storage readiness.

## Acceptance before sharing with customers

1. Inspect deployed identity/config and RBAC, disabled Shared Key/basic publishing,
   private containers, exact CORS and authenticated telemetry. Confirm there are
   no user-assigned runtime client-ID selectors.
   Confirm all three private endpoints are Approved, their DNS A records match
   NIC IPs, the private DNS VNet links are Completed and the Function retains
   the same principal ID with the correct delegated integration subnet.
   From the Function or an approved in-network worker, resolve storage names to
   those private IPs and exercise Entra-authenticated Blob/Table access. From
   outside the private network, verify storage data access remains blocked.
2. Submit as a separate requester with no owner Graph permissions; reload and
   confirm persistence. Missing mappings/ownerless budgets must fail visibly.
   Without balance-read access, verify that the balance error/explicit consent
   action does not block submission or erase its draft. With an authorized demo
   account, verify the balance uses that account's tenant-local object ID and
   displays real consumed/remaining credits, limits, timestamps and its effective
   policy. Zero, unlimited, unreported and unavailable must remain distinct.
   Policy-read denial must not hide a successful balance. This reading does not
   replace the server's requester-index/current-owner routing; no browser policy
   ID becomes authoritative, and no balance failure is replaced by demo numbers.
3. As owner, approve/reject; reload as requester to see the persisted decision.
   Verify unrelated users cannot see it, owner self-approval works, and competing
   decisions return conflicts rather than overwriting each other.
4. Confirm changed ownership/routing and expired temporary requests block
   decisions. Verify the captured exclusive UTC boundary.
5. With an explicitly authorized demo user, verify **Approve and move to
   Enhanced** adds/verifies Enhanced before removing Standard using `$ref`, and
   confirms final direct/transitive membership before recording completion.
   No policy limits, billing settings or directory user objects may be modified.
   Verify **Apply Enhanced upgrade** for a legacy approval, partial failure and
   uncertain-response retry without duplicate writes, an active/expired lease,
   changed owner/routing/group mappings, and inherited membership requiring
   administrator review. Loading/refreshing must not initiate membership writes.
   Check both permanent and temporary requests: temporary moves show manual
   return deadlines/overdue warnings, never promise automatic reversion, and
   expired requests cannot authorize further membership writes.
   Graph reads may lag behind verified membership changes. Rejection is API-only.
6. Confirm signed-out privacy and account-switch cleanup in the browser. Test
   storage denial as an unprivileged identity, without credential fallback.
7. Confirm Insights telemetry arrives with local ingestion auth disabled, and logs
   contain correlation IDs but neither tokens nor justifications.
8. Repeat unchanged infrastructure what-if and code deployment; verify requests,
   metadata/index and the Function principal ID are preserved.

Automated unit tests are not evidence that these live account-isolation and storage
RBAC acceptance steps passed. Use only explicitly approved demo users/data for
write acceptance checks.

An opt-in SDK/service integration suite uses a **dedicated** Entra-authorized
`sttest*` storage account: set `TEST_STORAGE_ACCOUNT` and run `npm test --prefix
api -- test/table.integration.test.ts`. It creates/uses a test `Requests` table,
exercises actual conditional create/update and list/read persistence, and removes
only its unique test record. It refuses production account names and is skipped
when no dedicated test account is configured. It does not grant permissions or
create an Azure account automatically.

## Local API development

Use Azure Functions Core Tools v4 and Node 24. Create ignored `api/local.settings.json`
with `IsEncrypted: false` and `Values`: `FUNCTIONS_WORKER_RUNTIME: node`,
`AZURE_FUNCTIONS_ENVIRONMENT: Development`, `DEMO_LOCAL_AUTH: azure-cli`,
`TENANT_ID`, `SPA_CLIENT_ID`, `API_AUDIENCE` (API client ID), `BUSINESS_TABLE_URL`,
`BUSINESS_BLOB_URL`. Use an authorized Azure CLI identity and dedicated development
data; do not point integration tests at the real request table. HTTP-only local
development needs no storage triggers. Allow only the localhost frontend origin
with `func start --cors http://localhost:5173` and configure its API URL to
`http://localhost:7071`. No production key credentials are used locally.
Local API access to private development storage also requires an approved
private network route and DNS; Entra roles alone do not provide connectivity.

## Operations and recovery

Failures include a correlation ID. Check API settings, current identity and
scoped RBAC, then metadata and routing. No successful response is returned when
storage access fails. JWTs and request text are not logged.

Rollback by redeploying an older **compatible, tested** application commit.
There are no Flex slots and no atomic API/frontend release guarantee. Use additive
schema changes; unsupported request versions fail explicitly. Do not roll back or
erase Tables with application code. Review infrastructure rollback with what-if.
The optional request `fulfilment` field preserves legacy records but older strict
readers cannot read newly upgraded records. Deploy the compatible API before
using the new UI; do not roll back to a version that rejects this field.
The API stores approval/progress and rechecks authorization using existing storage
roles; Graph remains delegated browser execution with no new application
permissions, proxy or background worker. A completed state is the authorized
browser's verification report, not an independent server-side Graph attestation.

LRS is redundancy, not backup or point-in-time restore. Take a private authorized
export before destructive data work. No automatic backup/restore or recovery SLA
is implemented. Teardown needs separate explicit approval identifying both the
resource group and the persisted data that would be lost.

References: [Flex Bicep](https://learn.microsoft.com/azure/azure-functions/functions-create-first-function-bicep),
[OIDC deployment](https://learn.microsoft.com/azure/azure-functions/functions-how-to-github-actions),
[Table authorization](https://learn.microsoft.com/azure/storage/tables/authorize-access-azure-active-directory).
