# Copilot budget demo

React/TypeScript frontend for a single-tenant, evaluation-only Copilot budget
portal. See [deployment](../docs/deployment.md) for the Azure and GitHub setup.

## Run locally

Use Node 24:

```powershell
npm ci
npm test
npm run build
npm run dev
```

Configure the public bootstrap with your tenant/client IDs, API scope and API
origin. The committed bootstrap is a non-working synthetic example; incomplete
hosted configuration fails visibly rather than showing a pretend submission form.
Never put private seeds, business metadata, tokens or credentials in `public`.
Retain `http://localhost:5173/auth.html` as an Entra SPA redirect. The dedicated
MSAL v5 redirect bridge and both HTML build entries are required.

## Hosted behaviour

- Nothing beyond sign-in is rendered before login. After login the API returns
  the caller's authorized context using stable tenant/user object IDs.
- Persistent **Admin**, **Policy owner** and **Request credits** switches share
  the signed-in identity. The view is reflected in `?view=admin`, `?view=owner`
  or the existing `?view=request`; browser back/forward restores the view.
  Switching views clears that view's draft and operation ID.
- A **Dark mode / Light mode** button is available on the sign-in screen and
  portal header. It switches the existing theme immediately and retains the
  selected theme while navigating views or signing in/out on the current page.
- **Admin** is shown only after the API verifies administrator access. Direct
  admin URLs do not grant access. **Setup requests** contains the review queue
  and the explicitly enabled private JSON importer. **Spending policies**
  automatically reads every spending-policy page directly through delegated Graph,
  not just configured budgets. The inventory is read-only. Business unit,
  cost centre, tier and owner mappings come from private metadata; missing
  mappings are explicitly **Unmapped**, never inferred from policy names or
  Entra group ownership. Owner names are resolved through Graph; lookup failures
  remain visible with the mapped object IDs. API/Graph failures are not replaced
  by an empty inventory. Before initial import, the list explicitly warns that
  portal ownership has not been configured.
- **Policy owner** opens on a clean **Business units** list showing only the
  caller's server-authorized BUs, their cost centres and planning budgets.
  **View policies** opens one BU's Standard/Enhanced policies, allocation,
  members and consumption, with a **Back to business units** action. No Graph
  reads are made for the list, unopened BUs, or simply opening Credit requests.
- The separate **Credit requests** owner page contains the approval inbox across
  all owned BUs, with friendly BU names on requests. The **Set up a BU** page holds
  the new-BU form and setup history; it remains available before budget metadata
  has been initialized. These pages do not expose admin review/bootstrap tools.
  URLs use `?view=owner&ownerPage=budgets`, `ownerPage=requests` or
  `ownerPage=setup`, with `&budget=<budget-id>` for an opened BU. Refresh and browser
  back/forward restore the selected screen; an unavailable/unowned BU is an
  explicit error, never substituted with another BU. Switching owner pages
  clears that page's draft and stops its in-flight reads/actions; partial
  provisioning/upgrade work remains persisted for explicit recovery.
- **Request credits** automatically reads the signed-in user's current Graph
  service balance and effective policy, alongside their allowance request and
  history. It shows consumed/remaining credits, configured limits, effective
  capacity, access status and source/fetch timestamps. These are current credit
  readings, not invoiced spend or a verified billing-period history.
  Missing balances are unavailable, not zero; policy-read errors preserve the
  balance and remain visible. BU setup requests go to administrators; credit
  increases go to budget owners.
- Initial requester sign-in requests only OIDC and the application's delegated
  API scope. The optional balance panel separately acquires only the existing
  `CopilotCostManagement-UserData.Read.All` and
  `CopilotCostManagement-Policy.Read.All` Graph read scopes, never the owner's
  directory/assignment/write scopes. The demo assumes the signed-in account has
  permission and any required directory role to read this data; it does not
  establish ordinary-user self-service support or grant those privileges.
  An explicit **Grant balance read access** button appears only for an MSAL
  interaction-required error. Permission/network failures are shown without
  simulated data, and do not block a server-routable request or erase its draft.
  The form submits business justification (20-2000 characters), permanent or
  temporary duration, and a client-generated operation ID. Routing is server-side.
  The browser's reported effective policy is informational, not routing authority:
  the API still uses the private requester index and current budget owners.
  Policy-to-budget/owner metadata remains necessary; no owner is inferred from
  a policy creator or Entra group owner.
- Temporary requests capture the next UTC calendar month's first instant,
  exclusive. This is a **demo convention**, not a verified tenant billing period.
- Requests and decisions are persisted in Azure Tables. Requesters see only
  their own history; current owners see requests for their budgets.
- Decisions use conditional ETags and recheck current ownership, routing,
  expiry. Owners **may approve their own requests in this MVP**; ownership is
  still required. Rejection requires 10-1000 characters.
- **Approve and move to Enhanced** first obtains the owner's delegated Graph
  access and records approval, then applies the mapped group move. It verifies
  both enabled policies, their single mapped group and enabled service, and that
  groups are assigned/cloud-managed/non-role-assignable security groups.
  Enhanced direct membership is added and verified **before** Standard direct
  membership is removed using its `$ref`. Final direct and transitive checks
  must confirm Enhanced without Standard before **Enhanced move applied** is
  recorded. No policy caps, billing settings or user objects are changed.
- Both permanent and temporary approvals apply this move. Temporary requests
  retain their recorded UTC boundary, but return-to-Standard workflows and
  overdue reminders are out of scope for this MVP. Group membership does not
  automatically revert when that timestamp passes. Expiry still prevents new
  late approvals/upgrades; it is not an access-expiry mechanism.
  There is no scheduler, automatic reversion, notification or native Cowork action.
- Credit requests start collapsed in both the owner inbox and requester history,
  showing status, BU, submitted date and duration. **Expand / Collapse** reveals
  IDs, justification, audit details and owner approval/retry actions. Collapsing
  retains the rejection draft and does not cancel an in-progress operation;
  progress/errors remain visible outside the collapsed details.
  Status badges use the same colours as BU requests: orange for pending or
  unfinished upgrades, green only after the Enhanced move is confirmed, and red
  for rejected requests. Approved-but-not-applied, in-progress, needs-attention
  and expired incomplete requests remain orange with their explanatory text.
  Applied temporary upgrades remain green after the requested period; that
  timestamp does not automatically revert membership.
- Request progress is durable in the existing Requests table. Five-minute leases
  and conditional ETags coordinate attempts; before each membership mutation the
  API rechecks current ownership, requester routing, frozen policy/group mappings
  and temporary expiry. The trusted delegated browser reports verification;
  the Function does not proxy Graph or independently attest to Graph membership.
  No new Graph scopes, Graph application permissions or storage roles are added.
- Existing approved requests expose **Apply Enhanced upgrade**; loading the page
  never applies them automatically. Partial failures retain approval and expose
  **Retry Enhanced upgrade**. Retry rereads real memberships and skips confirmed
  writes, including after uncertain responses. Inherited tier membership blocks
  automatic reconciliation. If a browser closes mid-attempt, wait for the lease
  to expire, refresh and retry. Do not resubmit an already-approved request.
- On an uncertain network outcome, refresh history before making a new request.
  Retrying the unchanged draft reuses its operation ID within the current page;
  reloading/switching account intentionally clears the draft and operation ID.
- Requester identity never comes from form fields, UPNs or browser JSON.

Owners automatically load Graph policy configuration, assigned groups, direct
members and current user/service consumption **after opening a BU**. Returning
to a BU after an approval reads fresh Graph data; no policy details are loaded
alongside the request inbox. Admin inventory also loads on entry. Policy cards
show the monthly allocation and per-user cap; settings/groups, member lists and
each member's detailed consumption start collapsed. Reads use
MSAL's cached/silently renewed Graph token; refresh does not force a popup.
**Grant Graph access** appears only when an automatic read requires user
interaction. Explicit provisioning, upgrade and allocation actions also preflight
Graph access silently and may open a popup only when interaction is required.
API tokens are never reused as Graph tokens; requester balance reads use the
separate read-only scope subset described above. Provisioning, upgrades, recovery and allocation writes require their
own explicit actions, not background effects. The existing eight
operator-selected delegated scopes are unchanged. Preview Graph APIs are for
evaluation, not supported production applications. Nested membership expansion,
historical spend and reconciled billing are not provided.

Each policy has an **Edit** button opening numeric **Monthly allocation** and
**Per-user credit limit** fields, not a slider. The editor reads fresh limits only
when opened; **Save policy limits** creates a durable, resumable allocation plan.
The other policy's current limits are preserved, and the combined monthly limits
cannot exceed the BU's stored approved ceiling. Reduce one policy before increasing
the other when reallocating a fully allocated budget. Per-user limits can be zero
or explicitly Unlimited; per-user/billing settings on the other policy are not
changed. Billing methods and onboarding defaults are not edited. Budgets without
an approved ceiling show why limit editing is unavailable. Stale selected-policy
values block saving; failed/partial edits expose **Resume saved policy edit**.

**Add member**, beside **Members and consumption**, looks up an existing enabled
directory user by their exact sign-in email (UPN) and requires an explicit
confirmation. It does not create accounts or move people between tiers. Direct
or inherited membership of the other tier blocks the add, as do unsafe groups,
changed policy assignments, and conflicting private BU routing. The browser uses
delegated Graph to add only the mapped group's member `$ref`, then verifies
direct/transitive membership. The API rechecks current ownership and mappings
before it additively saves that user's requester routing; another BU's mapping is
never overwritten. Existing direct membership is reconciled on retry, including
after a lost response or a failed routing save. Graph membership and private
routing are not one transaction; errors explicitly require retry/reconciliation.
The Function trusts the authorized browser's verification report and does not
proxy or independently attest to Graph. Successful edits/adds refresh the policy
cards and show confirmation; balance/policy readings may lag. No Graph writes
occur just from opening cards or disclosures.

Graph reads remain serialized with bounded throttling retries, pagination and
independent resource errors. Zero, unavailable and unlimited quantities remain
distinct. Consumption is not inferred from configured limits. Graph tokens and
API tokens are acquired separately and are never displayed or exported.

## Request a new BU

All signed-in users can submit a BU name, cost centre, monthly USD budget and
justification without Graph permissions. The API derives proposed ownership
from the validated object ID. A separate configured administrator reviews
cost-centre/cross-charge information and the fixed policy defaults, then approves
or rejects. There is no existing-BU template picker or dependency on Legal.
The BU request uses labels above full-width controls, a separate credit-allocation
summary and a full-width justification field. Name/cost centre sit side by side
on wider screens and stack on mobile; submission fields and validation are unchanged.
BU request cards start collapsed in both the owner's setup history and the admin
queue. The BU name, cost centre, monthly budget, submitted date and status badge
remain visible: orange **Awaiting approval** / **Provisioning**, green **Completed**,
red **Rejected**. Status text accompanies the colour; approval alone is not
completion. Expand to review details or use admin approval/recovery controls.
Collapsing preserves review drafts and in-flight work; recorded errors remain visible.

Fixed defaults live in private `metadata.onboardingDefaults` JSON:
Standard has **10,000 credits per person** and Enhanced **25,000**, both
`payAsYouGo` in the supplied seed example. These are per-person caps, not the
pooled BU policy budgets. No limits/billing defaults are silently invented.
Import the updated private seed through **Admin > Setup requests > Import private
setup JSON** after deploying this version. Reimporting adds missing defaults
without resetting requests or changing existing budgets/Graph policies.
A configuration-only seed with empty `budgets` and `requesterIndex` is also
supported. Different existing defaults conflict rather than being overwritten.
New approvals verify the reviewed settings against private configuration and
freeze them into the operation; older approved/recovering operations keep their
recorded settings. Missing defaults block approval but not request/rejection.

Approval provisions two security groups and two Cowork policies using the
administrator's delegated Graph connection. Groups are owned by the requester
but initially have no members. The demo's **USD 0.01/Credit** planning convention
makes **USD 20K = 2M credits**, initially **1M per policy**; this is not rated
spend or a purchase. Per-user caps/billing methods are reviewed independently.

Progress and IDs are persisted; partial or uncertain outcomes are not reported
as complete. Refresh/resume rather than submitting duplicates. New ownership is
published after browser verification. Owners can then edit each policy within the
stored approved ceiling, decreasing one tier before increasing the other.
Separately authorized admin-center changes can bypass this application ceiling.
Existing Legal resources, requester routing and allowance approvals are preserved.

See [deployment](../docs/deployment.md#bu-onboarding-and-browser-bootstrap) for the
private administrator allowlist, scoped runtime writes and explicitly enabled
initial JSON upload. No private seed is bundled or committed.

The earlier JSON-only `BudgetPortal`, `RequestPage` and example components remain
as isolated legacy test fixtures, not production entry points or live inboxes.
The hosted entry point is `HostedPortal`; synthetic requests are not seeded into
the real request store.
