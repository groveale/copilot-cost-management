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
- **Policy owner** retains authorized budgets, Graph details, allowance approvals
  and budget allocation. Its setup area allows a first or additional BU request
  and shows the caller's setup history, even when the main workspace has not yet
  been initialized. It does not expose administrator review or bootstrap tools.
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
- Allowance-increase approval means **Approved - manual fulfilment required**.
  That flow has no Graph writes, notifications, automatic expiry/reversion or
  native Cowork actions. BU provisioning is the separate workflow below.
- On an uncertain network outcome, refresh history before making a new request.
  Retrying the unchanged draft reuses its operation ID within the current page;
  reloading/switching account intentionally clears the draft and operation ID.
- Requester identity never comes from form fields, UPNs or browser JSON.

Owners automatically load Graph policy configuration, assigned groups, direct
members and current user/service consumption for their selected budget. Admin
inventory and approved-budget allocation limits also load on entry. Reads use
MSAL's cached/silently renewed Graph token; refresh does not force a popup.
**Grant Graph access** appears only when silent authentication requires user
interaction, and only its explicit action may open a consent/sign-in popup.
API tokens are never reused as Graph tokens; requester balance reads use the
separate read-only scope subset described above. Provisioning, recovery and allocation writes require their
own explicit actions, not background effects. The existing eight
operator-selected delegated scopes are unchanged. Preview Graph APIs are for
evaluation, not supported production applications. Nested membership expansion,
historical spend and reconciled billing are not provided.

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
published after browser verification. Owners can then reallocate within the
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
