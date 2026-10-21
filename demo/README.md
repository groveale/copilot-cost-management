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
- `?view=request` requests only OIDC and the application's delegated API scope.
  The form submits business justification (20-2000 characters), permanent or
  temporary duration, and a client-generated operation ID. Routing is server-side.
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

Owners can separately connect Graph to inspect policy configuration, assigned
groups, direct members and current user/service consumption. The existing eight
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
cost-centre/cross-charge information, loads existing tier templates, then
approves or rejects.

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
