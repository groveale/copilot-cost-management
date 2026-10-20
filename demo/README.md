# Copilot Budget Portal demo

Browser-only, operator-led evaluation over Microsoft Graph. All non-Graph business
data belongs in [public/demo-config.json](./public/demo-config.json). No custom API,
database, metadata service, application-data storage, downloads or usage history.

## Live budget details

The demo shows actual Graph data rather than a pass/fail checklist:

- Standard and Enhanced spending-policy configuration, limits and billing methods.
- Per-policy service settings and actual assigned groups.
- Directory group configuration and all direct members, including non-user types.
- Each direct user's current Cowork consumption, remaining credits, access decision,
  reported policy, source timestamp and fetched timestamp.

Business metadata and owner labels come from JSON and are labelled separately.
The view makes **no membership or policy changes**. The request form and example
approval inbox below are a separate, explicitly simulated experience.

## Cowork request example

Use this link for the requester-facing page:

```text
http://localhost:5173/?view=request
```

For a customer-accessible demo, use the corresponding URL on an approved HTTPS
static host and register that host's authentication callback. A Cowork message can
direct the user to the link; this change does not install or assume a native Cowork
action, callback or supported request-submission API.

The requester signs in before seeing the form. This path requests **only OIDC
identity scopes (`openid`, `profile`)**, not the owner's privileged Graph scopes.
It never calls policy, group, membership or balance APIs. An existing signed-in
account can be reused without acquiring a Graph token. The enterprise application's
sign-in/assignment policies must permit the chosen requester account; preview API
roles are needed for owner operations, not for filling in this form.

The form collects:

- Business justification (20-2000 characters).
- Permanent Enhanced allowance, or Enhanced for the rest of the current billing
  period. This demo uses the UTC calendar month and captures the next month's
  first instant as the exclusive expiry when previewing the request.

Selecting **Preview request and routing** does not submit, save, notify, approve,
create an owner-inbox entry, schedule an expiry or modify Graph. The preview exists
only until the form changes, the account changes or the page reloads. Permanent
means an ongoing entitlement subject to review, not a credit purchase or a limit
change. The UI does not claim to know the requester's current balance or tier.

### Routing index: user -> budget -> current owners

The same static JSON contains `requesterIndex` entries:

```json
{
  "username": "person@example.com",
  "budgetId": "legal"
}
```

Routing derives the requester from the signed-in identity, never from a query
parameter or an editable requester/policy/approver field. The index identifies the
budget rather than a fixed source policy, so a future Standard/Enhanced movement
does not require choosing a different business owner. Owners are resolved from
the budget's current `owners` array. Missing mappings, multiple possible budgets,
missing owners, and self-approval with no other owner all show **Needs routing**;
the app does not guess or claim delivery.

The Legal index includes the configured administrator and two synthetic `.test`
users. Replace or extend the synthetic entries with actual test-account UPNs to
exercise sign-in as a separate requester. If you use the administrator account
for its own request, it is currently Legal's only owner, so the preview correctly
requires another owner or administrative triage rather than self-approval.

The index is manually maintained demo metadata, **not evidence of current group
membership or the API-selected policy**. In a real implementation a privileged
owner/operator process could reconcile directory membership and validated active
policy observations into a timestamped index. The end user would still need no
balance API rights. Revalidate eligibility, current ownership and the requested
period before a real decision or fulfilment; do not execute from a stale index.
Use stable tenant/user object IDs rather than mutable UPNs for a production store.

### Owner inbox examples

An authenticated Legal owner sees two synthetic examples from `exampleRequests`
in the JSON: one permanent and one for the captured remainder of a billing period.
Example records contain the requester, budget reference, justification, grant kind
and UTC creation time; approvers are resolved from current budget ownership rather
than stored as a stale copy.

**Preview approval** and **Preview rejection** change only the displayed example
until reload. Rejection requires a reason. Self-approval, unresolved routing and
expired temporary requests cannot be preview-approved. No decision is persisted,
no group movement happens, no notification is sent, and no automatic reversion is
scheduled. Approval is explicitly separate from fulfilment.

The requester form and these seeded examples are deliberately **not connected by
a shared submission store**. A static browser app cannot write new requests back
to the hosted JSON. Real multi-user request delivery and an authoritative approval
queue require an agreed shared store/integration with identity validation and
server-enforced access. That remains outside this visual-demo scope.

## Owner sign-in experience

Before sign-in, only the login screen is rendered: no budget labels, owner
addresses, policy/group IDs, connection details or Graph data are shown. No Graph
budget reads run while signed out.

After sign-in (including a restored MSAL session), the app looks up the account's
`username` in each budget's `owners` array. Matching is exact after trimming and
ignoring case, and the account must belong to the configured tenant. Only matched
budgets appear in the workspace and selector, and only their Graph data is loaded.
An unmatched account sees **No budgets assigned** without Graph budget requests.
Sign-out and account switching clear the previous details and cancel pending reads.

For Legal, the configured owner is
`admin@m365cpi77517573.onmicrosoft.com`. Use the account's actual sign-in username
(UPN), not a display name or an assumed email alias. Co-owners can be added as
additional usernames in the same array. No Entra group ownership is changed.

This deliberately **simulates a database lookup**. The JSON remains a publicly
served static asset, and client-side filtering is not secure authorization.
Graph still enforces the signed-in account's actual privileges. A future secure
deployment needs server-enforced owner access, not just this UI behavior.

## Run

Requires Node.js 24 (or another version supported by Vite 8 and Vitest 5).

```powershell
Set-Location C:\Users\alexgrover\source\repos\copilot-cost-management\demo
npm install
npm run dev
```

Open `http://localhost:5173`. Use `npm test` for unit tests and `npm run build` for
type-checking and a static build. `npm run preview` serves the production build on
the same port after stopping the development server. Never open index.html as a
file URL. No application backend or Vite proxy is configured.

## Entra setup

The JSON currently reuses the tenant/client IDs from the existing PowerShell
scripts, as agreed. No credentials are embedded. In the selected app registration:

1. Under **Authentication**, add a **Single-page application** platform redirect:
   `http://localhost:5173/auth.html`. Do not register it only as a Web or
   mobile/desktop redirect. The dedicated callback runs the MSAL redirect bridge.
2. This evaluation explicitly requests the operator's existing Microsoft Graph
   **delegated** scope set below. Ensure administrator consent covers these exact
   scopes on this application:
   - `CopilotCostManagement-Assignment.Read.All`
   - `CopilotCostManagement-Assignment.ReadWrite.All`
   - `CopilotCostManagement-Policy.Read.All`
   - `CopilotCostManagement-Policy.ReadWrite.All`
   - `CopilotCostManagement-UserData.Read.All`
   - `Directory.AccessAsUser.All`
   - `Directory.Read.All`
   - `Directory.ReadWrite.All`
   The interactive and silent token requests use the same list in
   [src/auth.ts](./src/auth.ts), which also drives the UI permission display.
   The `Directory.*` scopes are broad, not a least-privilege deployment
   recommendation. They replace the earlier proposed `GroupMember.ReadWrite.All`
   and `User.ReadBasic.All` requests for this evaluation. Requested write scopes
   do not add live write actions to the detail view.
3. The operator must have the endpoint-specific active directory roles, any
   required enterprise-app assignment and tenant preview access. Consent alone
   does not provide these. Graph Explorer and PowerShell consent do not grant
   these permissions to this SPA.
4. Do not add a client secret or application permissions.

Allow the sign-in popup. If silent token acquisition later requires interaction,
use **Sign in / renew consent** again; failures are not silently ignored.

## Configure existing test resources

The JSON contains the operator-supplied Legal mappings and illustrative owner.
An empty `budgets` array permits sign-in but displays a configuration message
instead of inventing resources.

Add a budget using this shape, replacing all placeholders with designated
evaluation resources:

```json
{
  "id": "finance-demo",
  "displayName": "Finance",
  "businessArea": "Finance",
  "costCenter": "cc0001",
  "owners": ["owner@example.com"],
  "plannedUsd": 20000,
  "periodLabel": "Customer demonstration",
  "tiers": {
    "standard": {
      "policyId": "<existing-standard-policy-guid>",
      "groupId": "<existing-standard-group-guid>"
    },
    "enhanced": {
      "policyId": "<existing-enhanced-policy-guid>",
      "groupId": "<existing-enhanced-group-guid>"
    }
  }
}
```

For the intended tier workflow, policies should be enabled for `cowork` and
assigned to their designated dedicated groups. The read-only view displays the
actual status and assignments without hiding a disabled or differently configured
policy. Every group returned by the assignment endpoint is loaded, along with the
JSON-mapped group. An unverified or absent mapping is explicitly labelled, never
presented as an actual Graph assignment.

Direct members of every displayed group are listed; nested groups are not expanded.
Only direct users have service balances queried. Users appearing in multiple
groups have their balance fetched once per refresh. Resource IDs in the JSON must
be distinct across tiers and budgets. Cost centre remains text.

JSON is read-only at runtime: edit it manually and reload/redeploy. It is served
as a public static asset to anyone with access to the host. Use non-sensitive
demo metadata and an approved hosting audience. Owner entries filter this demo's
workspace but do not restrict Graph permissions or secure the underlying static
assets. Planned USD is not a credit conversion or a cap.

## Using the view

Sign in to look up and load an owner-assigned budget automatically. Use **Refresh Graph details**
for a new reading, or **Cancel loading** to stop outstanding/queued requests.
Each section loads independently: a policy/service/assignment error does not hide
directory data, and a failed user's balance does not hide other users' readings.
Errors retain Graph codes and request IDs without exposing tokens. A failed
membership read is incomplete coverage, not a group with zero members.

The view reads only the configured policies; it does not require a tenant-wide
policy-inventory request. Configured groups can still show live directory data if
the assignment route fails, but their binding remains explicitly unverified.
Disabled policies and group metadata are information, not up-front pass/fail gates.
Expandable sections show the returned policy, group, service and user-balance JSON.

User readings are current observations, not historical budget spend. Missing,
unlimited, invalid and unsupported quantities are distinguished from zero. Policy
mismatches and overlapping tiers remain visible alongside the actual reading;
they are not silently attributed to a tier. No combined historical spend total is
calculated. No automatic polling or proxy/sample-data fallback exists.

No application data is saved or exported. Refresh replaces the transient current
view. Membership mutations and policy-propagation verification are not enabled by
this read-only detail screen.

## Design context

The [enterprise specification](../SPEC.md) and [design report](../DESIGN-REPORT.html)
remain future reference. This MVP does not implement their secure owner isolation,
durable workflows, provisioning or consumption ledger.

Microsoft documents the [cost-management API](https://learn.microsoft.com/en-us/microsoft-365/copilot/extensibility/api/admin-settings/cost-management/overview)
as beta and unsupported for production applications.
