import { useState, type FormEvent } from "react";
import { budgetsForOwner, normalizeUsername, type BudgetConfig, type DemoConfig, type ExampleRequest } from "./config";
import { expiryLabel, requestExpiry, resolveRequestRoute, type RequestIdentity } from "./requests";

function ExampleRequestCard({ request, config, account }: {
  request: ExampleRequest; config: DemoConfig; account: RequestIdentity;
}) {
  const [decision, setDecision] = useState<"approval" | "rejection" | null>(null);
  const [reason, setReason] = useState("");
  const [error, setError] = useState("");
  const route = resolveRequestRoute(config, { tenantId: config.tenantId, username: request.requesterUsername });
  const expiresAt = requestExpiry(request.grantKind, new Date(request.requestedAt));
  const expired = expiresAt !== null && Date.parse(expiresAt) <= Date.now();
  const canDecide = route.state === "ready" && route.budget.id === request.budgetId &&
    route.approvers.includes(normalizeUsername(account.username)) && !expired;

  function decide(kind: "approval" | "rejection") {
    setError("");
    if (!canDecide || (expiresAt !== null && Date.parse(expiresAt) <= Date.now())) {
      setError("This example is not currently routed to you or its requested billing period has ended.");
    } else if (kind === "rejection" && (reason.trim().length < 10 || reason.trim().length > 1000)) {
      setError("Add a rejection reason of 10-1000 characters.");
    } else {
      setDecision(kind);
    }
  }

  function reject(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    decide("rejection");
  }

  return <article className="example-request">
    <div className="section-heading"><h3>{request.id}</h3><span className="badge">{decision ? `${decision} preview` : expired ? "Past billing period / example" : "Pending owner decision / example"}</span></div>
    <dl className="metadata-grid">
      <div><dt>Requester / synthetic example</dt><dd>{request.requesterUsername}</dd></div>
      <div><dt>Requested tier</dt><dd>Enhanced</dd></div>
      <div><dt>Duration</dt><dd>{expiryLabel(expiresAt)}</dd></div>
      <div><dt>Created</dt><dd>{new Date(request.requestedAt).toLocaleString()}</dd></div>
    </dl>
    <p className="justification">{request.justification}</p>
    <p className="small muted">Routing: {route.state === "ready" ? route.approvers.join(", ") : route.reason}</p>
    {route.state === "ready" && route.budget.id !== request.budgetId &&
      <p className="reading-note">The requester index no longer matches this example's budget. Administrative rerouting is required.</p>}
    {expired && <p className="reading-note">The captured temporary period has ended. This example cannot be approved as a current-period request.</p>}
    {decision ? <div className="reading-note" role="status">
      <strong>{decision === "approval" ? "Approval preview - fulfilment would be a separate step." : "Rejection preview."}</strong>
      {decision === "rejection" && <p>{reason.trim()}</p>}
      <p>Simulation only: no decision was saved, no notification sent, and no Graph change or expiry job created. Reload to reset the example.</p>
    </div> : <form className="request-form" onSubmit={reject}>
      <label htmlFor={`${request.id}-reason`}>Reason if rejecting</label>
      <textarea id={`${request.id}-reason`} value={reason} rows={2} minLength={10} maxLength={1000} required
        disabled={!canDecide} onChange={(event) => setReason(event.target.value)} />
      {error && <p className="error" role="alert">{error}</p>}
      <div className="actions">
        <button type="button" disabled={!canDecide} onClick={() => decide("approval")}>Preview approval</button>
        <button className="secondary" type="submit" disabled={!canDecide}>Preview rejection</button>
      </div>
      {!canDecide && !expired && <p className="reading-note">You are not an eligible approver for this example. Self-approval and uncertain routing are not permitted.</p>}
    </form>}
  </article>;
}

export function ExampleRequests({ config, budget, account }: {
  config: DemoConfig; budget: BudgetConfig; account: RequestIdentity;
}) {
  if (!budgetsForOwner(config, account).some((item) => item.id === budget.id)) return null;
  const requests = (config.exampleRequests ?? []).filter((request) => request.budgetId === budget.id);
  return <section className="card request-inbox" aria-labelledby="inbox-heading">
    <div className="section-heading"><div><span className="eyebrow">APPROVAL INBOX / JSON EXAMPLES</span><h2 id="inbox-heading">What requests could look like</h2></div>
      <a className="button-link" href={`${import.meta.env.BASE_URL}?view=request`}>Open requester form</a></div>
    <p className="muted">These are synthetic examples from JSON, not live submissions. Previewing a decision only changes this screen until reload; it does not change policy limits, group membership or credits.</p>
    {requests.length === 0 && <p>No example requests are configured for this budget.</p>}
    {requests.map((request) => <ExampleRequestCard key={request.id} request={request} config={config} account={account} />)}
  </section>;
}
