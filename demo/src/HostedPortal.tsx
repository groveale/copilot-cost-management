import { useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import { DemoAuth } from "./auth";
import type { BudgetConfig, DemoConfig, GrantKind } from "./config";
import { RequestApi, type Context, type Page, type RequestRecord, type Submission } from "./api";
import { GraphClient } from "./graph";
import { loadBudgetDetails, type BudgetSnapshot } from "./details";
import { BudgetDetails } from "./BudgetDetails";
import { BusinessUnitOnboarding, BudgetAllocation } from "./Onboarding";

const message = (error: unknown) => error instanceof Error ? error.message : String(error);
const statusLabel = (record: RequestRecord) => record.status === "approved" ? "Approved - manual fulfilment required" :
  record.status === "rejected" ? "Rejected" : record.expiresAt && Date.parse(record.expiresAt) <= Date.now() ? "Expired / not approved" : "Pending owner decision";

export function HostedPortal({ config, auth, requestPage }: { config: DemoConfig; auth: DemoAuth; requestPage: boolean }) {
  const [account, setAccount] = useState(auth.account());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [context, setContext] = useState<Context | null>(null);
  const [refresh, setRefresh] = useState(0);
  const [onboardingOpen, setOnboardingOpen] = useState(false);
  const api = useMemo(() => new RequestApi(config.apiBaseUrl ?? "", () => auth.apiToken(account)), [auth, config.apiBaseUrl, account]);
  useEffect(() => {
    setContext(null);
    if (!account || busy) return;
    const controller = new AbortController();
    void api.context(controller.signal).then((value) => {
      if (!controller.signal.aborted) setContext(value);
    }).catch((error: unknown) => { if (!controller.signal.aborted) setError(message(error)); });
    return () => controller.abort();
  }, [account, api, busy, refresh]);

  async function authenticate(signOut = false) {
    setOnboardingOpen(false);
    setAccount(null);
    setContext(null);
    setError("");
    setBusy(true);
    try {
      if (signOut) await auth.signOut();
      else setAccount(await auth.signIn("requester"));
    } catch (error) { setError(message(error)); }
    finally { setBusy(false); }
  }
  if (!account || busy) return <main className="login-shell"><section className="card login-card">
    <span className="eyebrow">COPILOT BUDGET PORTAL</span>
    <h1>{requestPage ? "Request a higher allowance" : "Sign in to your budgets"}</h1>
    <p className="muted">Use your work account. Requests do not require Graph permissions.</p>
    <button disabled={busy} onClick={() => void authenticate()}>{busy ? "Completing sign-in / sign-out..." : "Sign in with Microsoft"}</button>
    {error && <p className="error" role="alert">{error}</p>}
  </section></main>;

  return <main className="shell">
    <header className="masthead"><div><span className="eyebrow">COPILOT BUDGET PORTAL</span>
      <h1>{requestPage ? "Request an Enhanced allowance" : "Your owner workspace"}</h1>
      <p className="muted">Shared requests. Current owners. Manual fulfilment.</p></div>
      <div className="operator"><p>{account.username}</p><div className="actions">
        <button className="secondary" onClick={() => void authenticate()}>Switch account / renew sign-in</button>
        <button className="secondary" onClick={() => void authenticate(true)}>Sign out</button>
      </div></div></header>
    <nav className="actions"><a href="?view=request">Request an increase</a><a href="?">Owner workspace</a>
      <button className="secondary" onClick={() => setOnboardingOpen((value) => !value)}>{onboardingOpen ? "Back to workspace" : "Request a new BU"}</button>
      <button className="secondary" onClick={() => { setError(""); setRefresh((n) => n + 1); }}>Refresh workspace</button></nav>
    {onboardingOpen && <BusinessUnitOnboarding key={`${account.homeAccountId}:${account.localAccountId}`} api={api} auth={auth} account={account}
      onPublished={() => { setError(""); setRefresh((value) => value + 1); }} />}
    {!onboardingOpen && error && <p className="error" role="alert">{error}</p>}
    {!onboardingOpen && !context && !error && <p role="status">Loading your authorized workspace...</p>}
    {!onboardingOpen && context && <>
      <div className="notice">Approval records a decision only. It does not change a policy, allowance or group membership.
        Owners may approve their own requests in this MVP. Temporary increases use the UTC calendar month for this demo; no automatic reversion is scheduled.</div>
      {requestPage ? <SubmissionForm api={api} context={context} /> : <>
        {!context.ownedBudgets.length ? <section className="card"><h2>No budgets assigned</h2><p>Contact the administrator if you should be an owner. No Graph data has been requested.</p></section> :
          <><RequestList api={api} inbox /><OwnerGraph budgets={context.ownedBudgets} serviceId={context.serviceId} auth={auth} account={account} api={api} /></>}
      </>}
    </>}
    <footer>Evaluation only. Requests are stored privately. Graph preview APIs are not supported for production applications.</footer>
  </main>;
}

function SubmissionForm({ api, context }: { api: RequestApi; context: Context }) {
  const [justification, setJustification] = useState("");
  const [grantKind, setGrantKind] = useState<GrantKind>("billingPeriod");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [saved, setSaved] = useState<RequestRecord | null>(null);
  const [refresh, setRefresh] = useState(0);
  const operation = useRef<Submission | null>(null);
  const controller = useRef<AbortController | null>(null);
  useEffect(() => () => controller.current?.abort(), []);
  async function submit(event: FormEvent) {
    event.preventDefault();
    setBusy(true); setError(""); setSaved(null);
    const current = new AbortController();
    controller.current = current;
    const text = justification.trim();
    if (!operation.current || operation.current.justification !== text || operation.current.grantKind !== grantKind) {
      operation.current = { operationId: crypto.randomUUID(), justification: text, grantKind };
    }
    try {
      const record = await api.submit(operation.current, current.signal);
      if (!current.signal.aborted) {
        setSaved(record); setJustification(""); operation.current = null; setRefresh((n) => n + 1);
      }
    } catch (error) { if (!current.signal.aborted) setError(`Request not confirmed. Your draft is retained. ${message(error)}`); }
    finally { if (!current.signal.aborted) setBusy(false); }
  }
  return <>
    <section className="card"><h2>Business need</h2>
      <p>{context.route.message}{context.route.budgetName && ` Budget: ${context.route.budgetName}.`}</p>
      <p className="small muted">Routing uses your verified tenant/user ID and the private requester index. No policy selection or balance API access is needed.</p>
      <form className="request-form" onSubmit={(event) => void submit(event)}>
        <label htmlFor="justification">Business justification (20-2000 characters)</label>
        <textarea id="justification" minLength={20} maxLength={2000} required rows={6} value={justification}
          disabled={busy} onChange={(event) => setJustification(event.target.value)} />
        <fieldset disabled={busy}><legend>How long is the increase needed?</legend>
          <label className="radio-choice"><input type="radio" name="duration" checked={grantKind === "billingPeriod"} onChange={() => setGrantKind("billingPeriod")} />Rest of this billing period (demo: UTC calendar month)</label>
          <label className="radio-choice"><input type="radio" name="duration" checked={grantKind === "permanent"} onChange={() => setGrantKind("permanent")} />Permanent increase</label>
        </fieldset>
        <button disabled={busy || context.route.state !== "ready"}>{busy ? "Saving request..." : "Submit request"}</button>
      </form>
      {error && <p className="error" role="alert">{error}</p>}
      {saved && <p role="status">Request saved: {saved.id}. {statusLabel(saved)}. No allowance has changed.</p>}
    </section>
    <RequestList key={refresh} api={api} />
  </>;
}

function RequestList({ api, inbox = false }: { api: RequestApi; inbox?: boolean }) {
  const [page, setPage] = useState<Page<RequestRecord>>({ items: [] });
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const controller = useRef<AbortController | null>(null);
  async function load(cursor?: string) {
    controller.current?.abort();
    const current = new AbortController(); controller.current = current;
    setBusy(true); setError("");
    try {
      const next = await (inbox ? api.inbox(current.signal, cursor) : api.history(current.signal, cursor));
      if (!current.signal.aborted) setPage((prev) => ({ ...next, items: cursor ? [...prev.items, ...next.items] : next.items }));
    } catch (error) { if (!current.signal.aborted) setError(message(error)); }
    finally { if (!current.signal.aborted) setBusy(false); }
  }
  useEffect(() => { void load(); return () => controller.current?.abort(); }, [api, inbox]);
  return <section className="card"><div className="section-heading"><h2>{inbox ? "Owner approval inbox" : "Your request history"}</h2>
    <button className="secondary" disabled={busy} onClick={() => void load()}>Refresh requests</button></div>
    {error && <p className="error" role="alert">{error}</p>}
    {busy && <p role="status">Loading requests...</p>}
    {!busy && !error && !page.items.length && <p>No requests found.</p>}
    {page.items.map((record) => <RequestCard key={`${record.id}:${record.etag}`} record={record} api={api} inbox={inbox}
      onDecision={(saved) => setPage((prev) => ({ ...prev, items: prev.items.map((r) => r.id === saved.id ? saved : r) }))} />)}
    {page.cursor && <button disabled={busy} onClick={() => void load(page.cursor)}>Load more requests</button>}
  </section>;
}

function RequestCard({ record, api, inbox, onDecision }: {
  record: RequestRecord; api: RequestApi; inbox: boolean; onDecision: (record: RequestRecord) => void;
}) {
  const [reason, setReason] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const controller = useRef<AbortController | null>(null);
  useEffect(() => () => controller.current?.abort(), []);
  async function decide(status: "approved" | "rejected") {
    setBusy(true); setError("");
    const current = new AbortController(); controller.current = current;
    try {
      const saved = await api.decide(record.id, { status, etag: record.etag, ...(status === "rejected" ? { reason } : {}) }, current.signal);
      if (!current.signal.aborted) onDecision(saved);
    } catch (error) { if (!current.signal.aborted) setError(message(error)); }
    finally { if (!current.signal.aborted) setBusy(false); }
  }
  const expired = !!record.expiresAt && Date.parse(record.expiresAt) <= Date.now();
  return <article className="card"><h3>{statusLabel(record)}</h3>
    <dl><dt>Request</dt><dd className="mono">{record.id}</dd><dt>Budget</dt><dd>{record.budgetId}</dd>
      {inbox && <><dt>Requester object ID</dt><dd className="mono">{record.requesterId}</dd></>}
      <dt>Submitted</dt><dd>{new Date(record.createdAt).toLocaleString()}</dd><dt>Duration</dt>
      <dd>{record.expiresAt ? `Until ${record.expiresAt} (exclusive)` : "Permanent"}</dd></dl>
    <p className="justification">{record.justification}</p>
    {record.rejectionReason && <p>Rejection reason: {record.rejectionReason}</p>}
    {record.decisionAt && <p className="small muted">Decision recorded {record.decisionAt}</p>}
    {inbox && record.status === "pending" && !expired && <div className="request-form">
      <label htmlFor={`reason-${record.id}`}>Reason if rejecting (10-1000 characters)</label>
      <textarea id={`reason-${record.id}`} value={reason} maxLength={1000} disabled={busy} onChange={(e) => setReason(e.target.value)} />
      <div className="actions"><button disabled={busy} onClick={() => void decide("approved")}>Approve / manual fulfilment</button>
        <button className="secondary" disabled={busy || reason.trim().length < 10} onClick={() => void decide("rejected")}>Reject</button></div>
    </div>}
    {error && <p className="error" role="alert">{error}</p>}
  </article>;
}

function OwnerGraph({ budgets, serviceId, auth, account, api }: {
  budgets: BudgetConfig[]; serviceId: string; auth: DemoAuth; account: AccountInfo; api: RequestApi;
}) {
  const [budgetId, setBudgetId] = useState(budgets[0]?.id ?? "");
  const [snapshot, setSnapshot] = useState<BudgetSnapshot | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const controller = useRef<AbortController | null>(null);
  const budget = budgets.find((b) => b.id === budgetId);
  useEffect(() => { setSnapshot(null); controller.current?.abort(); }, [budget]);
  useEffect(() => () => controller.current?.abort(), []);
  async function load() {
    if (!budget) return;
    controller.current?.abort();
    const current = new AbortController(); controller.current = current;
    setBusy(true); setError(""); setSnapshot(null);
    try {
      await auth.consentGraph(account);
      if (current.signal.aborted) return;
      await loadBudgetDetails(budget, serviceId, new GraphClient(() => auth.token(account)), current.signal,
        (next) => { if (!current.signal.aborted) setSnapshot(next); });
    } catch (error) { if (!current.signal.aborted) setError(message(error)); }
    finally { if (!current.signal.aborted) setBusy(false); }
  }
  return <><section className="card"><h2>Live Graph details</h2>
    <p className="muted">Only your server-authorized budgets are available. Graph uses a separate delegated token and the existing eight owner permissions.</p>
    <label>Budget <select value={budgetId} disabled={busy} onChange={(e) => setBudgetId(e.target.value)}>
      {budgets.map((b) => <option key={b.id} value={b.id}>{b.displayName}</option>)}</select></label>
    <button disabled={busy || !budget} onClick={() => void load()}>{busy ? "Loading Graph..." : "Connect / refresh Graph details"}</button>
    {error && <p className="error" role="alert">{error}</p>}
  </section>{budget?.approvedCreditCeiling && <BudgetAllocation key={budget.id} budget={budget} api={api} auth={auth} account={account} onApplied={() => setSnapshot(null)} />}
    {budget && snapshot?.budgetId === budget.id && <BudgetDetails budget={budget} snapshot={snapshot} serviceId={serviceId} />}</>;
}
