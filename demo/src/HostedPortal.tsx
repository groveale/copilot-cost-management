import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import { DemoAuth, isGraphInteractionRequired } from "./auth";
import type { BudgetConfig, DemoConfig, GrantKind } from "./config";
import { RequestApi, type Context, type OnboardingCapabilities, type Page, type RequestRecord, type Submission } from "./api";
import { GraphClient } from "./graph";
import { loadBudgetDetails, type BudgetSnapshot } from "./details";
import { BudgetDetails } from "./BudgetDetails";
import { BusinessUnitOnboarding, BudgetAllocation } from "./Onboarding";
import { AdminPolicies } from "./AdminPolicies";
import { PersonalBalance } from "./PersonalBalance";
import { ThemeToggle } from "./ThemeToggle";
import { requestStatusLabel as statusLabel, runFulfilment } from "./fulfilment";

const message = (error: unknown) => error instanceof Error ? error.message : String(error);

export type PortalView = "admin" | "owner" | "request";
export function portalView(search: string): PortalView {
  const view = new URLSearchParams(search).get("view");
  return view === "admin" || view === "request" ? view : "owner";
}
const viewNames: Record<PortalView, string> = { admin: "Admin", owner: "Policy owner", request: "Request credits" };
export type OwnerPage = "budgets" | "requests" | "setup";
export interface OwnerLocation { page: OwnerPage; budgetId?: string }
export function ownerLocation(search: string): OwnerLocation {
  const params = new URLSearchParams(search);
  const page = params.get("ownerPage");
  if (page === "requests" || page === "setup") return { page };
  const budgetId = params.get("budget");
  return { page: "budgets", ...(budgetId ? { budgetId } : {}) };
}
const ownerPageNames: Record<OwnerPage, string> = { budgets: "Business units", requests: "Credit requests", setup: "Set up a BU" };

export function OwnerNavigation({ location, onNavigate }: {
  location: OwnerLocation; onNavigate: (location: OwnerLocation) => void;
}) {
  return <nav className="actions owner-navigation" aria-label="Policy owner pages">
    {(Object.keys(ownerPageNames) as OwnerPage[]).map((page) => <button key={page}
      className={location.page === page ? "" : "secondary"} aria-pressed={location.page === page}
      onClick={() => onNavigate({ page })}>{ownerPageNames[page]}</button>)}
  </nav>;
}

export function PortalNavigation({ view, isAdmin, onSelect }: {
  view: PortalView; isAdmin: boolean; onSelect: (view: PortalView) => void;
}) {
  const views: PortalView[] = isAdmin ? ["admin", "owner", "request"] : ["owner", "request"];
  return <nav className="actions portal-navigation" aria-label="Portal views">
    {views.map((item) => <button key={item} className={view === item ? "" : "secondary"} aria-pressed={view === item}
      onClick={() => onSelect(item)}>{viewNames[item]}</button>)}
  </nav>;
}

export function HostedPortal({ config, auth, initialView }: { config: DemoConfig; auth: DemoAuth; initialView: PortalView }) {
  const [account, setAccount] = useState(auth.account());
  const [view, setView] = useState(initialView);
  const [owner, setOwner] = useState(() => ownerLocation(typeof window === "undefined" ? "" : window.location.search));
  const [adminSection, setAdminSection] = useState<"requests" | "policies">("requests");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [capabilities, setCapabilities] = useState<OnboardingCapabilities | null>(null);
  const [capabilityError, setCapabilityError] = useState("");
  const [context, setContext] = useState<Context | null>(null);
  const [refresh, setRefresh] = useState(0);
  const api = useMemo(() => new RequestApi(config.apiBaseUrl ?? "", () => auth.apiToken(account)), [auth, config.apiBaseUrl, account]);
  const onPublished = useCallback(() => setRefresh((value) => value + 1), []);
  useEffect(() => {
    const onPopState = () => {
      setView(portalView(window.location.search));
      setOwner(ownerLocation(window.location.search));
    };
    window.addEventListener("popstate", onPopState);
    return () => window.removeEventListener("popstate", onPopState);
  }, []);
  useEffect(() => {
    document.title = `Copilot Budget Portal | ${view === "owner" ? `Policy owner | ${ownerPageNames[owner.page]}` : viewNames[view]}`;
  }, [view, owner.page]);
  useEffect(() => {
    if (!account || busy) return;
    const controller = new AbortController();
    setCapabilityError("");
    void api.onboardingCapabilities(controller.signal).then((value) => {
      if (!controller.signal.aborted) setCapabilities(value);
    }).catch((error: unknown) => {
      if (!controller.signal.aborted) { setCapabilities(null); setCapabilityError(message(error)); }
    });
    return () => controller.abort();
  }, [account, api, busy, refresh]);
  useEffect(() => {
    setContext(null);
    if (!account || busy) return;
    setError("");
    if (view === "admin") return;
    const controller = new AbortController();
    void api.context(controller.signal).then((value) => {
      if (!controller.signal.aborted) setContext(value);
    }).catch((error: unknown) => { if (!controller.signal.aborted) setError(message(error)); });
    return () => controller.abort();
  }, [account, api, busy, refresh, view]);

  function navigate(next: PortalView) {
    const url = new URL(window.location.href);
    url.searchParams.set("view", next);
    url.searchParams.delete("ownerPage");
    url.searchParams.delete("budget");
    window.history.pushState(null, "", url);
    setView(next);
    setOwner({ page: "budgets" });
  }
  function navigateOwner(next: OwnerLocation) {
    const url = new URL(window.location.href);
    url.searchParams.set("view", "owner");
    url.searchParams.set("ownerPage", next.page);
    if (next.page === "budgets" && next.budgetId) url.searchParams.set("budget", next.budgetId);
    else url.searchParams.delete("budget");
    window.history.pushState(null, "", url);
    setOwner(next);
  }

  async function authenticate(signOut = false) {
    setAccount(null);
    setContext(null);
    setCapabilities(null);
    setCapabilityError("");
    setError("");
    setBusy(true);
    try {
      if (signOut) await auth.signOut();
      else setAccount(await auth.signIn("requester"));
    } catch (error) { setError(message(error)); }
    finally { setBusy(false); }
  }
  if (!account || busy) return <main className="login-shell"><section className="card login-card">
    <div className="section-heading"><span className="eyebrow">COPILOT BUDGET PORTAL</span><ThemeToggle /></div>
    <h1>{view === "request" ? "Request a higher allowance" : "Sign in to your budgets"}</h1>
    <p className="muted">Use your work account. Requests do not require Graph permissions.</p>
    <button disabled={busy} onClick={() => void authenticate()}>{busy ? "Completing sign-in / sign-out..." : "Sign in with Microsoft"}</button>
    {error && <p className="error" role="alert">{error}</p>}
  </section></main>;

  const workspaceRequired = view === "request" || (view === "owner" && owner.page !== "setup");
  return <main className="shell">
    <header className="masthead"><div><span className="eyebrow">COPILOT BUDGET PORTAL</span>
      <h1>{viewNames[view]}</h1>
      <p className="muted">{view === "admin" ? "Review setup requests and inspect spending-policy ownership." :
        view === "owner" ? "Browse your business units or review employee credit requests." : "Ask your policy owner for an Enhanced allowance."}</p></div>
      <div className="operator"><p>{account.username}</p><div className="actions">
        <ThemeToggle />
        <button className="secondary" onClick={() => void authenticate()}>Switch account / renew sign-in</button>
        <button className="secondary" onClick={() => void authenticate(true)}>Sign out</button>
      </div></div></header>
    <PortalNavigation view={view} isAdmin={capabilities?.isAdmin === true} onSelect={navigate} />
    <div className="actions"><button className="secondary" onClick={onPublished}>Refresh workspace</button></div>
    {capabilityError && <p className="error" role="alert">Could not verify administrator access. {capabilityError}</p>}
    {view === "admin" && <>
      {!capabilities && !capabilityError && <p role="status">Checking administrator access...</p>}
      {capabilities && !capabilities.isAdmin && <p className="error" role="alert">Administrator access is required. Choose Policy owner or Request credits.</p>}
      {capabilities?.isAdmin && <>
        <nav className="actions" aria-label="Admin sections">
          <button className={adminSection === "requests" ? "" : "secondary"} aria-pressed={adminSection === "requests"}
            onClick={() => setAdminSection("requests")}>Setup requests</button>
          <button className={adminSection === "policies" ? "" : "secondary"} aria-pressed={adminSection === "policies"}
            onClick={() => setAdminSection("policies")}>Spending policies</button>
        </nav>
        {adminSection === "requests" ? <BusinessUnitOnboarding api={api} auth={auth} account={account}
          mode="admin" capabilities={capabilities} onPublished={onPublished} /> :
          <AdminPolicies key={refresh} api={api} auth={auth} account={account} />}
      </>}
    </>}
    {workspaceRequired && error && <p className="error" role="alert">{error}</p>}
    {view === "request" && <PersonalBalance key={`${account.homeAccountId}:${account.localAccountId}:${refresh}`}
      auth={auth} account={account} serviceId={config.serviceId} />}
    {workspaceRequired && !context && !error && <p role="status">Loading your authorized workspace...</p>}
    {view === "request" && context && <>
      <CreditApprovalNotice />
      <SubmissionForm api={api} context={context} />
    </>}
    {view === "owner" && <OwnerWorkspace location={owner} onNavigate={navigateOwner} context={context}
      api={api} auth={auth} account={account} capabilities={capabilities} onPublished={onPublished} />}
    <footer>Evaluation only. Requests are stored privately. Graph preview APIs are not supported for production applications.</footer>
  </main>;
}

function CreditApprovalNotice() {
  return <p className="notice">Approval moves the requester from Standard to Enhanced using the owner's delegated Graph access.
    Failed or partial upgrades require explicit retry. Temporary increases require manual return to Standard at the UTC calendar month boundary.
    Owners may approve their own requests in this MVP.</p>;
}

export function BusinessUnitList({ budgets, onNavigate }: {
  budgets: BudgetConfig[]; onNavigate: (location: OwnerLocation) => void;
}) {
  return <section aria-labelledby="business-units-heading">
    <div className="section-heading workspace-heading"><div><h2 id="business-units-heading">Your business units</h2>
      <p className="muted">Select a business unit to view its Standard and Enhanced policies.</p></div>
      <button className="secondary" onClick={() => onNavigate({ page: "setup" })}>Set up a new BU</button>
    </div>
    {!budgets.length ? <div className="card"><h3>No business units yet</h3>
      <p>Set up your first BU to create its Standard and Enhanced policies. An administrator must approve it.
        If you already manage a BU, ask the administrator to check its owner mapping.</p></div> :
      <ul className="business-unit-list">{budgets.map((budget) => <li key={budget.id} className="card business-unit-card">
        <div><h3>{budget.displayName}</h3><p className="small muted">{budget.businessArea}</p></div>
        <dl className="metadata-grid">
          <div><dt>Cost centre</dt><dd>{budget.costCenter}</dd></div>
          <div><dt>Planning budget</dt><dd>{budget.plannedUsd === undefined ? "Not configured" :
            new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" }).format(budget.plannedUsd)}</dd></div>
          <div><dt>Policies</dt><dd>Standard + Enhanced</dd></div>
        </dl>
        <button aria-label={`View policies for ${budget.displayName}`}
          onClick={() => onNavigate({ page: "budgets", budgetId: budget.id })}>View policies</button>
      </li>)}</ul>}
  </section>;
}

export function OwnerWorkspace({ location, onNavigate, context, api, auth, account, capabilities, onPublished }: {
  location: OwnerLocation; onNavigate: (location: OwnerLocation) => void; context: Context | null;
  api: RequestApi; auth: DemoAuth; account: AccountInfo; capabilities: OnboardingCapabilities | null; onPublished: () => void;
}) {
  const budget = context?.ownedBudgets.find((item) => item.id === location.budgetId);
  return <div className="owner-workspace">
    <OwnerNavigation location={location} onNavigate={onNavigate} />
    {location.page === "setup" ? <>
      <div className="workspace-heading"><h2>Set up a business unit</h2>
        <p className="muted">Request a new BU and follow its setup progress. Employee credit approvals are on the Credit requests page.</p></div>
      <BusinessUnitOnboarding api={api} auth={auth} account={account} mode="owner" capabilities={capabilities} onPublished={onPublished} />
    </> : context && (location.page === "requests" ? <>
      <div className="workspace-heading"><h2>Credit requests</h2>
        <p className="muted">Review employee allowance requests across all your business units.</p></div>
      {!context.ownedBudgets.length ? <section className="card"><p>You do not own a business unit yet, so there are no credit requests to review.</p>
        <button onClick={() => onNavigate({ page: "setup" })}>Set up your first BU</button></section> :
        <><CreditApprovalNotice /><RequestList api={api} inbox auth={auth} account={account} serviceId={context.serviceId}
          budgets={context.ownedBudgets} /></>}
    </> : location.budgetId ? <>
      <div className="actions"><button className="secondary" onClick={() => onNavigate({ page: "budgets" })}>Back to business units</button></div>
      {!budget ? <p className="error" role="alert">This business unit is unavailable or you are no longer an owner. Return to your business unit list.</p> :
        <OwnerGraph key={budget.id} budget={budget} serviceId={context.serviceId} auth={auth} account={account} api={api} />}
    </> : <BusinessUnitList budgets={context.ownedBudgets} onNavigate={onNavigate} />)}
  </div>;
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
      <p className="small muted">Routing uses your verified tenant/user ID and the private requester index, not the policy reported by your browser.
        Balance access is optional and does not block this form.</p>
      <form className="request-form" onSubmit={(event) => void submit(event)}>
        <label htmlFor="justification">Business justification (20-2000 characters)</label>
        <textarea id="justification" minLength={20} maxLength={2000} required rows={6} value={justification}
          disabled={busy} onChange={(event) => setJustification(event.target.value)} />
        <fieldset disabled={busy}><legend>How long is the increase needed?</legend>
          <label className="radio-choice"><input type="radio" name="duration" checked={grantKind === "billingPeriod"} onChange={() => setGrantKind("billingPeriod")} />Rest of this billing period (demo: UTC calendar month)</label>
          <label className="radio-choice"><input type="radio" name="duration" checked={grantKind === "permanent"} onChange={() => setGrantKind("permanent")} />Permanent increase</label>
        </fieldset>
        {grantKind === "billingPeriod" && <p className="small muted">An owner or administrator must manually return you to Standard at period-end. This demo does not schedule automatic reversion.</p>}
        <button disabled={busy || context.route.state !== "ready"}>{busy ? "Saving request..." : "Submit request"}</button>
      </form>
      {error && <p className="error" role="alert">{error}</p>}
      {saved && <p role="status">Request saved: {saved.id}. {statusLabel(saved)}. An owner must approve and apply the upgrade.</p>}
    </section>
    <RequestList key={refresh} api={api} />
  </>;
}

function RequestList({ api, inbox = false, auth, account, serviceId, budgets }: {
  api: RequestApi; inbox?: boolean; auth?: DemoAuth; account?: AccountInfo; serviceId?: string;
  budgets?: BudgetConfig[];
}) {
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
    {page.items.map((record) => <RequestCard key={record.id} record={record} api={api} inbox={inbox}
      auth={auth} account={account} serviceId={serviceId}
      budgetName={budgets?.find((budget) => budget.id === record.budgetId)?.displayName}
      onDecision={(saved) => setPage((prev) => ({ ...prev, items: prev.items.map((r) => r.id === saved.id ? saved : r) }))} />)}
    {page.cursor && <button disabled={busy} onClick={() => void load(page.cursor)}>Load more requests</button>}
  </section>;
}

export function RequestCard({ record, api, inbox, onDecision, auth, account, serviceId, budgetName }: {
  record: RequestRecord; api: RequestApi; inbox: boolean; onDecision: (record: RequestRecord) => void;
  auth?: DemoAuth; account?: AccountInfo; serviceId?: string;
  budgetName?: string;
}) {
  const [reason, setReason] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const controller = useRef<AbortController | null>(null);
  useEffect(() => () => controller.current?.abort(), []);
  async function decide(status: "approved" | "rejected") {
    setBusy(true); setError("");
    const current = new AbortController(); controller.current = current;
    let approved = record.status === "approved";
    try {
      if (status === "rejected") {
        const saved = await api.decide(record.id, { status, etag: record.etag, reason }, current.signal);
        if (!current.signal.aborted) onDecision(saved);
      } else {
        if (!auth || !account || !serviceId) throw new Error("Owner Graph access is not configured. Refresh the owner workspace.");
        await auth.consentGraph(account);
        current.signal.throwIfAborted();
        let saved = record;
        if (!approved) {
          saved = await api.decide(record.id, { status: "approved", etag: record.etag }, current.signal);
          approved = true;
          if (!current.signal.aborted) onDecision(saved);
        }
        await runFulfilment(api, new GraphClient(() => auth.token(account)), saved, serviceId, current.signal,
          (value) => { if (!current.signal.aborted) onDecision(value); });
      }
    } catch (error) {
      if (!current.signal.aborted) setError(`${approved ? "Approval is saved, but the upgrade is not confirmed. Refresh and retry. " : ""}${message(error)}`);
    }
    finally { if (!current.signal.aborted) setBusy(false); }
  }
  const expired = !!record.expiresAt && Date.parse(record.expiresAt) <= Date.now();
  const activeLease = !!record.fulfilment?.lease && Date.parse(record.fulfilment.lease.until) > Date.now();
  return <article className="card"><h3>{statusLabel(record)}</h3>
    <dl><dt>Request</dt><dd className="mono">{record.id}</dd><dt>Business unit</dt><dd>{budgetName ?? record.budgetId}</dd>
      {inbox && <><dt>Requester object ID</dt><dd className="mono">{record.requesterId}</dd></>}
      <dt>Submitted</dt><dd>{new Date(record.createdAt).toLocaleString()}</dd><dt>Duration</dt>
      <dd>{record.expiresAt ? `Until ${record.expiresAt} (exclusive)` : "Permanent"}</dd></dl>
    <p className="justification">{record.justification}</p>
    {record.rejectionReason && <p>Rejection reason: {record.rejectionReason}</p>}
    {record.decisionAt && <p className="small muted">Decision recorded {record.decisionAt}</p>}
    {record.fulfilment?.appliedAt && <p className="small muted">Group move verified by the delegated browser {record.fulfilment.appliedAt}. Policy/balance updates may take time to propagate.</p>}
    {record.status === "approved" && record.expiresAt && <p className="notice">Manual return to Standard {expired ? "is overdue" : "is required"} at {record.expiresAt}.
      No automatic reversion is scheduled. Review actual membership before making the return.</p>}
    {record.fulfilment?.lease && <p className="small muted">Upgrade attempt lease until {record.fulfilment.lease.until}. If the browser was closed, wait for the lease to expire, refresh and retry.</p>}
    {inbox && record.status === "pending" && !expired && <div className="request-form">
      <label htmlFor={`reason-${record.id}`}>Reason if rejecting (10-1000 characters)</label>
      <textarea id={`reason-${record.id}`} value={reason} maxLength={1000} disabled={busy} onChange={(e) => setReason(e.target.value)} />
      <div className="actions"><button disabled={busy} onClick={() => void decide("approved")}>Approve and move to Enhanced</button>
        <button className="secondary" disabled={busy || reason.trim().length < 10} onClick={() => void decide("rejected")}>Reject</button></div>
    </div>}
    {inbox && record.status === "approved" && record.fulfilment?.state !== "applied" && !expired &&
      <button disabled={busy || activeLease} onClick={() => void decide("approved")}>
        {record.fulfilment?.state === "needs-attention" ? "Retry Enhanced upgrade" : "Apply Enhanced upgrade"}
      </button>}
    {error && <p className="error" role="alert">{error}</p>}
  </article>;
}

function OwnerGraph({ budget, serviceId, auth, account, api }: {
  budget: BudgetConfig; serviceId: string; auth: DemoAuth; account: AccountInfo; api: RequestApi;
}) {
  const [snapshot, setSnapshot] = useState<BudgetSnapshot | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [needsConsent, setNeedsConsent] = useState(false);
  const [graphReady, setGraphReady] = useState(false);
  const controller = useRef<AbortController | null>(null);
  useEffect(() => {
    void load();
    return () => controller.current?.abort();
  }, [budget, serviceId, auth, account]);
  async function load(interactive = false) {
    controller.current?.abort();
    const current = new AbortController(); controller.current = current;
    setBusy(true); setError(""); setSnapshot(null);
    try {
      if (interactive) await auth.consentGraph(account);
      await auth.token(account);
      setNeedsConsent(false);
      setGraphReady(true);
      if (current.signal.aborted) return;
      await loadBudgetDetails(budget, serviceId, new GraphClient(() => auth.token(account).catch((error: unknown) => {
        if (!current.signal.aborted) setNeedsConsent(isGraphInteractionRequired(error));
        throw error;
      })), current.signal,
        (next) => { if (!current.signal.aborted) setSnapshot(next); });
    } catch (error) {
      if (!current.signal.aborted) {
        setNeedsConsent(isGraphInteractionRequired(error));
        setError(message(error));
      }
    }
    finally { if (!current.signal.aborted) setBusy(false); }
  }
  return <><section className="card"><span className="eyebrow">BUSINESS UNIT / POLICIES</span><h2>{budget.displayName}</h2>
    <p className="muted">{budget.businessArea} / Cost centre: {budget.costCenter}</p>
    <p className="small muted">Standard and Enhanced policy settings, members and current consumption. Readings come directly from Graph and may lag behind membership changes.</p>
    <div className="actions">
      <button disabled={busy} onClick={() => void load()}>{busy ? "Loading Graph..." : "Refresh policies"}</button>
      {needsConsent && <button disabled={busy} onClick={() => void load(true)}>Grant Graph access</button>}
    </div>
    {error && <p className="error" role="alert">{error}</p>}
  </section>{graphReady && budget.approvedCreditCeiling && <BudgetAllocation key={budget.id} budget={budget} api={api} auth={auth} account={account} onApplied={() => void load()} />}
    {snapshot?.budgetId === budget.id && <BudgetDetails budget={budget} snapshot={snapshot} serviceId={serviceId} />}</>;
}
