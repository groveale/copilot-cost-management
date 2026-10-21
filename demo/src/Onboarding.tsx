import { useEffect, useId, useRef, useState, type FormEvent } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import type { Budget, Page, PolicyDefaults } from "../../api/src/contracts";
import type { BuSubmission, OnboardingRecord } from "../../api/src/onboarding";
import { allocationCeiling, checkAllocation, splitCredits, usdToCredits } from "../../api/src/budget-math";
import { RequestApi, type OnboardingCapabilities } from "./api";
import { DemoAuth, isGraphInteractionRequired } from "./auth";
import { GraphClient } from "./graph";
import { runProvisioning, readAllocation } from "./provisioning";
import type { TierName } from "./config";

const message = (error: unknown) => error instanceof Error ? error.message : String(error);
const format = (value: number) => value.toLocaleString("en-US");
const statusLabel = (record: OnboardingRecord) => ({
  pending: "Pending administrator review", provisioning: "Approved - provisioning not yet complete",
  rejected: "Rejected", completed: "Completed - owner workspace available",
})[record.status];
const summaryStatus: Record<OnboardingRecord["status"], string> = {
  pending: "Awaiting approval", provisioning: "Provisioning", rejected: "Rejected", completed: "Completed",
};
type Connection = { api: RequestApi; auth: DemoAuth; account: AccountInfo };

export function BusinessUnitOnboarding({ api, auth, account, onPublished, mode, capabilities }: Connection & {
  onPublished: () => void; mode: "owner" | "admin"; capabilities: OnboardingCapabilities | null;
}) {
  const admin = mode === "admin" && capabilities?.isAdmin === true;
  const [defaults, setDefaults] = useState<PolicyDefaults | null>(null);
  const [page, setPage] = useState<Page<OnboardingRecord>>({ items: [] });
  const [refresh, setRefresh] = useState(0);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");
  const [usd, setUsd] = useState("20000");
  const operation = useRef<{ signature: string; id: string } | null>(null);
  const controller = useRef<AbortController | null>(null);
  const published = useRef(new Set<string>());
  useEffect(() => () => controller.current?.abort(), []);
  useEffect(() => {
    if (mode === "admin" && !admin) return;
    const current = new AbortController();
    void (async () => {
      const result = await api.onboardingList(current.signal, admin);
      const metadata = admin ? await api.onboardingSetup(current.signal) : null;
      if (!current.signal.aborted) { setPage(result); setDefaults(metadata?.onboardingDefaults ?? null); }
    })().catch((error: unknown) => { if (!current.signal.aborted) setError(message(error)); });
    return () => current.abort();
  }, [api, mode, admin, refresh]);
  useEffect(() => {
    for (const item of page.items) {
      if (item.status === "completed" && !published.current.has(item.id)) {
        published.current.add(item.id); onPublished();
      }
    }
    if (busy || !page.items.some((item) => item.status === "pending" || item.status === "provisioning")) return;
    const timer = window.setInterval(() => setRefresh((value) => value + 1), 10000);
    return () => window.clearInterval(timer);
  }, [page.items, busy, onPublished]);
  async function perform(action: (signal: AbortSignal) => Promise<void>, reload = true) {
    const current = new AbortController(); controller.current = current;
    setBusy(true); setError(""); setNotice("");
    try { await action(current.signal); if (!current.signal.aborted && reload) setRefresh((value) => value + 1); }
    catch (error) { if (!current.signal.aborted) setError(message(error)); }
    finally { if (!current.signal.aborted) setBusy(false); }
  }
  function progress(record: OnboardingRecord) {
    setPage((value) => ({ ...value, items: value.items.map((item) => item.id === record.id ? record : item) }));
  }
  async function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const form = event.currentTarget;
    const data = new FormData(form);
    const fields = { name: String(data.get("name")).trim(), costCenter: String(data.get("costCenter")).trim(),
      monthlyBudgetUsd: usd, justification: String(data.get("justification")).trim() };
    const signature = JSON.stringify(fields);
    if (operation.current?.signature !== signature) operation.current = { signature, id: crypto.randomUUID() };
    const input: BuSubmission = { ...fields, operationId: operation.current.id };
    await perform(async (signal) => {
      await api.onboardingSubmit(input, signal);
      if (!signal.aborted) { form.reset(); operation.current = null; setNotice("BU request saved. An administrator must validate and approve it before anything is provisioned."); }
    });
  }
  let calculation = "";
  try {
    const total = usdToCredits(usd); const split = splitCredits(total);
    calculation = `${format(total)} credits/month: ${format(split.standard)} Standard + ${format(split.enhanced)} Enhanced.`;
  } catch (error) { calculation = message(error); }
  if (mode === "admin" && !admin) return <p className="error" role="alert">Administrator access is required.</p>;
  return <section className="onboarding-workspace">
    {mode === "owner" && <section className="card"><h2>Request a new business unit</h2>
      <p>Your proposed ownership comes from your signed-in identity. Approval does not grant you directory roles.</p>
      <p>Proposed owner: {capabilities?.displayName ?? account.name ?? account.username}.
        {" "}Contact: {capabilities?.contactHint ?? account.username ?? "Not supplied by sign-in"}.</p>
      <form className="request-form bu-request-form" onSubmit={(event) => void submit(event)}>
        <fieldset className="bu-form-fields" disabled={busy}><legend>Business unit details</legend>
          <div className="bu-form-grid">
            <label><span>Business unit name</span><input name="name" required minLength={2} maxLength={80} /></label>
            <label><span>Cost centre</span><input name="costCenter" required maxLength={80} /></label>
          </div>
          <label><span>Requested monthly planning budget (USD)</span><input name="budget" type="number" min="0.01" step="0.01" required
            value={usd} onChange={(event) => setUsd(event.target.value)} /></label>
          <div className="bu-budget-preview"><span className="small muted">Proposed monthly credit allocation</span>
            <p role="status">{calculation}</p></div>
          <p className="small muted">Demo planning conversion: USD 0.01 per Copilot Credit. Not actual invoiced cost or a credit purchase.
            Policies reset each UTC calendar month. A spare credit goes to Enhanced when the total is odd.</p>
          <label><span>Business justification</span><textarea name="justification" required minLength={20} maxLength={2000} rows={5} /></label>
          <div><button>Submit BU request</button></div>
        </fieldset>
      </form>
    </section>}
    {error && <p className="error" role="alert">{error}</p>}
    {notice && <p role="status">{notice}</p>}
    <section className="card"><h2>{admin ? "Administrator onboarding queue" : "Your BU requests"}</h2>
      <button className="secondary" disabled={busy} onClick={() => { setError(""); setRefresh((value) => value + 1); }}>Refresh onboarding</button>
      {!page.items.length && <p>No onboarding requests loaded.</p>}
    </section>
    {page.items.map((record) => <OnboardingCard key={record.id} record={record} admin={admin}
      defaults={defaults} api={api} auth={auth} account={account} busy={busy} perform={perform} progress={progress} onPublished={onPublished} />)}
    {page.cursor && <button disabled={busy} onClick={() => void perform(async (signal) => {
      const next = await api.onboardingList(signal, admin, page.cursor);
      if (!signal.aborted) setPage({ ...next, items: [...page.items, ...next.items] });
    }, false)}>Load more onboarding requests</button>}
  </section>;
}

export function OnboardingCard({ record, admin, defaults, api, auth, account, busy, perform, progress, onPublished }: Connection & {
  record: OnboardingRecord; admin: boolean; defaults: PolicyDefaults | null; busy: boolean;
  perform: (action: (signal: AbortSignal) => Promise<void>) => Promise<void>;
  progress: (record: OnboardingRecord) => void; onPublished: () => void;
}) {
  const [expanded, setExpanded] = useState(false);
  const detailsId = useId();
  const [needsConsent, setNeedsConsent] = useState(false);
  const [confirmed, setConfirmed] = useState(false);
  const [note, setNote] = useState("");
  async function graphToken() {
    try {
      const token = await auth.token(account);
      setNeedsConsent(false);
      return token;
    } catch (error) {
      setNeedsConsent(isGraphInteractionRequired(error));
      throw error;
    }
  }
  return <article className="card request-card bu-request-card">
    <div className="request-card-heading">
      <div><h3>{record.name}</h3><p className="small muted">{record.costCenter} / ${record.monthlyBudgetUsd} per month
        {" / "}<time dateTime={record.createdAt}>{new Date(record.createdAt).toLocaleDateString()}</time></p></div>
      <div className="request-card-controls">
        <span className={`badge request-status-badge request-status-${record.status}`}>{summaryStatus[record.status]}</span>
        <button type="button" className="secondary" aria-expanded={expanded} aria-controls={detailsId}
          aria-label={`${expanded ? "Collapse" : "Expand"} BU request ${record.name}`}
          onClick={() => setExpanded((value) => !value)}>{expanded ? "Collapse" : "Expand"}</button>
      </div>
    </div>
    <div id={detailsId} className="request-card-details" hidden={!expanded}>
    <p className="small muted">{statusLabel(record)}</p>
    <dl><dt>Cost centre</dt><dd>{record.costCenter}</dd><dt>Requested monthly USD</dt><dd>${record.monthlyBudgetUsd}</dd>
      <dt>Proposed owner</dt><dd>{record.displayName ?? record.requesterId} {record.contactHint && `(${record.contactHint})`}</dd>
      <dt>Standard / Enhanced credits</dt><dd>{format(record.standard.credits)} / {format(record.enhanced.credits)}</dd>
      {record.approvedCredits > 0 && <><dt>Approved monthly ceiling</dt><dd>{format(record.approvedCredits)} credits</dd></>}
    </dl><p className="justification">{record.justification}</p>
    {record.reviewNote && <p>Finance review: {record.reviewNote}</p>}
    {record.rejectionReason && <p>Rejection reason: {record.rejectionReason}</p>}
    {record.status === "provisioning" && <p>Created resources: {Object.keys(record.resources).length}/4.
      Incomplete operations retain their progress. No new ownership is published until verification completes.</p>}
    {admin && record.status === "pending" && <div className="request-form">
      {defaults ? <p>Fixed per-person defaults: Standard {defaults.standard.perUserCredits === null ? "Unlimited" : format(defaults.standard.perUserCredits)} credits,
        {" "}{defaults.standard.billingMethods}; Enhanced {defaults.enhanced.perUserCredits === null ? "Unlimited" : format(defaults.enhanced.perUserCredits)} credits,
        {" "}{defaults.enhanced.billingMethods}. Existing BU policies are not changed.</p> :
        <p className="notice">Fixed defaults are not loaded. Refresh setup requests or ask an onboarding administrator to configure the private setup metadata before approving.</p>}
      <label><input type="checkbox" checked={confirmed} disabled={busy} onChange={(event) => setConfirmed(event.target.checked)} />
        I have validated this cost centre and cross-charging arrangement.</label>
      <label>Validation note / rejection reason (10-1000 characters)
        <textarea value={note} disabled={busy} maxLength={1000} onChange={(event) => setNote(event.target.value)} /></label>
      <div className="actions">
        <button disabled={busy || !confirmed || note.trim().length < 10 || !defaults} onClick={() => void perform(async (signal) => {
          if (!defaults) throw new Error("Configure the fixed private policy defaults before approving.");
          await graphToken();
          signal.throwIfAborted();
          const approved = await api.onboardingDecide(record.id, { status: "approved", etag: record.etag,
            financeConfirmed: true, reviewNote: note.trim(), ...defaults }, signal);
          progress(approved);
          const saved = await runProvisioning(api, new GraphClient(graphToken), approved, signal, progress);
          if (!signal.aborted && saved.status === "completed") onPublished();
        })}>Approve and provision</button>
        <button className="secondary" disabled={busy || note.trim().length < 10} onClick={() => void perform(async (signal) => {
          progress(await api.onboardingDecide(record.id, { status: "rejected", etag: record.etag, reason: note.trim() }, signal));
        })}>Reject</button>
      </div>
    </div>}
    {admin && record.status === "provisioning" && <button disabled={busy} onClick={() => void perform(async (signal) => {
      await graphToken();
      const saved = await runProvisioning(api, new GraphClient(graphToken), record, signal, progress);
      if (!signal.aborted && saved.status === "completed") onPublished();
    })}>Resume provisioning</button>}
    {admin && needsConsent && <button disabled={busy} onClick={() => void perform(async () => {
      await auth.consentGraph(account);
      await graphToken();
    })}>Grant Graph access</button>}
    </div>
    {record.lastError && <p className="error" role="alert">{record.lastError}</p>}
  </article>;
}

export function BudgetAllocation({ budget, tier, api, auth, account, onApplied, disabled = false, onBusyChange }: Connection & {
  budget: Budget; tier: TierName; onApplied: () => void; disabled?: boolean; onBusyChange: (busy: boolean) => void;
}) {
  const [editing, setEditing] = useState(false);
  const detailsId = useId();
  const [monthly, setMonthly] = useState("");
  const [perUser, setPerUser] = useState("");
  const [unlimited, setUnlimited] = useState(false);
  const [loaded, setLoaded] = useState<Awaited<ReturnType<typeof readAllocation>> | null>(null);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const [error, setError] = useState("");
  const [needsConsent, setNeedsConsent] = useState(false);
  const controller = useRef<AbortController | null>(null);
  const operation = useRef<{ signature: string; id: string } | null>(null);
  useEffect(() => () => controller.current?.abort(), []);
  const other = tier === "standard" ? "enhanced" : "standard";
  const ceiling = allocationCeiling(budget);
  const label = tier === "standard" ? "Standard" : "Enhanced";
  async function perform(action: (graph: GraphClient, signal: AbortSignal) => Promise<void>, interactive = false) {
    controller.current?.abort();
    const current = new AbortController(); controller.current = current;
    setBusy(true); onBusyChange(true); setError(""); setNotice("");
    try {
      if (interactive) await auth.consentGraph(account);
      await auth.token(account);
      setNeedsConsent(false);
      await action(new GraphClient(() => auth.token(account)), current.signal);
    } catch (error) {
      if (!current.signal.aborted) { setNeedsConsent(isGraphInteractionRequired(error)); setError(message(error)); }
    }
    finally { if (!current.signal.aborted) setBusy(false); onBusyChange(false); }
  }
  async function loadLimits(graph: GraphClient, signal: AbortSignal) {
    const values = await readAllocation(graph, budget, signal);
    if (!signal.aborted) {
      setLoaded(values); setMonthly(String(values[tier]));
      setUnlimited(values.perUser[tier] === null); setPerUser(values.perUser[tier] === null ? "" : String(values.perUser[tier]));
    }
  }
  async function run(graph: GraphClient, record: OnboardingRecord, signal: AbortSignal) {
    const saved = await runProvisioning(api, graph, record, signal, () => undefined);
    if (!signal.aborted && saved.status === "completed") {
      setNotice("Policy limits verified in Graph. Live details are refreshing.");
      operation.current = null; onApplied();
    }
  }
  return <div className="policy-editor">
    <button type="button" className="secondary" disabled={disabled || busy} aria-expanded={editing} aria-controls={detailsId}
      aria-label={`Edit ${label} policy limits`} onClick={() => {
        setEditing((value) => !value);
        if (!editing) void perform(loadLimits);
      }}>{editing ? "Close editor" : "Edit"}</button>
    {budget.approvedCreditCeiling === undefined && <p className="notice small">
      {"Demo planning budget: \u00a35,000/month = 500,000 credits shared by Standard and Enhanced, assuming \u00a30.01 per credit. Not financial approval or a currency conversion."}
      {" "}Existing Graph limits change only when you explicitly save.
    </p>}
    <div id={detailsId} hidden={!editing}>
    <p className="small muted">BU monthly ceiling: {format(ceiling)} Copilot Credits.
      The other policy's limits are left unchanged. This does not prevent separately authorized admin-center changes.</p>
    <form className="policy-limit-form" onSubmit={(event) => {
      event.preventDefault();
      void perform(async (graph, signal) => {
        if (!loaded || monthly.trim() === "" || (!unlimited && perUser.trim() === "")) {
          throw new Error("Load the current limits and enter both values before saving.");
        }
        const credits = Number(monthly);
        const userCredits = unlimited ? null : Number(perUser);
        if (userCredits !== null && (!Number.isFinite(userCredits) || userCredits < 0 || userCredits > Number.MAX_SAFE_INTEGER)) {
          throw new Error("The per-user limit must be a non-negative number of credits, or explicitly Unlimited.");
        }
        const latest = await readAllocation(graph, budget, signal);
        if (latest[tier] !== loaded[tier] || latest.perUser[tier] !== loaded.perUser[tier]) {
          throw new Error("This policy changed since the editor opened. Refresh the limits, or resume a saved edit before making further changes.");
        }
        const standard = tier === "standard" ? credits : latest.standard;
        const enhanced = tier === "enhanced" ? credits : latest.enhanced;
        checkAllocation(standard, enhanced, ceiling);
        const perUserEdit = { tier, credits: userCredits };
        const signature = JSON.stringify({ standard, enhanced, perUserEdit });
        if (operation.current?.signature !== signature) operation.current = { signature, id: crypto.randomUUID() };
        const saved = await api.allocation(budget.id, signal, { operationId: operation.current.id, standard, enhanced, perUserEdit });
        if (!saved) throw new Error("No saved policy edit was returned.");
        await run(graph, saved, signal);
      });
    }}><fieldset disabled={disabled || busy || !loaded}><legend>{label} policy limits</legend>
      <label>Monthly allocation (credits)<input type="number" min="0" step="1"
        max={loaded ? Math.max(0, ceiling - loaded[other]) : undefined}
        required value={monthly} onChange={(event) => setMonthly(event.target.value)} /></label>
      <label>Per-user credit limit<input type="number" min="0" step="any" max={Number.MAX_SAFE_INTEGER}
        required={!unlimited} disabled={unlimited} value={perUser} onChange={(event) => setPerUser(event.target.value)} /></label>
      <label className="checkbox-choice"><input type="checkbox" checked={unlimited}
        onChange={(event) => setUnlimited(event.target.checked)} />Unlimited per-user credits</label>
      <button type="submit">Save policy limits</button>
    </fieldset></form>
      <div className="actions">
        <button type="button" className="secondary" disabled={disabled || busy} onClick={() => void perform(loadLimits)}>Refresh limits</button>
        <button type="button" className="secondary" disabled={disabled || busy} onClick={() => void perform(async (graph, signal) => {
          const saved = await api.allocation(budget.id, signal);
          if (!saved || saved.status === "completed") throw new Error("There is no unfinished allocation to resume.");
          if (saved.perUserEdit && saved.perUserEdit.tier !== tier) throw new Error("The unfinished edit belongs to the other policy. Open that policy's editor to resume it.");
          await run(graph, saved, signal);
        })}>Resume saved policy edit</button>
      </div>
    </div>
    {needsConsent && <button disabled={disabled || busy} onClick={() => void perform(loadLimits, true)}>Grant Graph access</button>}
    {busy && <p role="status">Working with Graph. Keep this page open.</p>}
    {error && <p className="error" role="alert">{error}</p>}
    {notice && <p role="status">{notice}</p>}
  </div>;
}
