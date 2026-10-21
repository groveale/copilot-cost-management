import { useEffect, useRef, useState, type FormEvent } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import type { Budget, Page, PolicyDefaults } from "../../api/src/contracts";
import type { BuSubmission, OnboardingRecord } from "../../api/src/onboarding";
import { checkAllocation, splitCredits, usdToCredits } from "../../api/src/budget-math";
import { RequestApi, type OnboardingCapabilities } from "./api";
import { DemoAuth, isGraphInteractionRequired } from "./auth";
import { GraphClient } from "./graph";
import { runProvisioning, readAllocation } from "./provisioning";

const message = (error: unknown) => error instanceof Error ? error.message : String(error);
const format = (value: number) => value.toLocaleString("en-US");
const statusLabel = (record: OnboardingRecord) => ({
  pending: "Pending administrator review", provisioning: "Approved - provisioning not yet complete",
  rejected: "Rejected", completed: "Completed - owner workspace available",
})[record.status];
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
      <form className="request-form" onSubmit={(event) => void submit(event)}>
        <fieldset disabled={busy}><legend>Business unit details</legend>
          <label>Business unit name <input name="name" required minLength={2} maxLength={80} /></label>
          <label>Cost centre <input name="costCenter" required maxLength={80} /></label>
          <label>Requested monthly planning budget (USD) <input name="budget" type="number" min="0.01" step="0.01" required
            value={usd} onChange={(event) => setUsd(event.target.value)} /></label>
          <p role="status">{calculation}</p>
          <p className="small muted">Demo planning conversion: USD 0.01 per Copilot Credit. Not actual invoiced cost or a credit purchase.
            Policies reset each UTC calendar month. A spare credit goes to Enhanced when the total is odd.</p>
          <label>Business justification <textarea name="justification" required minLength={20} maxLength={2000} rows={4} /></label>
          <button>Submit BU request</button>
        </fieldset>
      </form>
    </section>}
    {admin && <section className="card"><h2>Onboarding administration</h2>
      <p>Validate cross-charging and the cost centre outside this demo. These fields are metadata, not a finance-system integration.</p>
      <p>New BUs always receive Standard and Enhanced policies. Per-person caps and billing methods come from fixed private JSON, not another BU.</p>
      {capabilities?.bootstrapEnabled && <label>Import private setup JSON
        <input type="file" accept=".json,application/json" disabled={busy} onChange={(event) => {
          const file = event.target.files?.[0];
          if (!file) return;
          void perform(async (signal) => {
            if (file.size > 1_048_576) throw new Error("Seed file exceeds 1 MiB.");
            const input: unknown = JSON.parse(await file.text());
            try {
              const result = await api.bootstrap(input, signal);
              if (!signal.aborted) { setNotice(`Imported ${result.budgets} budgets and ${result.mappings} mappings. Existing requests were not reset.`); onPublished(); }
            } catch (error) {
              throw new Error(`Import was not confirmed and may be partial. Retry the same file after resolving the error; existing records are not cleared. ${message(error)}`);
            }
          });
        }} />
      </label>}
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

function OnboardingCard({ record, admin, defaults, api, auth, account, busy, perform, progress, onPublished }: Connection & {
  record: OnboardingRecord; admin: boolean; defaults: PolicyDefaults | null; busy: boolean;
  perform: (action: (signal: AbortSignal) => Promise<void>) => Promise<void>;
  progress: (record: OnboardingRecord) => void; onPublished: () => void;
}) {
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
  return <article className="card"><h3>{record.name}: {statusLabel(record)}</h3>
    <dl><dt>Cost centre</dt><dd>{record.costCenter}</dd><dt>Requested monthly USD</dt><dd>${record.monthlyBudgetUsd}</dd>
      <dt>Proposed owner</dt><dd>{record.displayName ?? record.requesterId} {record.contactHint && `(${record.contactHint})`}</dd>
      <dt>Standard / Enhanced credits</dt><dd>{format(record.standard.credits)} / {format(record.enhanced.credits)}</dd>
      {record.approvedCredits > 0 && <><dt>Approved monthly ceiling</dt><dd>{format(record.approvedCredits)} credits</dd></>}
    </dl><p className="justification">{record.justification}</p>
    {record.reviewNote && <p>Finance review: {record.reviewNote}</p>}
    {record.rejectionReason && <p>Rejection reason: {record.rejectionReason}</p>}
    {record.lastError && <p className="error" role="alert">{record.lastError}</p>}
    {record.status === "provisioning" && <p>Created resources: {Object.keys(record.resources).length}/4.
      Incomplete operations retain their progress. No new ownership is published until verification completes.</p>}
    {admin && record.status === "pending" && <div className="request-form">
      {defaults ? <p>Fixed per-person defaults: Standard {defaults.standard.perUserCredits === null ? "Unlimited" : format(defaults.standard.perUserCredits)} credits,
        {" "}{defaults.standard.billingMethods}; Enhanced {defaults.enhanced.perUserCredits === null ? "Unlimited" : format(defaults.enhanced.perUserCredits)} credits,
        {" "}{defaults.enhanced.billingMethods}. Existing BU policies are not changed.</p> :
        <p className="notice">Fixed defaults are not loaded. Refresh setup requests or import private setup JSON before approving.</p>}
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
  </article>;
}

export function BudgetAllocation({ budget, api, auth, account, onApplied }: Connection & { budget: Budget; onApplied: () => void }) {
  const [standard, setStandard] = useState("");
  const [enhanced, setEnhanced] = useState("");
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const [error, setError] = useState("");
  const [needsConsent, setNeedsConsent] = useState(false);
  const controller = useRef<AbortController | null>(null);
  const operation = useRef<{ signature: string; id: string } | null>(null);
  useEffect(() => {
    if (!budget.approvedCreditCeiling) return;
    void perform(loadLimits);
    return () => controller.current?.abort();
  }, [budget, auth, account]);
  if (!budget.approvedCreditCeiling) return null;
  async function perform(action: (graph: GraphClient, signal: AbortSignal) => Promise<void>, interactive = false) {
    controller.current?.abort();
    const current = new AbortController(); controller.current = current;
    setBusy(true); setError(""); setNotice("");
    try {
      if (interactive) await auth.consentGraph(account);
      await auth.token(account);
      setNeedsConsent(false);
      await action(new GraphClient(() => auth.token(account)), current.signal);
    } catch (error) {
      if (!current.signal.aborted) { setNeedsConsent(isGraphInteractionRequired(error)); setError(message(error)); }
    }
    finally { if (!current.signal.aborted) setBusy(false); }
  }
  async function loadLimits(graph: GraphClient, signal: AbortSignal) {
    const values = await readAllocation(graph, budget, signal);
    if (!signal.aborted) { setStandard(String(values.standard)); setEnhanced(String(values.enhanced)); }
  }
  async function run(graph: GraphClient, record: OnboardingRecord, signal: AbortSignal) {
    const saved = await runProvisioning(api, graph, record, signal, () => undefined);
    if (!signal.aborted && saved.status === "completed") {
      setStandard(String(saved.standard.credits)); setEnhanced(String(saved.enhanced.credits));
      setNotice("Allocation verified in Graph. Live details are refreshing.");
      operation.current = null; onApplied();
    }
  }
  return <section className="card"><h2>Allocate the approved monthly budget</h2>
    <p>Ceiling: {format(budget.approvedCreditCeiling)} Copilot Credits (${(budget.approvedCreditCeiling / 100).toLocaleString("en-US")} planning equivalent).
      This application cannot prevent separately authorized changes in the admin center.</p>
    <div className="request-form"><fieldset disabled={busy}><legend>Tier credit limits</legend>
      <label>Standard credits <input type="number" min="0" step="1" value={standard} onChange={(event) => setStandard(event.target.value)} /></label>
      <label>Enhanced credits <input type="number" min="0" step="1" value={enhanced} onChange={(event) => setEnhanced(event.target.value)} /></label>
      <div className="actions">
        <button className="secondary" onClick={() => void perform(loadLimits)}>Refresh actual Graph limits</button>
        <button disabled={!standard || !enhanced} onClick={() => void perform(async (graph, signal) => {
          checkAllocation(Number(standard), Number(enhanced), budget.approvedCreditCeiling!);
          const signature = `${standard}:${enhanced}`;
          if (operation.current?.signature !== signature) operation.current = { signature, id: crypto.randomUUID() };
          const saved = await api.allocation(budget.id, signal, { operationId: operation.current.id, standard: Number(standard), enhanced: Number(enhanced) });
          if (!saved) throw new Error("No allocation operation was returned.");
          await run(graph, saved, signal);
        })}>Apply allocation</button>
        <button className="secondary" onClick={() => void perform(async (graph, signal) => {
          const saved = await api.allocation(budget.id, signal);
          if (!saved || saved.status === "completed") throw new Error("There is no unfinished allocation to resume.");
          await run(graph, saved, signal);
        })}>Resume saved allocation</button>
      </div>
    </fieldset></div>
    {needsConsent && <button disabled={busy} onClick={() => void perform(async (graph, signal) => {
      if (!standard && !enhanced) await loadLimits(graph, signal);
    }, true)}>Grant Graph access</button>}
    {busy && <p role="status">Working with Graph. Keep this page open.</p>}
    {error && <p className="error" role="alert">{error}</p>}
    {notice && <p role="status">{notice}</p>}
  </section>;
}
