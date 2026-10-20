import { useEffect, useMemo, useRef, useState } from "react";
import { DemoAuth, scopes } from "./auth";
import { budgetsForOwner, type DemoConfig } from "./config";
import { GraphClient } from "./graph";
import { loadBudgetDetails, stopLoading, type BudgetSnapshot } from "./details";
import { BudgetDetails } from "./BudgetDetails";
import { ExampleRequests } from "./ExampleRequests";

const errorMessage = (error: unknown) => error instanceof Error ? error.message : String(error);

export function BudgetPortal({ config, auth }: { config: DemoConfig; auth: DemoAuth }) {
  const [account, setAccount] = useState(auth.account());
  const [authBusy, setAuthBusy] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [budgetId, setBudgetId] = useState("");
  const [refresh, setRefresh] = useState(0);
  const [snapshot, setSnapshot] = useState<BudgetSnapshot | null>(null);
  const controller = useRef<AbortController | null>(null);
  const graph = useMemo(() => new GraphClient(() => auth.token(account)), [auth, account]);
  const ownedBudgets = budgetsForOwner(config, account);
  const budget = ownedBudgets.find((item) => item.id === budgetId) ?? ownedBudgets[0];
  const redirectUri = new URL("auth.html", new URL(import.meta.env.BASE_URL, window.location.origin)).href;

  useEffect(() => {
    setSnapshot(null);
    setLoading(false);
    if (!account || authBusy || !budget) return;
    const current = new AbortController();
    controller.current = current;
    setLoading(true);
    setError("");
    void loadBudgetDetails(budget, config.serviceId, graph, current.signal, (next) => {
      if (controller.current === current && !current.signal.aborted) setSnapshot(next);
    })
      .catch((error: unknown) => {
        if (controller.current !== current) return;
        const message = current.signal.aborted ? "Loading cancelled. Already loaded details remain visible." : errorMessage(error);
        setError(message);
        setSnapshot((previous) => previous ? stopLoading(previous, message) : null);
      })
      .finally(() => {
        if (controller.current === current) {
          controller.current = null;
          setLoading(false);
        }
      });
    return () => {
      controller.current = null;
      current.abort();
    };
  }, [account, authBusy, budget, config.serviceId, graph, refresh]);

  function clearLoad() {
    const current = controller.current;
    controller.current = null;
    current?.abort();
    setLoading(false);
    setSnapshot(null);
  }

  async function signIn() {
    clearLoad();
    setAccount(null);
    setBudgetId("");
    setAuthBusy(true);
    setError("");
    setSnapshot(null);
    try {
      setAccount(await auth.signIn());
      setRefresh((value) => value + 1);
    } catch (error) {
      setError(errorMessage(error));
    } finally {
      setAuthBusy(false);
    }
  }

  async function signOut() {
    clearLoad();
    setAccount(null);
    setBudgetId("");
    setAuthBusy(true);
    setError("");
    setSnapshot(null);
    try {
      await auth.signOut();
    } catch (error) {
      setError(errorMessage(error));
    } finally {
      setAuthBusy(false);
    }
  }

  if (!account || authBusy) {
    return <main className="login-shell">
      <section className="card login-card" aria-labelledby="login-heading">
        <span className="eyebrow">COPILOT BUDGET PORTAL</span>
        <h1 id="login-heading">Sign in to your budgets</h1>
        <p className="muted">Use your work account to view the budgets you manage.</p>
        <button onClick={signIn} disabled={authBusy}>{authBusy ? "Completing sign-in / sign-out..." : "Sign in with Microsoft"}</button>
        {error && <p className="error" role="alert">{error}</p>}
      </section>
    </main>;
  }

  return (
    <main className="shell">
      <header className="masthead">
        <div><span className="eyebrow">COPILOT BUDGET PORTAL</span><h1>{budget?.displayName ?? "Your budgets"}</h1><p className="muted">Policy configuration, groups, people and current consumption.</p></div>
        <div className="operator"><span className="badge warning">Evaluation / read only</span><p className="small muted">{account?.username ?? "Not signed in"}</p>
          <div className="actions"><button onClick={signIn} disabled={authBusy}>{account ? "Sign in / renew consent" : "Sign in with Microsoft"}</button>
            {account && <button className="secondary" onClick={signOut} disabled={authBusy}>Sign out</button>}</div>
        </div>
      </header>
      <div className="notice"><strong>Your owner workspace.</strong> Your signed-in account is matched to the owners in the demo JSON. Only matching budgets are shown and loaded. This simulates a database lookup, not server-enforced access control. No live writes or consumption history.</div>
      {!budget ? <section className="card budget-summary" aria-labelledby="empty-heading">
        <h2 id="empty-heading">No budgets assigned</h2>
        <p>No owner entries match your signed-in work account. Contact the demo administrator to update your owner assignment.</p>
        <p className="small muted">No budget policies, groups or consumption have been requested from Graph.</p>
        {error && <p className="error" role="alert">{error}</p>}
      </section> : <>
      <section className="card budget-summary" aria-labelledby="budget-heading">
        <div className="section-heading">
          <div><span className="eyebrow">BUDGET / JSON METADATA</span><h2 id="budget-heading">{budget?.displayName ?? "No budgets configured"}</h2></div>
          <div className="actions">
            {ownedBudgets.length > 1 && <label className="budget-select">Budget<select value={budget.id} onChange={(event) => { clearLoad(); setBudgetId(event.target.value); }} disabled={authBusy}>
              {ownedBudgets.map((item) => <option value={item.id} key={item.id}>{item.displayName}</option>)}
            </select></label>}
            <button onClick={() => setRefresh((value) => value + 1)} disabled={!account || !budget || authBusy || loading}>{loading ? "Loading from Graph..." : "Refresh Graph details"}</button>
            {loading && <button className="secondary" onClick={() => controller.current?.abort()}>Cancel loading</button>}
          </div>
        </div>
        {budget && <dl className="metadata-grid">
          <div><dt>Business area</dt><dd>{budget.businessArea}</dd></div>
          <div><dt>Cost centre</dt><dd>{budget.costCenter}</dd></div>
          <div><dt>Budget owners</dt><dd>{budget.owners.join(", ")}</dd></div>
          {budget.plannedUsd !== undefined && <div><dt>Planned USD / metadata only</dt><dd>{budget.plannedUsd.toLocaleString("en-US", { style: "currency", currency: "USD" })}</dd></div>}
          {budget.periodLabel && <div><dt>Business period label</dt><dd>{budget.periodLabel}</dd></div>}
        </dl>}
        {error && <p className="error" role="alert">{error}</p>}
      </section>
      <ExampleRequests key={`${account.homeAccountId}:${budget.id}`} config={config} budget={budget} account={account} />
      {budget && snapshot?.budgetId === budget.id && <BudgetDetails budget={budget} snapshot={snapshot} serviceId={config.serviceId} />}
      </>}
      <section className="card connection-details">
        <details><summary>Connection and static configuration</summary>
          <dl><dt>Tenant</dt><dd className="mono">{config.tenantId}</dd><dt>Application</dt><dd className="mono">{config.clientId}</dd><dt>SPA redirect URI</dt><dd className="mono">{redirectUri}</dd></dl>
          <p className="small muted">All business data outside Graph lives in public\demo-config.json. Edit it manually and reload; no metadata service or application database exists.</p>
        </details>
        <details><summary>Requested delegated permissions</summary>
          <ul className="small">
            {scopes.map((scope) => <li key={scope}>{scope.replace("https://graph.microsoft.com/", "")}</li>)}
          </ul>
          <p className="small muted">These operator-selected delegated scopes include broad directory access and policy/assignment write permissions. This view makes read-only requests. No application permissions are requested, and no access tokens are displayed or exported.</p>
        </details>
      </section>
      <footer>Current user/service readings only, not reconciled billing or historical budget spend. Preview APIs are not supported for production applications.</footer>
    </main>
  );
}
