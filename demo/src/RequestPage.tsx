import { useState, type FormEvent } from "react";
import { DemoAuth } from "./auth";
import type { DemoConfig, GrantKind } from "./config";
import { createRequestPreview, expiryLabel, resolveRequestRoute, type RequestPreview } from "./requests";

export function RequestPage({ config, auth }: { config: DemoConfig; auth: DemoAuth }) {
  const [account, setAccount] = useState(auth.account());
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [justification, setJustification] = useState("");
  const [grantKind, setGrantKind] = useState<GrantKind>("billingPeriod");
  const [preview, setPreview] = useState<RequestPreview | null>(null);

  function clearForm() {
    setAccount(null);
    setPreview(null);
    setJustification("");
    setGrantKind("billingPeriod");
    setError("");
  }

  async function authenticate(signOut = false) {
    clearForm();
    setBusy(true);
    try {
      if (signOut) await auth.signOut();
      else setAccount(await auth.signIn("requester"));
    } catch (error) {
      setError(error instanceof Error ? error.message : String(error));
    } finally {
      setBusy(false);
    }
  }

  function showPreview(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setError("");
    setPreview(null);
    if (!account) {
      setError("Sign in before previewing a request.");
      return;
    }
    try {
      setPreview(createRequestPreview(config, account, justification, grantKind));
    } catch (error) {
      setError(error instanceof Error ? error.message : String(error));
    }
  }

  if (!account || busy) return <main className="login-shell">
    <section className="card login-card" aria-labelledby="request-login">
      <span className="eyebrow">COPILOT / ALLOWANCE REQUEST</span>
      <h1 id="request-login">Request a higher allowance</h1>
      <p className="muted">Sign in with your work account. You do not need a policy ID or access to cost-management APIs.</p>
      <button onClick={() => void authenticate()} disabled={busy}>{busy ? "Completing sign-in / sign-out..." : "Sign in with Microsoft"}</button>
      {error && <p className="error" role="alert">{error}</p>}
    </section>
  </main>;

  const route = resolveRequestRoute(config, account);
  return <main className="shell request-shell">
    <header className="masthead">
      <div><span className="eyebrow">COWORK / REQUEST AN INCREASE</span><h1>A little more room to work.</h1><p className="muted">Explain the business need. We find the right owner.</p></div>
      <div className="operator"><span className="badge warning">Visual demo only</span><p className="small muted">{account.username}</p>
        <div className="actions"><button className="secondary" onClick={() => void authenticate()}>Switch account</button><button className="secondary" onClick={() => void authenticate(true)}>Sign out</button></div>
      </div>
    </header>
    <div className="notice"><strong>Preview, not submission.</strong> Nothing is sent, saved or added to the owner inbox. This form demonstrates the experience; the inbox uses separate JSON examples.</div>
    <div className="request-layout">
      <section className="card" aria-labelledby="request-heading">
        <span className="eyebrow">01 / BUSINESS NEED</span><h2 id="request-heading">Request the Enhanced allowance</h2>
        <p className="muted">No credit quantity or policy selection is needed. Approval would use your budget's existing Enhanced tier, not change its limits or purchase credits.</p>
        <form onSubmit={showPreview} className="request-form">
          <label htmlFor="justification">Business justification</label>
          <p className="small muted" id="justification-help">Describe the work, business impact and why the additional allowance is needed. 20-2000 characters.</p>
          <textarea id="justification" value={justification} required minLength={20} maxLength={2000} rows={6}
            aria-describedby="justification-help" onChange={(event) => { setJustification(event.target.value); setPreview(null); }} />
          <span className="small muted">{justification.length}/2000 characters</span>
          <fieldset><legend>How long is the increase needed?</legend>
            <label className="radio-choice"><input type="radio" name="grantKind" value="billingPeriod" checked={grantKind === "billingPeriod"}
              onChange={() => { setGrantKind("billingPeriod"); setPreview(null); }} />
              <span><strong>Rest of this billing period</strong><span className="small muted block">For this demo, the period is the UTC calendar month. The preview captures its end.</span></span></label>
            <label className="radio-choice"><input type="radio" name="grantKind" value="permanent" checked={grantKind === "permanent"}
              onChange={() => { setGrantKind("permanent"); setPreview(null); }} />
              <span><strong>Permanent increase</strong><span className="small muted block">An ongoing Enhanced entitlement, subject to future review or revocation.</span></span></label>
          </fieldset>
          {error && <p className="error" role="alert">{error}</p>}
          <button type="submit">Preview request and routing</button>
        </form>
      </section>
      <aside className="card" aria-labelledby="routing-heading">
        <span className="eyebrow">02 / ROUTING</span><h2 id="routing-heading">Matched from your identity</h2>
        <p className="muted">Your signed-in username is looked up in the demo index. You never need to identify a spending policy.</p>
        <dl><dt>Requester</dt><dd>{account.username}</dd><dt>Service</dt><dd>{config.serviceId}</dd>
          {route.budget && <><dt>Budget</dt><dd>{route.budget.displayName}</dd></>}
          <dt>Routing status</dt><dd>{route.state === "ready" ? "Owner found / preview only" : "Needs routing"}</dd>
          {route.state === "ready" && <><dt>Current eligible owners</dt><dd>{route.approvers.join(", ")}</dd></>}
        </dl>
        {route.state === "needs-routing" && <p className="reading-note" role="status">{route.reason}</p>}
        <div className="routing-flow"><span>Signed-in user</span><span aria-hidden="true">↓</span><span>JSON user-to-budget index</span><span aria-hidden="true">↓</span><span>Current budget owner / co-owners</span></div>
        <p className="small muted">No balance, policy, group or membership API is called by this page. An index is a routing hint, not proof of an active spending policy.</p>
      </aside>
    </div>
    {preview && <section className="card request-preview" aria-labelledby="preview-heading" role="status">
      <span className="eyebrow">03 / REQUEST PREVIEW</span><h2 id="preview-heading">{preview.route.state === "ready" ? "This is what the owner would receive" : "This would need administrative routing"}</h2>
      <dl className="metadata-grid">
        <div><dt>From</dt><dd>{preview.requester}</dd></div>
        <div><dt>Requested tier</dt><dd>Enhanced</dd></div>
        <div><dt>Duration</dt><dd>{expiryLabel(preview.expiresAt)}</dd></div>
        <div><dt>Routing</dt><dd>{preview.route.state === "ready" ? preview.route.approvers.join(", ") : preview.route.reason}</dd></div>
      </dl>
      <h3>Business justification</h3><p className="justification">{preview.justification}</p>
      <p className="reading-note">Not submitted, not approved, and no membership or allowance has changed. Reloading clears this preview. Temporary reversion is not scheduled by this demo.</p>
    </section>}
    <footer>Evaluation only. JSON routing is not secure authorization or a live submission service. <a href={import.meta.env.BASE_URL}>Owner workspace</a></footer>
  </main>;
}
