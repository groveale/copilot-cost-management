import { useEffect, useId, useRef, useState, type FormEvent } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import type { BudgetConfig, TierName } from "./config";
import { RequestApi } from "./api";
import { DemoAuth, isGraphInteractionRequired } from "./auth";
import { GraphClient } from "./graph";
import { addPolicyMember, findDirectoryUser, type DirectoryUser } from "./membership";

export function AddPolicyMember({ api, auth, account, budget, tier, serviceId, disabled, onBusyChange, onAdded }: {
  api: RequestApi; auth: DemoAuth; account: AccountInfo; budget: BudgetConfig; tier: TierName; serviceId: string;
  disabled: boolean; onBusyChange: (busy: boolean) => void; onAdded: (user: DirectoryUser) => void;
}) {
  const [expanded, setExpanded] = useState(false);
  const detailsId = useId();
  const [upn, setUpn] = useState("");
  const [user, setUser] = useState<DirectoryUser | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [needsConsent, setNeedsConsent] = useState(false);
  const controller = useRef<AbortController | null>(null);
  useEffect(() => () => controller.current?.abort(), []);
  const label = tier === "standard" ? "Standard" : "Enhanced";
  async function perform(action: (graph: GraphClient, signal: AbortSignal) => Promise<void>, interactive = false) {
    controller.current?.abort();
    const current = new AbortController(); controller.current = current;
    setBusy(true); onBusyChange(true); setError("");
    try {
      if (interactive) await auth.consentGraph(account);
      await auth.token(account);
      setNeedsConsent(false);
      await action(new GraphClient(() => auth.token(account)), current.signal);
    } catch (error) {
      if (!current.signal.aborted) {
        setNeedsConsent(isGraphInteractionRequired(error));
        setError(`${error instanceof Error ? error.message : String(error)} If an add was interrupted, retry verifies the actual membership before writing again.`);
      }
    } finally {
      if (!current.signal.aborted) setBusy(false);
      onBusyChange(false);
    }
  }
  function lookup(event: FormEvent<HTMLFormElement>) {
    event.preventDefault(); setUser(null);
    void perform(async (graph, signal) => {
      const found = await findDirectoryUser(graph, upn, signal);
      if (!signal.aborted) setUser(found);
    });
  }
  return <div className="policy-member-editor">
    <button type="button" className="secondary" disabled={disabled || busy} aria-expanded={expanded} aria-controls={detailsId}
      aria-label={`Add member to ${label} policy`} onClick={() => setExpanded((value) => !value)}>{expanded ? "Close add member" : "Add member"}</button>
    <div id={detailsId} hidden={!expanded}>
      <p className="small muted">Find an existing directory user by their exact sign-in email, then confirm the addition.
        People already in the other tier cannot be added here. No accounts are created or moved.</p>
      <form className="policy-member-form" onSubmit={lookup}>
        <label>Sign-in email (UPN)<input type="email" required maxLength={254} disabled={disabled || busy} value={upn}
          onChange={(event) => { setUpn(event.target.value); setUser(null); }} /></label>
        <button type="submit" disabled={disabled || busy || !upn.trim()}>Find user</button>
      </form>
      {user && <div className="member-confirmation"><p><strong>{user.displayName}</strong><span className="block small muted">{user.userPrincipalName}</span></p>
        <button type="button" disabled={disabled || busy} onClick={() => void perform(async (graph, signal) => {
          await addPolicyMember(api, graph, budget, tier, user, serviceId, signal);
          if (!signal.aborted) onAdded(user);
        })}>Add {user.displayName} to {label}</button></div>}
    </div>
    {needsConsent && <button type="button" disabled={disabled || busy} onClick={() => void perform(async () => undefined, true)}>Grant Graph access</button>}
    {busy && <p className="small muted" role="status">Checking directory membership. Keep this page open.</p>}
    {error && <p className="error" role="alert">{error}</p>}
  </div>;
}
