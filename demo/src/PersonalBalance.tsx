import { useEffect, useMemo, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import { DemoAuth, isGraphInteractionRequired } from "./auth";
import { BalanceDetails } from "./BudgetDetails";
import { formatField, loadSpendingPolicy, loadUserBalance, type Entity, type Fields, type Resource } from "./details";
import { GraphClient } from "./graph";

export interface PersonalReading {
  balance: Fields | null;
  fetchedAt: string;
  policy: Resource<Entity> | null;
  interactionRequired: boolean;
}
const message = (error: unknown) => error instanceof Error ? error.message : String(error);

export async function loadPersonalBalance(
  graph: GraphClient, userId: string, serviceId: string, signal: AbortSignal,
): Promise<PersonalReading> {
  const balance = await loadUserBalance(graph, userId, serviceId, signal);
  signal.throwIfAborted();
  const reading: PersonalReading = { balance, fetchedAt: new Date().toISOString(), policy: null, interactionRequired: false };
  const id = balance?.spendingPolicyId;
  if (id !== undefined && id !== null && id !== "") {
    try {
      const data = await loadSpendingPolicy(graph, id, signal);
      reading.policy = { state: "ready", data, fetchedAt: new Date().toISOString() };
    } catch (error) {
      signal.throwIfAborted();
      reading.policy = { state: "error", error: message(error) };
      reading.interactionRequired = isGraphInteractionRequired(error);
    }
  }
  signal.throwIfAborted();
  return reading;
}

export function PersonalBalanceReading({ reading, serviceId }: { reading: PersonalReading; serviceId: string }) {
  const policy = reading.policy?.state === "ready" ? reading.policy.data.fields : null;
  return reading.balance ? <>
    <BalanceDetails row={reading.balance} fetchedAt={reading.fetchedAt} showLimits
      policyName={policy ? formatField(policy.displayName) : reading.policy ? "Policy details unavailable" : "Not reported"} note={null} />
    {policy && <p className="small muted">Policy status: {formatField(policy.status)}. Billing methods: {formatField(policy.billingMethods)}.</p>}
    {reading.policy?.state === "error" && <p className="error" role="alert">Could not load the effective policy details. {reading.policy.error}</p>}
    {!reading.policy && <p className="reading-note">Graph did not report an effective spending policy. No policy or owner is inferred.</p>}
  </> : <p className="reading-note">No {serviceId} balance was returned. Consumption is unavailable, not zero.</p>;
}

export function PersonalBalance({ auth, account, serviceId }: { auth: DemoAuth; account: AccountInfo; serviceId: string }) {
  const graph = useMemo(() => new GraphClient(() => auth.balanceToken(account)), [auth, account]);
  const [reading, setReading] = useState<Resource<PersonalReading>>({ state: "loading" });
  const [needsConsent, setNeedsConsent] = useState(false);
  const [consenting, setConsenting] = useState(false);
  const [consentError, setConsentError] = useState("");
  const [refresh, setRefresh] = useState(0);
  const controller = useRef<AbortController | null>(null);
  useEffect(() => {
    const current = new AbortController();
    controller.current = current;
    setReading({ state: "loading" }); setNeedsConsent(false); setConsentError("");
    void loadPersonalBalance(graph, account.localAccountId, serviceId, current.signal).then((data) => {
      if (!current.signal.aborted) {
        setReading({ state: "ready", data, fetchedAt: data.fetchedAt });
        setNeedsConsent(data.interactionRequired);
      }
    }).catch((error: unknown) => {
      if (!current.signal.aborted) {
        setReading({ state: "error", error: message(error) });
        setNeedsConsent(isGraphInteractionRequired(error));
      }
    });
    return () => current.abort();
  }, [graph, account.localAccountId, serviceId, refresh]);
  async function consent() {
    const current = controller.current;
    setConsenting(true); setConsentError("");
    try {
      await auth.consentBalance(account);
      if (!current?.signal.aborted) setRefresh((value) => value + 1);
    } catch (error) {
      if (!current?.signal.aborted) setConsentError(message(error));
    } finally {
      if (!current?.signal.aborted) setConsenting(false);
    }
  }
  const busy = reading.state === "loading" || consenting;
  return <section className="card personal-balance" aria-labelledby="personal-balance-heading">
    <div className="section-heading"><div><span className="eyebrow">{serviceId} / CURRENT GRAPH READING</span>
      <h2 id="personal-balance-heading">Your current usage and balance</h2></div>
      <button className="secondary" disabled={busy} onClick={() => setRefresh((value) => value + 1)}>Refresh my balance</button>
    </div>
    <p className="small muted">Current credit consumption and allowance, not invoiced spend. Graph readings may lag behind usage.
      This demo assumes your account can read its balance and effective policy; those permissions are not granted by this screen.</p>
    {reading.state === "loading" && <p role="status">Loading your balance from Graph...</p>}
    {reading.state === "error" && <p className="error" role="alert">Your balance is unavailable. {reading.error}</p>}
    {reading.state === "ready" && <PersonalBalanceReading reading={reading.data} serviceId={serviceId} />}
    {needsConsent && <button disabled={busy} onClick={() => void consent()}>{consenting ? "Authorizing..." : "Grant balance read access"}</button>}
    {consentError && <p className="error" role="alert">{consentError}</p>}
    <p className="small muted">Balance reads are optional: you can still submit a request when server routing is ready.
      The effective policy shown here does not override your verified request route below.</p>
  </section>;
}
