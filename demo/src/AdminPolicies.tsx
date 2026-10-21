import { useEffect, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import type { Budget, Metadata } from "../../api/src/contracts";
import { RequestApi } from "./api";
import { DemoAuth, isGraphInteractionRequired } from "./auth";
import { guid, object, text, type TierName } from "./config";
import { tierNames, type Resource } from "./details";
import { GraphClient } from "./graph";

interface PolicyMapping { budget: Budget; tier: TierName }
interface PolicyRow { id: string; displayName: string; mappings: PolicyMapping[] }
interface DirectoryOwner { displayName: string; username?: string }
export interface PolicyInventory {
  policies: PolicyRow[];
  owners: Record<string, Resource<DirectoryOwner>>;
  initialized: boolean;
}
const message = (error: unknown) => error instanceof Error ? error.message : String(error);

export async function loadPolicyInventory(graph: GraphClient, metadata: Metadata | null, signal: AbortSignal): Promise<PolicyInventory> {
  const mappings = new Map<string, PolicyMapping[]>();
  for (const budget of metadata?.budgets ?? []) {
    for (const tier of tierNames) {
      const id = guid(budget.tiers[tier].policyId, "Mapped policy ID");
      mappings.set(id, [...mappings.get(id) ?? [], { budget, tier }]);
    }
  }
  const rows = await graph.list("/beta/copilot/costManagement/spendingPolicies", signal);
  const policies = new Map<string, PolicyRow>();
  for (const value of rows) {
    const fields = object(value, "Spending policy");
    const id = guid(fields.id, "Spending policy ID");
    policies.set(id, { id, displayName: text(fields.displayName, "Spending policy name"), mappings: mappings.get(id) ?? [] });
  }
  const ownerIds = new Set([...policies.values()].flatMap((policy) => policy.mappings.flatMap((mapping) => mapping.budget.owners)));
  const owners: PolicyInventory["owners"] = {};
  for (const id of ownerIds) {
    try {
      const user = object(await graph.get(`/v1.0/users/${guid(id, "Mapped owner ID")}?$select=id,displayName,userPrincipalName`, signal), "Mapped owner");
      if (guid(user.id, "Returned owner ID") !== id.toLowerCase()) throw new Error("Directory lookup returned a different owner ID.");
      const displayName = text(user.displayName, "Owner display name");
      owners[id] = { state: "ready", data: { displayName, username: typeof user.userPrincipalName === "string" ? user.userPrincipalName : undefined },
        fetchedAt: new Date().toISOString() };
    } catch (error) {
      signal.throwIfAborted();
      owners[id] = { state: "error", error: message(error) };
    }
  }
  signal.throwIfAborted();
  return { policies: [...policies.values()].sort((a, b) => a.displayName.localeCompare(b.displayName)),
    owners, initialized: metadata !== null };
}

export function PolicyInventoryTable({ inventory }: { inventory: PolicyInventory }) {
  const unmapped = inventory.policies.filter((policy) => !policy.mappings.length).length;
  return <>
    {!inventory.initialized && <p className="notice">Initial private metadata has not been imported. Graph policies are listed, but portal ownership is not configured.</p>}
    <p role="status">{inventory.policies.length} Graph policies. {unmapped} unmapped.</p>
    {!inventory.policies.length ? <p>No spending policies were returned by Graph.</p> :
      <div className="table-scroll" role="region" aria-label="Spending policies and owner mappings" tabIndex={0}>
        <table className="policy-inventory"><caption>All spending policies returned by Graph; ownership and tiers come from private portal metadata.</caption>
          <thead><tr><th scope="col">Spending policy</th><th scope="col">Mapping</th><th scope="col">Tier</th>
            <th scope="col">Business unit / budget</th><th scope="col">Cost centre</th><th scope="col">Mapped owner(s)</th></tr></thead>
          <tbody>{inventory.policies.map((policy) => <tr key={policy.id}>
            <td><strong>{policy.displayName}</strong><span className="mono block">{policy.id}</span></td>
            <td><span className={`badge ${policy.mappings.length ? "passed" : "warning"}`}>{policy.mappings.length ? "Mapped" : "Unmapped"}</span></td>
            <td>{policy.mappings.length ? policy.mappings.map(({ budget, tier }) => <div key={`${budget.id}:${tier}`}>{tier}</div>) : "Not mapped"}</td>
            <td>{policy.mappings.length ? policy.mappings.map(({ budget, tier }) => <div key={`${budget.id}:${tier}`}>
              {budget.displayName}<span className="small muted block">{budget.businessArea}</span></div>) : "Unmapped"}</td>
            <td>{policy.mappings.length ? policy.mappings.map(({ budget, tier }) => <div key={`${budget.id}:${tier}`}>{budget.costCenter}</div>) : "Not mapped"}</td>
            <td>{policy.mappings.length ? [...new Set(policy.mappings.flatMap(({ budget }) => budget.owners))].map((id) => {
              const owner = inventory.owners[id];
              return <div className="mapped-owner" key={id}>
                {owner?.state === "ready" ? <><strong>{owner.data.displayName}</strong>
                  {owner.data.username && <span className="small block">{owner.data.username}</span>}</> : <strong>Owner name unavailable</strong>}
                <span className="mono block">{id}</span>
                {owner?.state === "error" && <span className="small error block" role="alert">{owner.error}</span>}
              </div>;
            }) : "Unmapped"}</td>
          </tr>)}</tbody>
        </table>
      </div>}
  </>;
}

export function AdminPolicies({ api, auth, account }: { api: RequestApi; auth: DemoAuth; account: AccountInfo }) {
  const [inventory, setInventory] = useState<PolicyInventory | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [needsConsent, setNeedsConsent] = useState(false);
  const controller = useRef<AbortController | null>(null);
  useEffect(() => {
    void load();
    return () => controller.current?.abort();
  }, [api, auth, account]);
  async function load(interactive = false) {
    controller.current?.abort();
    const current = new AbortController(); controller.current = current;
    setBusy(true); setError(""); setInventory(null);
    try {
      if (interactive) await auth.consentGraph(account);
      await auth.token(account);
      setNeedsConsent(false);
      const metadata = await api.onboardingSetup(current.signal);
      const result = await loadPolicyInventory(new GraphClient(() => auth.token(account)), metadata, current.signal);
      if (!current.signal.aborted) setInventory(result);
    } catch (error) {
      if (!current.signal.aborted) { setNeedsConsent(isGraphInteractionRequired(error)); setError(message(error)); }
    }
    finally { if (!current.signal.aborted) setBusy(false); }
  }
  return <section className="card"><h2>Spending policies and owners</h2>
    <p>Read-only inventory of all policies available through Graph, including policies with no portal mapping.
      Mapped owners are portal budget owners, not inferred from Entra group ownership.</p>
    <div className="actions">
      <button disabled={busy} onClick={() => void load()}>{busy ? "Loading policies..." : "Refresh policy list"}</button>
      {needsConsent && <button disabled={busy} onClick={() => void load(true)}>Grant Graph access</button>}
    </div>
    {busy && <p role="status">Loading all policy pages and resolving mapped owner names...</p>}
    <p className="small muted">Uses your cached Graph sign-in, renewing silently when possible. No policy or owner mapping is changed.</p>
    {error && <p className="error" role="alert">{error}</p>}
    {inventory && <PolicyInventoryTable inventory={inventory} />}
  </section>;
}
