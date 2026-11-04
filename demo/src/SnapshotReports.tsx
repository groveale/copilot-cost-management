import { useEffect, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import { ApiRequestError, type RequestApi } from "./api";
import { DemoAuth } from "./auth";
import { formatField, formatQuantity, formatTime } from "./details";
import { GraphClient, GraphError } from "./graph";
import { SharePointReporting, type SnapshotPolicy, type SnapshotReading } from "./reporting";
import { parallelMap } from "./sync-collector";

type Report = { policy: SnapshotPolicy; observations: SnapshotReading["observations"] | null };
export async function loadSnapshotReports(api: RequestApi, graph: GraphClient, tenantId: string, signal: AbortSignal,
  options: { policyIds?: string[]; inventoryOnly?: boolean; userId?: string } = {}): Promise<Report[]> {
  const settings = await api.reportingSettings(signal);
  if (options.userId && !settings.isAdmin) throw new Error("Personal synced balances are available to administrators only in v1.");
  const reader = new SharePointReporting(graph, settings, tenantId);
  const selected = (reading: SnapshotReading): Report => ({ ...reading,
    observations: options.userId ? reading.observations.filter((row) => row.userId === options.userId) : reading.observations });
  if (options.policyIds) return parallelMap(options.policyIds, 3, async (id) => selected(await reader.read(id, signal)));
  const inventory = await reader.inventory(signal);
  if (options.inventoryOnly) return inventory.map((policy) => ({ policy, observations: null }));
  return parallelMap(inventory.map((policy) => policy.id), 3, async (id) => {
    return selected(await reader.read(id, signal));
  });
}

export function SnapshotPolicyPanel({ report, onSync, personal = false }: { report: Report; onSync?: (id: string) => void; personal?: boolean }) {
  const { policy, observations } = report;
  const coverage = policy.coverage;
  return <article className="card snapshot-policy">
    <span className="eyebrow">SHAREPOINT SNAPSHOT</span><h3>{policy.displayName}</h3>
    <p className="mono small">{policy.id}</p>
    {onSync && <button className="secondary" onClick={() => onSync(policy.id)}>Sync this policy</button>}
    {!policy.activeSnapshotId ? <p className="notice">No successful snapshot has been published for this policy. Consumption is unavailable, not zero.</p> : <>
      <p className="small">Collected {formatTime(policy.startedAt)} to {formatTime(policy.observedAt)}. This is not a real-time reading.</p>
      {policy.stale && <p className="notice">Stale policy: absent from the last full inventory. The previous successful snapshot is retained.</p>}
      {!!coverage?.exclusions && <p className="notice">{coverage.exclusions} observations or entities were excluded; this report is incomplete. Mismatched balances are not attributed to this policy.</p>}
      {coverage?.unsupported ? <p className="notice">Inventory only: {coverage.reason}. Consumption was not collected.</p> : !personal && <>
        <p>{coverage?.users} discovered users across {coverage?.groups} assigned groups. Membership includes nested users, deduplicated within this policy.</p>
        <dl className="metadata-grid">{policy.totals?.map((total) => <div key={`${total.serviceId}:${total.unit}`}>
          <dt>{total.serviceId} / matched consumption</dt><dd>{total.consumed.toLocaleString()} {total.unit} ({total.users} readings)</dd>
        </div>)}</dl>
        {policy.totals?.length === 0 && <p>No policy-matched consumption was observed. This does not imply a zero allowance or a zero invoice.</p>}
      </>}
      {personal && <p className="small">Only your stored observations are listed below. Policy-wide consumption totals are not shown.</p>}
      <details><summary>Stored configuration, services and ownership</summary>
        <dl><dt>Period limit</dt><dd>{formatQuantity(policy.configuration?.policy.periodLimit)}</dd>
          <dt>Per-user limit</dt><dd>{formatQuantity(policy.configuration?.policy.perUserLimit)}</dd></dl>
        <p className="small muted">Configuration and owner mappings as collected, not current authorization. Services omitted from the snapshot are not assumed disabled.</p>
        <h4>Assigned groups</h4>
        {!policy.configuration?.groups.length ? <p>No assigned groups were captured.</p> :
          <ul>{policy.configuration.groups.map((group) => <li key={group.id}>{group.displayName ?? "Name unavailable"} <span className="mono">{group.id}</span></li>)}</ul>}
        <h4>Portal mapping</h4>
        {!policy.configuration?.mappings.length ? <p>Unmapped at collection time; administrator-only reporting.</p> :
          <ul>{policy.configuration.mappings.map((mapping, index) => <li key={index}>
            {formatField(mapping.budgetId)} / {formatField(mapping.tier)} / {formatField(mapping.businessArea)} / {formatField(mapping.costCenter)}
            <span className="small block">Owner IDs: {formatField(mapping.owners)}</span>
          </li>)}</ul>}
        <pre>{JSON.stringify(policy.configuration, null, 2)}</pre>
      </details>
      {observations !== null && <div className="table-scroll" tabIndex={0} role="region" aria-label={`${policy.displayName} synced user balances`}>
        <table><caption>Stored user/service observations for this policy and collection time. These are normalized snapshots, not full Graph responses.</caption>
          <thead><tr><th>User / discovered groups</th><th>Service</th><th>Consumed</th><th>Remaining</th><th>Access / reported policy</th><th>Time</th></tr></thead>
          <tbody>{observations.map((row) => <tr key={`${row.userId}:${row.serviceId}`}>
            <td><strong>{row.displayName ?? row.userId}</strong><span className="small block">{row.username}</span>
              <details><summary className="small">Discovered group IDs</summary>{row.groups.map((id) => <span className="mono block" key={id}>{id}</span>)}</details></td>
            <td>{row.serviceId ?? "No balance returned"}</td>
            <td>{row.consumed === null ? "Unavailable, not zero" : `${row.consumed.toLocaleString()} ${row.unit}`}
              {row.match === "mismatch" && <span className="small block">Excluded from this policy's totals</span>}</td>
            <td>{row.remainingKind === "unlimited" ? "Unlimited" : row.remainingKind === "limited" ? `${row.remaining?.toLocaleString()} ${row.unit}` : "Not reported"}</td>
            <td>{row.allowed === null ? "Not reported" : row.allowed ? "Allowed" : "Not allowed"}
              <span className="small block">{row.reason}</span><span className="mono block">{row.reportedPolicyId ?? "Not reported"}</span></td>
            <td><span className="small block">Observed: {formatTime(row.observedAt)}</span>
              <span className="small block">Source: {formatTime(row.sourceTime)}</span></td>
          </tr>)}</tbody>
        </table>
        {!observations.length && <p>No stored user observations for this selection. Consumption is unavailable, not zero.</p>}
        <p className="small muted">Per-user configured limits and effective capacity were not captured by sync and are unavailable here.</p>
      </div>}
    </>}
  </article>;
}

export function SnapshotReports({ api, auth, account, policyIds, inventoryOnly, userId, onSync }: {
  api: RequestApi; auth: DemoAuth; account: AccountInfo; policyIds?: string[]; inventoryOnly?: boolean;
  userId?: string; onSync?: (id: string) => void;
}) {
  const [reports, setReports] = useState<Report[] | null>(null);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const challenges = useRef<{ api?: string; graph?: string }>({});
  const controller = useRef<AbortController | null>(null);
  const selection = policyIds?.join(",");
  useEffect(() => {
    void load();
    return () => controller.current?.abort();
  }, [api, auth, account, selection, inventoryOnly, userId]);
  async function load(interactive = false) {
    controller.current?.abort();
    const current = new AbortController(); controller.current = current;
    setReports(null); setBusy(true); setError("");
    try {
      if (interactive) {
        await auth.refreshApi(account, challenges.current.api);
        current.signal.throwIfAborted();
        await auth.refreshReporting(account, challenges.current.graph);
      }
      current.signal.throwIfAborted();
      const reports = await loadSnapshotReports(api, new GraphClient(() => auth.reportingToken(account)),
        account.tenantId, current.signal, { policyIds, inventoryOnly, userId });
      if (!current.signal.aborted) { setReports(reports); challenges.current = {}; }
    } catch (error) {
      if (!current.signal.aborted) {
        setError(error instanceof Error ? error.message : String(error));
        if (error instanceof ApiRequestError && error.claims) challenges.current.api = error.claims;
        if (error instanceof GraphError && error.claims) challenges.current.graph = error.claims;
      }
    } finally { if (!current.signal.aborted) setBusy(false); }
  }
  return <section aria-label="SharePoint reporting">
    <div className="card"><h2>{userId ? "Your synced usage and balance" : inventoryOnly ? "Synced spending policies" : "Synced policy details"}</h2>
      <p className="muted">Delegated SharePoint reporting. No live Cost Management calls are made for these readings.
        Actions use live APIs; snapshots remain unchanged until the next successful sync.</p>
      <div className="actions">
        <button disabled={busy} onClick={() => void load()}>Refresh snapshots</button>
        {error && <button disabled={busy} onClick={() => void load(true)}>Refresh sign-in / SharePoint access</button>}
      </div>
      {busy && <p role="status">Reading SharePoint snapshots...</p>}
      {error && <p className="error" role="alert">{error}</p>}
      {reports?.length === 0 && <p className="notice">No SharePoint snapshots are available to this account. Ask an administrator to sync and verify permissions.</p>}
      {reports && reports.length > 0 && <p className="small">{reports.length} SharePoint policy records in this selection.</p>}
      {userId && reports && <p className="small">Admin-only personal reporting, shown separately by discovery policy and observation time.
        Repeated users across policies are not summed. Request routing is unaffected.</p>}
    </div>
    {reports?.map((report) => <SnapshotPolicyPanel key={report.policy.id} report={report} onSync={onSync} personal={!!userId} />)}
  </section>;
}
