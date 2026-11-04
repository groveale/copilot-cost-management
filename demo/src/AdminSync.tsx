import { useEffect, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import type { SyncJob, SyncSettings } from "../../api/src/sync-contracts";
import { reportingAccessMode, type ReportingAccessMode } from "../../api/src/reporting-contracts";
import { ApiRequestError, RequestApi } from "./api";
import { DemoAuth } from "./auth";
import { guid } from "./config";
import { GraphError } from "./graph";
import { mirrorSyncJob, runSync, syncClients, validateSyncAccess } from "./sync-runner";
import { useOperationLock } from "./OperationLock";
import { ReportingAccessNotice } from "./ReportingAccessNotice";

const message = (error: unknown) => error instanceof Error ? error.message : String(error);
const terminal = (job: SyncJob) => ["completed", "completed-with-exclusions", "partially-failed", "failed", "cancelled"].includes(job.state);
export function SyncJobSummary({ job }: { job: SyncJob }) {
  return <section aria-label={`Sync job ${job.id}`}>
    <h3>Sync {job.policyId ? "one policy" : "all policies"}</h3>
    <p><strong>{job.state}</strong> - {job.phase}</p>
    <p className="small">Job <span className="mono">{job.id}</span> | Concurrency {job.concurrency} | Updated {job.updatedAt}</p>
    <p className="small">Initiated by {job.initiatorId}. Current executor: {job.executorId ?? "not claimed"}.
      {job.actors.length > 1 && ` ${job.actors.length} administrators have participated.`}</p>
    {job.lease && <p className="small">Execution lease expires {job.lease.expiresAt}. A different browser must wait for expiry or release.</p>}
    {job.mirrorPending && <p className="notice">SharePoint job reporting needs reconciliation. Private execution state below is authoritative.</p>}
    <p>{job.policies.filter((policy) => ["completed", "excluded"].includes(policy.state)).length} / {job.policies.length} policies complete.</p>
    {!!job.policies.length && <div className="table-scroll"><table><caption>Per-policy sync outcomes</caption>
      <thead><tr><th>Policy</th><th>Status</th><th>Unique users</th><th>Exclusions / issues</th></tr></thead>
      <tbody>{job.policies.map((policy) => <tr key={policy.id}><td>{policy.displayName}<span className="mono block">{policy.id}</span></td>
        <td>{policy.state}</td><td>{policy.users ?? "Not collected"}</td>
        <td>{policy.exclusions ?? "Not collected"}{policy.error && <span className="block error">{policy.error}</span>}</td></tr>)}</tbody>
    </table></div>}
  </section>;
}
export function AdminSync({ api, auth, account, selectedPolicyId = "" }: {
  api: RequestApi; auth: DemoAuth; account: AccountInfo; selectedPolicyId?: string;
}) {
  const [settings, setSettings] = useState<SyncSettings | null>(null);
  const [accessMode, setAccessMode] = useState<ReportingAccessMode | null>(null);
  const [job, setJob] = useState<SyncJob | null>(null);
  const [history, setHistory] = useState<SyncJob[]>([]);
  const [cursor, setCursor] = useState<string>();
  const [policyId, setPolicyId] = useState(selectedPolicyId);
  const [concurrency, setConcurrency] = useState(3);
  const [busy, setBusy] = useState(false);
  const [running, setRunning] = useState(false);
  const [phase, setPhase] = useState("");
  const [error, setError] = useState("");
  const [checks, setChecks] = useState<string[]>([]);
  const [notices, setNotices] = useState<string[]>([]);
  const [claims, setClaims] = useState<string>();
  const [apiClaims, setApiClaims] = useState<string>();
  const [links, setLinks] = useState<{ label: string; url: string }[]>([]);
  const controller = useRef<AbortController | null>(null);
  const creation = useRef<{ key: string; id: string } | null>(null);
  useOperationLock(busy);
  useEffect(() => {
    const request = new AbortController();
    void load(request.signal).catch((error: unknown) => { if (!request.signal.aborted) setError(message(error)); });
    return () => { request.abort(); controller.current?.abort(); };
  }, [api]);
  useEffect(() => { setPolicyId(selectedPolicyId); }, [selectedPolicyId]);
  async function load(signal: AbortSignal) {
    const config = await api.syncSettings(signal);
    setSettings(config); setConcurrency(config.defaultConcurrency);
    if (!config.enabled) return;
    const [page, active, reporting] = await Promise.all([api.syncJobs(signal), api.syncCurrent(signal), api.reportingSettings(signal)]);
    setAccessMode(reportingAccessMode(reporting.accessMode));
    setHistory(page.items); setCursor(page.cursor); setJob(active);
  }
  async function act(operation: (signal: AbortSignal) => Promise<void>) {
    const request = new AbortController(); controller.current = request;
    setBusy(true); setError(""); setNotices([]);
    try { await operation(request.signal); }
    catch (error) {
      if (error instanceof GraphError && error.claims) setClaims(error.claims);
      if (error instanceof ApiRequestError && error.claims) setApiClaims(error.claims);
      setError(request.signal.aborted ? "Sync paused. Refresh task status if a lease or reporting update could not be saved." : message(error));
    } finally { setBusy(false); setRunning(false); setPhase(""); }
  }
  const notice = (value: string) => setNotices((values) => [...values, value]);
  async function execute(current: SyncJob, signal: AbortSignal) {
    if (!settings) throw new Error("Sync settings are not loaded.");
    setJob(current); setRunning(true);
    const result = await runSync(api, auth, account, settings, current, signal,
      (value, phase) => { setJob(value); setPhase(phase ?? value.phase); }, notice);
    setJob(result);
    const page = await api.syncJobs(signal); setHistory(page.items); setCursor(page.cursor);
  }
  return <section className="card">
    <h2>Sync data</h2>
    <p>Save spending-policy configuration and user/service balances to SharePoint. Keep this browser open:
      closing it pauses work until an administrator resumes. Existing owner and request workflows are unchanged.</p>
    {!settings && !error && <p role="status">Loading sync configuration...</p>}
    {settings && !settings.enabled && <p className="notice">SharePoint sync is not configured. An operator must run the SharePoint setup script and deploy its destination settings.</p>}
    {settings?.enabled && <>
      <ReportingAccessNotice accessMode={accessMode} />
      <p className="small">Destination site: {settings.siteId}. Job history retention: {settings.retentionDays} days.
        Policy and balance reporting retains only the latest successful snapshot per policy.</p>
      <div className="actions">
        <label>Policy ID (blank for all)<input value={policyId} disabled={busy} onChange={(event) => setPolicyId(event.target.value)} /></label>
        <label>Parallel requests<input type="number" min={1} max={settings.maxConcurrency} value={concurrency} disabled={busy}
          onChange={(event) => setConcurrency(Number(event.target.value))} /></label>
      </div>
      <div className="actions">
        {accessMode === "sharepoint-acl" && <button disabled={busy} onClick={() => void act(async (signal) => {
          const result = await api.reportingReconcile(signal);
          setChecks([`SharePoint row permissions verified: ${result.policies} policies, ${result.rows} rows, ${result.checkedAt}.`]);
        })}>Verify SharePoint permissions</button>}
        <button disabled={busy} onClick={() => void act(async (signal) => {
          const access = await validateSyncAccess(api, auth, account, settings, concurrency, signal, policyId ? guid(policyId, "Policy ID") : undefined);
          setChecks(access.checks); setLinks(access.links); setClaims(undefined);
        })}>Validate access</button>
        <button disabled={busy} onClick={() => void act(async (signal) => {
          await auth.refreshApi(account, apiClaims);
          await auth.refreshSync(account, claims);
          const access = await validateSyncAccess(api, auth, account, settings, concurrency, signal, policyId ? guid(policyId, "Policy ID") : undefined);
          setChecks(access.checks); setLinks(access.links); setClaims(undefined); setApiClaims(undefined);
        })}>Refresh access / sign in</button>
        <button disabled={busy || !!(job && !terminal(job))} onClick={() => void act(async (signal) => {
          const selected = policyId ? guid(policyId, "Policy ID") : undefined;
          const access = await validateSyncAccess(api, auth, account, settings, concurrency, signal, selected);
          setChecks(access.checks); setLinks(access.links);
          const key = JSON.stringify([selected, concurrency]);
          if (creation.current?.key !== key) creation.current = { key, id: crypto.randomUUID() };
          const created = await api.syncCreate({ operationId: creation.current.id, policyId: selected, concurrency }, signal);
          creation.current = null;
          await execute(created, signal);
        })}>Start sync</button>
        <button className="secondary" disabled={busy} onClick={() => void act(load)}>Refresh task status</button>
        {running && <button className="secondary" onClick={() => controller.current?.abort()}>Pause sync</button>}
      </div>
      <p className="small muted">If access is denied, check consent, site permissions and any required active PIM role.
        Activate the appropriate role in <a href="https://entra.microsoft.com/#view/Microsoft_Azure_PIMCommon/ActivationMenuBlade/~/aadRoles"
          target="_blank" rel="noreferrer">Microsoft Entra PIM</a>, then refresh access. A valid token alone does not prove authorization.</p>
      {!!checks.length && <ul aria-label="Access validation results">{checks.map((check) => <li key={check}>{check}</li>)}</ul>}
      {!!links.length && <nav className="actions" aria-label="SharePoint reporting lists">{links.map((link) =>
        <a key={link.url} href={link.url} target="_blank" rel="noreferrer">{link.label}</a>)}</nav>}
      {job && <>
        <SyncJobSummary job={job} />
        <div className="actions">
          {!terminal(job) && <button disabled={busy || !!(job.lease && Date.parse(job.lease.expiresAt) > Date.now())}
            onClick={() => void act((signal) => execute(job, signal))}>Resume / take over</button>}
          {["partially-failed", "failed"].includes(job.state) && <button disabled={busy} onClick={() => void act(async (signal) => {
            const retry = await api.syncUpdate(job.id, { action: "retry" }, signal);
            await execute(retry, signal);
          })}>Retry failed policies</button>}
          {!terminal(job) && <button className="secondary" disabled={busy} onClick={() => void act(async (signal) => {
            await auth.syncToken(account, true);
            const reporting = syncClients(auth, account, settings, 1).reporting;
            const current = await api.syncJob(job.id, signal);
            await reporting.invalidate(current, signal);
            const cancelled = await api.syncUpdate(job.id, { action: "cancel" }, signal); setJob(cancelled);
            setJob(await mirrorSyncJob(api, reporting, cancelled, signal));
          })}>Cancel task</button>}
          <button className="secondary" disabled={busy || !terminal(job) || !!(job.lease && Date.parse(job.lease.expiresAt) > Date.now())}
            onClick={() => void act(async (signal) => {
              await auth.syncToken(account, true);
              const reporting = syncClients(auth, account, settings, 1).reporting;
              const maintenance = await api.syncUpdate(job.id, { action: "claim-maintenance" }, signal);
              const leaseId = maintenance.lease!.id;
              const guard = async () => { await api.syncUpdate(job.id, { action: "heartbeat", leaseId }, signal); };
              try {
                for (const policy of maintenance.policies) {
                  await api.syncUpdate(job.id, { action: "heartbeat", leaseId }, signal);
                  await reporting.cleanupPolicy(policy.id, signal, guard);
                }
                await api.syncUpdate(job.id, { action: "heartbeat", leaseId }, signal);
                await reporting.pruneJobs(signal, job.id, guard);
                await api.syncMaintain(job.id, leaseId, signal);
              } finally {
                const released = await api.syncUpdate(job.id, { action: "release-maintenance", leaseId }, AbortSignal.timeout(15_000));
                setJob(released);
              }
              setJob(await mirrorSyncJob(api, reporting, await api.syncJob(job.id, signal), signal));
            })}>Reconcile reporting / cleanup</button>
        </div>
      </>}
      {!!history.length && <details><summary>Sync job history</summary>
        {history.filter((entry) => entry.id !== job?.id).map((entry) => <div key={entry.id}>
          <SyncJobSummary job={entry} />
          <button className="secondary" disabled={busy} onClick={() => void act(async (signal) => {
            await auth.syncToken(account, true);
            const saved = await mirrorSyncJob(api, syncClients(auth, account, settings, 1).reporting, entry, signal);
            setHistory((values) => values.map((value) => value.id === saved.id ? saved : value));
          })}>Reconcile this job</button>
        </div>)}
        {cursor && <button disabled={busy} onClick={() => void act(async (signal) => {
          const next = await api.syncJobs(signal, cursor);
          setHistory((values) => [...values, ...next.items.filter((item) => !values.some((value) => value.id === item.id))]);
          setCursor(next.cursor);
        })}>Load more jobs</button>}
      </details>}
    </>}
    {busy && <p role="status">{phase || "Checking access and updating sync state..."}</p>}
    {error && <p className="error" role="alert">{error}</p>}
    {notices.map((value, index) => <p className="notice" role="status" key={index}>{value}</p>)}
    <p className="small muted">Preview APIs. Policy totals exclude mismatched/no-balance observations, are grouped by service/unit,
      and are not invoiced currency. Unsupported policy types have inventory only, not zero consumption.
      SharePoint readers must join balances using each policy's ActiveSnapshotId.</p>
  </section>;
}
