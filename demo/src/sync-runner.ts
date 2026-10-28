import type { AccountInfo } from "@azure/msal-browser";
import type { SyncJob, SyncSettings, SyncUpdate } from "../../api/src/sync-contracts";
import { ApiRequestError, RequestApi } from "./api";
import { DemoAuth, isGraphInteractionRequired } from "./auth";
import { GraphClient, GraphError, RequestScheduler } from "./graph";
import { collectPolicy, policyInventory, probeCollectionAccess } from "./sync-collector";
import { SharePointSync } from "./sync-sharepoint";

export const syncAccessFailure = (error: unknown) => isGraphInteractionRequired(error) ||
  (error instanceof GraphError && [401, 403].includes(error.status)) ||
  (error instanceof ApiRequestError && [401, 403].includes(error.status));
export function syncClients(auth: DemoAuth, account: AccountInfo, settings: SyncSettings, concurrency: number) {
  const graph = new GraphClient(() => auth.syncToken(account), fetch.bind(globalThis), new RequestScheduler(concurrency));
  return { graph, reporting: new SharePointSync(graph, settings, account.tenantId) };
}
export async function validateSyncAccess(api: RequestApi, auth: DemoAuth, account: AccountInfo,
  settings: SyncSettings, concurrency: number, signal: AbortSignal, policyId?: string) {
  if (!settings.enabled) throw new Error("Provision and configure the SharePoint sync destination first.");
  if (!Number.isInteger(concurrency) || concurrency < 1 || concurrency > settings.maxConcurrency) throw new Error("Invalid sync concurrency.");
  const capabilities = await api.onboardingCapabilities(signal);
  if (!capabilities.isAdmin) throw new Error("Administrator access is required.");
  await auth.syncToken(account, true);
  const { graph, reporting } = syncClients(auth, account, settings, concurrency);
  const policies = await policyInventory(graph, signal, policyId);
  const checks = [...await reporting.validate(signal), ...await probeCollectionAccess(graph, policies, signal)];
  const links = await reporting.links(signal);
  return { graph, reporting, policies, checks, links };
}
export async function mirrorSyncJob(api: RequestApi, reporting: SharePointSync, job: SyncJob, signal: AbortSignal) {
  await reporting.mirror(job, signal);
  return api.syncUpdate(job.id, { action: "mirrored", revision: job.revision }, signal);
}
export async function runSync(api: RequestApi, auth: DemoAuth, account: AccountInfo, settings: SyncSettings,
  original: SyncJob, externalSignal: AbortSignal, onProgress: (job: SyncJob, phase?: string) => void,
  onNotice: (message: string) => void): Promise<SyncJob> {
  const access = await validateSyncAccess(api, auth, account, settings, original.concurrency, externalSignal, original.policyId);
  let job = await api.syncUpdate(original.id, { action: "claim" }, externalSignal);
  const leaseId = job.lease!.id;
  const controller = new AbortController();
  const signal = AbortSignal.any([externalSignal, controller.signal]);
  let tail = Promise.resolve();
  const update = (input: SyncUpdate, requestSignal = signal) => {
    const operation = tail.then(async () => {
      job = await api.syncUpdate(job.id, input, requestSignal);
      onProgress(job);
      return job;
    });
    tail = operation.then(() => undefined, () => undefined);
    return operation;
  };
  let heartbeatBusy = false;
  const timer = setInterval(() => {
    if (heartbeatBusy) return;
    heartbeatBusy = true;
    void update({ action: "heartbeat", leaseId }).catch((error: unknown) => controller.abort(error))
      .finally(() => { heartbeatBusy = false; });
  }, 20_000);
  const mirror = async (required = false) => {
    try { job = await mirrorSyncJob(api, access.reporting, job, signal); onProgress(job); }
    catch (error) {
      signal.throwIfAborted();
      if (required || syncAccessFailure(error)) throw error;
      onNotice(`SharePoint job mirror is out of date: ${error instanceof Error ? error.message : String(error)}. Use Reconcile reporting.`);
    }
  };
  try {
    onProgress(job);
    await mirror(true);
    if (!job.inventoryComplete) {
      await update({ action: "inventory", leaseId, policies: access.policies.map(({ id, displayName, type }) => ({ id, displayName, type })) });
      await mirror();
    }
    const metadata = await api.onboardingSetup(signal);
    for (const descriptor of job.policies) {
      if (["completed", "excluded"].includes(descriptor.state)) continue;
      signal.throwIfAborted();
      await update({ action: "heartbeat", leaseId });
      const recovered = await access.reporting.recover(descriptor, job, signal);
      if (recovered) {
        await update({ action: "policy", leaseId, policy: recovered });
        await mirror();
        continue;
      }
      await update({ action: "policy", leaseId, policy: { ...descriptor, state: "running" } });
      const fence = await access.reporting.fence(descriptor, job, signal);
      let snapshot;
      try {
        snapshot = await collectPolicy(access.graph, descriptor, job.concurrency, signal,
          (phase) => onProgress(job, `${descriptor.displayName}: ${phase}`));
      } catch (error) {
        signal.throwIfAborted();
        if (syncAccessFailure(error) || error instanceof TypeError ||
            (error instanceof GraphError && [429, 500, 502, 503, 504].includes(error.status))) throw error;
        const code = error instanceof GraphError ? `Graph HTTP ${error.status}` : "Invalid or incomplete source data";
        onNotice(`${descriptor.displayName}: ${error instanceof Error ? error.message : String(error)}`);
        await update({ action: "policy", leaseId, policy: { ...descriptor, state: "failed", error: code } });
        await mirror();
        continue;
      }
      onProgress(job, `${descriptor.displayName}: Publishing SharePoint snapshot`);
      await access.reporting.publish(fence, snapshot, descriptor, job, metadata, signal,
        async () => { await update({ action: "heartbeat", leaseId }); });
      let cleanupRequired = false;
      try { await access.reporting.cleanupPolicy(descriptor.id, signal, async () => { await update({ action: "heartbeat", leaseId }); }); }
      catch (error) {
        signal.throwIfAborted();
        cleanupRequired = true;
        onNotice(`Snapshot published; superseded-row cleanup needs retry: ${error instanceof Error ? error.message : String(error)}`);
      }
      await update({ action: "policy", leaseId, policy: {
        id: descriptor.id, displayName: descriptor.displayName, type: descriptor.type,
        state: snapshot.unsupported ? "excluded" : "completed", users: snapshot.users, groups: snapshot.groups.length,
        exclusions: snapshot.exclusions, ...(cleanupRequired ? { error: "CleanupRequired" } : {}),
      } });
      await mirror();
    }
    if (!job.policyId && job.inventoryComplete) {
      await update({ action: "heartbeat", leaseId });
      await access.reporting.markAbsent(job.policies, signal);
    }
    clearInterval(timer);
    await tail;
    await update({ action: "finish", leaseId });
    await mirror();
    return job;
  } catch (error) {
    clearInterval(timer);
    await tail;
    const failure = controller.signal.aborted ? controller.signal.reason : error;
    const recoverySignal = AbortSignal.timeout(15_000);
    try {
      await update({ action: "pause", leaseId, accessRequired: syncAccessFailure(failure) }, recoverySignal);
      await mirrorSyncJob(api, access.reporting, job, recoverySignal);
    } catch (pauseError) {
      onNotice(`Could not record/reconcile pause: ${pauseError instanceof Error ? pauseError.message : String(pauseError)}. Refresh task status; the lease will expire if it is no longer renewed.`);
    }
    throw failure;
  } finally { clearInterval(timer); }
}
