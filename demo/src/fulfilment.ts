import type { FulfilmentUpdate, RequestRecord } from "../../api/src/contracts";
import { RequestApi } from "./api";
import { guid, object } from "./config";
import { GraphClient } from "./graph";
import { directMember, transitiveGroups, verifyMembershipTargets } from "./membership";

const message = (error: unknown) => error instanceof Error ? error.message : String(error);
export function requestStatusLabel(record: RequestRecord, now = Date.now()): string {
  const expired = !!record.expiresAt && Date.parse(record.expiresAt) <= now;
  if (record.status === "rejected") return "Rejected";
  if (record.status !== "approved") return expired ? "Expired / not approved" : "Pending owner decision";
  if (record.fulfilment?.state === "applied") return "Approved - Enhanced move applied";
  if (expired) return "Approved - request period expired; review any partial upgrade manually";
  if (record.fulfilment?.state === "applying") return "Approved - Enhanced upgrade attempt in progress";
  return record.fulfilment?.state === "needs-attention" ?
    "Approved - Enhanced upgrade needs attention" : "Approved - Enhanced upgrade not yet applied";
}

export async function applyMembershipMove(
  graph: GraphClient, record: RequestRecord, serviceId: string, signal: AbortSignal, beforeWrite: () => Promise<void>,
): Promise<void> {
  const plan = record.fulfilment;
  if (!plan || record.status !== "approved") throw new Error("An approved, server-mapped upgrade plan is required.");
  const userId = guid(record.requesterId, "Requester ID");
  const standard = guid(plan.standard.groupId, "Standard group ID");
  const enhanced = guid(plan.enhanced.groupId, "Enhanced group ID");
  if (standard === enhanced) throw new Error("Standard and Enhanced must use different groups.");
  const checkPeriod = () => {
    signal.throwIfAborted();
    if (record.expiresAt && Date.parse(record.expiresAt) <= Date.now()) {
      throw new Error("The request period expired. Review any partial upgrade manually; no further membership changes will be made.");
    }
  };
  checkPeriod();
  await verifyMembershipTargets(graph, plan, serviceId, signal);
  const user = object(await graph.get(`/v1.0/users/${userId}?$select=id,accountEnabled`, signal), "Requester");
  if (guid(user.id, "Returned requester ID") !== userId || user.accountEnabled !== true) throw new Error("The requester must be the expected enabled directory user.");
  const inStandard = await directMember(graph, standard, userId, signal);
  const inEnhanced = await directMember(graph, enhanced, userId, signal);
  const inherited = await transitiveGroups(graph, userId, [standard, enhanced], signal);
  if ((!inStandard && inherited.includes(standard)) || (!inEnhanced && inherited.includes(enhanced))) {
    throw new Error("Tier membership is inherited through another group. An administrator must reconcile it; no membership changes were made.");
  }
  if (!inEnhanced) {
    await beforeWrite(); checkPeriod();
    await graph.post(`/v1.0/groups/${enhanced}/members/$ref`, { "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}` }, signal);
  }
  if (!await directMember(graph, enhanced, userId, signal)) throw new Error("Enhanced membership is not yet confirmed. Standard was not removed; refresh and retry.");
  if (inStandard) {
    await beforeWrite(); checkPeriod();
    await graph.deleteReference(`/v1.0/groups/${standard}/members/${userId}/$ref`, signal);
  }
  if (await directMember(graph, standard, userId, signal) || !await directMember(graph, enhanced, userId, signal)) {
    throw new Error("The final direct memberships are not confirmed. Refresh and retry; no completed upgrade was recorded.");
  }
  const finalGroups = await transitiveGroups(graph, userId, [standard, enhanced], signal);
  if (finalGroups.includes(standard) || !finalGroups.includes(enhanced)) {
    throw new Error("Graph still reports an unexpected or inherited tier membership. Allow for propagation, then refresh and retry; administrator reconciliation may be needed.");
  }
}

export async function runFulfilment(
  api: RequestApi, graph: GraphClient, input: RequestRecord, serviceId: string,
  signal: AbortSignal, progress: (record: RequestRecord) => void,
): Promise<RequestRecord> {
  let record = await api.fulfilment(input.id, { action: "claim", etag: input.etag }, signal);
  progress(record);
  if (record.fulfilment?.state === "applied") return record;
  const leaseId = record.fulfilment?.lease?.id;
  if (!leaseId) throw new Error("No upgrade lease was returned. Refresh before retrying.");
  async function save(action: FulfilmentUpdate["action"]) {
    record = await api.fulfilment(record.id, { action, etag: record.etag, leaseId, ...(action === "complete" ? { verified: true } : {}) }, signal);
    progress(record);
  }
  try {
    await applyMembershipMove(graph, record, serviceId, signal, () => save("heartbeat"));
    await save("complete");
    return record;
  } catch (error) {
    if (!signal.aborted) {
      try { await save("release"); }
      catch (saveError) {
        throw new Error(`${message(error)} Upgrade progress could not be confirmed: ${message(saveError)} Refresh before retrying.`);
      }
    }
    throw error;
  }
}
