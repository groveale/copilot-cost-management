import type { OnboardingRecord, ResourceKey, TierSettings, WorkUpdate } from "../../api/src/onboarding";
import { checkAllocation } from "../../api/src/budget-math";
import { RequestApi } from "./api";
import { GraphClient, GraphError, wait } from "./graph";
import { guid, object, type TierName, type BudgetConfig } from "./config";

const base = "/beta/copilot/costManagement/spendingPolicies";
const tiers: TierName[] = ["standard", "enhanced"];
export function resourceNames(record: OnboardingRecord, tier: TierName) {
  const name = record.name.replace(/[^A-Za-z0-9 ]/g, " ").replace(/\s+/g, " ").trim().slice(0, 36) || "Business Unit";
  return { policy: `${name} ${tier === "standard" ? "Standard" : "Enhanced"} ${record.id.slice(0, 10)}`,
    alias: `bu${record.id.slice(0, 24)}${tier}`, marker: `Copilot BU request ${record.id} ${tier}` };
}
function quantity(value: unknown, whole = true): number | null {
  const data = object(value, "Policy quantity");
  if (data.unit !== "copilotCredit") throw new Error("Policy limits must use Copilot Credits.");
  if (data["@odata.type"] === "#microsoft.graph.unlimitedCostManagementQuantity") return null;
  if (data["@odata.type"] !== "#microsoft.graph.limitedCostManagementQuantity" ||
      typeof data.value !== "number" || !Number.isFinite(data.value) || data.value < 0 || data.value > Number.MAX_SAFE_INTEGER ||
      (whole && !Number.isSafeInteger(data.value))) {
    throw new Error("Expected a non-negative credit limit (whole credits for the policy cap).");
  }
  return data.value;
}
export function templateSettings(value: unknown): TierSettings {
  const data = object(value, "Template policy");
  if (typeof data.billingMethods !== "string") throw new Error("The template has no billing method.");
  const flags = data.billingMethods.split(",").map((flag) => flag.trim());
  const billingMethods = flags.includes("prepaidCredits") && flags.includes("payAsYouGo") && flags.length === 2 ?
    "prepaidCredits,payAsYouGo" : data.billingMethods;
  if (!["notSpecified", "prepaidCredits", "payAsYouGo", "prepaidCredits,payAsYouGo"].includes(billingMethods)) {
    throw new Error("The template has an unsupported billing method. Review the policy before onboarding.");
  }
  if (billingMethods !== "notSpecified" && billingMethods !== "prepaidCredits" &&
      billingMethods !== "payAsYouGo" && billingMethods !== "prepaidCredits,payAsYouGo") throw new Error("Invalid billing method.");
  return { perUserCredits: quantity(data.perUserLimit, false), billingMethods };
}
const limited = (value: number) => ({ "@odata.type": "#microsoft.graph.limitedCostManagementQuantity", value, unit: "copilotCredit" });
function policyBody(record: OnboardingRecord, tier: TierName) {
  return {
    "@odata.type": "#microsoft.graph.selectedGroupsSpendingPolicy", displayName: resourceNames(record, tier).policy,
    periodLimit: limited(record[tier].credits),
    perUserLimit: record[tier].perUserCredits === null ?
      { "@odata.type": "#microsoft.graph.unlimitedCostManagementQuantity", unit: "copilotCredit" } : limited(record[tier].perUserCredits),
    billingMethods: record[tier].billingMethods, serviceSettings: [{ serviceId: "cowork", state: "enabled" }],
  };
}
async function policy(graph: GraphClient, id: string, signal: AbortSignal) {
  const result = await graph.getWithEtag(`${base}/${id}`, signal);
  const data = object(result.data, "Spending policy");
  if (guid(data.id, "Policy ID") !== id || data["@odata.type"] !== "#microsoft.graph.selectedGroupsSpendingPolicy") {
    throw new Error("Unexpected policy identity or audience type.");
  }
  if (!result.etag) throw new Error("Graph returned no policy ETag.");
  return { data, etag: result.etag };
}
export async function readAllocation(graph: GraphClient, budget: BudgetConfig, signal: AbortSignal) {
  const standardPolicy = (await policy(graph, budget.tiers.standard.policyId, signal)).data;
  const enhancedPolicy = (await policy(graph, budget.tiers.enhanced.policyId, signal)).data;
  const standard = quantity(standardPolicy.periodLimit);
  const enhanced = quantity(enhancedPolicy.periodLimit);
  if (standard === null || enhanced === null) throw new Error("An actual policy is unlimited; administrator reconciliation is required.");
  return { standard, enhanced, perUser: {
    standard: quantity(standardPolicy.perUserLimit, false), enhanced: quantity(enhancedPolicy.perUserLimit, false),
  } };
}
async function ready(graph: GraphClient, id: string, signal: AbortSignal) {
  for (let attempt = 0; attempt < 9; attempt++) {
    const current = await policy(graph, id, signal);
    if (current.data.status === "enabled") return current;
    if (current.data.status !== "provisioning" && current.data.status !== "disabled") throw new Error(`Policy ${id} is ${String(current.data.status)}; administrator action is required.`);
    if (attempt < 8) await wait(5000, signal);
  }
  throw new Error("Policy activation is still pending. Refresh and resume; recorded resource IDs are retained.");
}
async function verifyGroup(graph: GraphClient, id: string, record: OnboardingRecord, tier: TierName, signal: AbortSignal) {
  const group = object(await graph.get(`/v1.0/groups/${id}`, signal), "Security group");
  if (guid(group.id, "Group ID") !== id || group.securityEnabled !== true || group.mailEnabled !== false ||
      group.description !== resourceNames(record, tier).marker || group.isAssignableToRole === true ||
      !Array.isArray(group.groupTypes) || group.groupTypes.length !== 0) throw new Error("Group does not match this onboarding operation.");
  const owners = await graph.list(`/v1.0/groups/${id}/owners`, signal);
  if (!owners.some((owner) => object(owner, "Group owner").id === record.requesterId)) throw new Error("The proposed BU owner is not an owner of the group.");
}
export async function verifyAudience(graph: GraphClient, id: string, groupId: string, signal: AbortSignal) {
  const groups = await graph.list(`${base}/${id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups`, signal);
  if (groups.length !== 1 || object(groups[0], "Assigned group").id !== groupId) throw new Error("Policy group assignments differ from this BU.");
}
async function verifyPolicy(graph: GraphClient, record: OnboardingRecord, tier: TierName, signal: AbortSignal) {
  const id = record.resources[tier === "standard" ? "standardPolicy" : "enhancedPolicy"];
  const groupId = record.resources[tier === "standard" ? "standardGroup" : "enhancedGroup"];
  if (!id || !groupId) throw new Error("Missing recorded group or policy ID.");
  const current = await ready(graph, id, signal);
  if (quantity(current.data.periodLimit) !== record[tier].credits) throw new Error("Actual policy cap differs from the approved allocation. Reconcile before completing.");
  if (record.kind === "allocation" && record.perUserEdit?.tier === tier &&
      quantity(current.data.perUserLimit, false) !== record.perUserEdit.credits) {
    throw new Error("Actual per-user limit differs from the saved policy edit. Refresh and resume.");
  }
  if (record.kind === "onboarding") {
    if (current.data.displayName !== resourceNames(record, tier).policy ||
        JSON.stringify(templateSettings(current.data)) !== JSON.stringify({
          perUserCredits: record[tier].perUserCredits, billingMethods: record[tier].billingMethods,
        })) throw new Error("Policy settings differ from the approved configuration.");
    await verifyGroup(graph, groupId, record, tier, signal);
  }
  await verifyAudience(graph, id, groupId, signal);
  const services = object(await graph.get(`${base}/${id}/serviceSettings`, signal), "Policy services");
  if (!Array.isArray(services.value) || !services.value.some((value) => {
    const service = object(value, "Policy service"); return service.serviceId === "cowork" && service.state === "enabled";
  })) throw new Error("Cowork is not enabled for this policy.");
}
export async function runProvisioning(api: RequestApi, graph: GraphClient, input: OnboardingRecord,
  signal: AbortSignal, progress: (record: OnboardingRecord) => void): Promise<OnboardingRecord> {
  let record = await api.onboardingGet(input.id, signal);
  if (record.status === "completed") { progress(record); return record; }
  record = await api.onboardingWork(record.id, { action: "claim", etag: record.etag }, signal);
  const leaseId = record.lease?.id;
  if (!leaseId) throw new Error("No provisioning lease was returned.");
  progress(record);
  async function save(update: Omit<WorkUpdate, "etag" | "leaseId">) {
    record = await api.onboardingWork(record.id, { ...update, etag: record.etag, leaseId }, signal);
    progress(record);
  }
  async function resource(key: ResourceKey, find: () => Promise<unknown[]>, create: () => Promise<{ data: unknown }>) {
    if (record.resources[key]) return;
    const previousIntent = record.intents.includes(key);
    const found = await find();
    if (found.length > 1) throw new Error(`Multiple resources match ${key}. Administrator reconciliation is required.`);
    if (!found.length && previousIntent) throw new Error(`${key} creation has an uncertain outcome. Wait and resume to search again; no duplicate will be created automatically.`);
    if (!previousIntent) await save({ action: "intent", resource: key });
    let data: unknown;
    try { data = found.length ? found[0] : (await create()).data; }
    catch (error) {
      if (error instanceof GraphError && [400, 401, 403, 429].includes(error.status)) await save({ action: "clear-intent", resource: key });
      throw error;
    }
    await save({ action: "resource", resource: key, resourceId: guid(object(data, key).id, `${key} ID`) });
  }
  try {
    checkAllocation(record.standard.credits, record.enhanced.credits, record.approvedCredits);
    if (record.kind === "onboarding") {
      for (const tier of tiers) {
        const names = resourceNames(record, tier);
        const key = tier === "standard" ? "standardGroup" : "enhancedGroup";
        await resource(key,
          () => graph.list(`/v1.0/groups?$filter=${encodeURIComponent(`mailNickname eq '${names.alias}'`)}`, signal),
          () => graph.post("/v1.0/groups", { displayName: names.policy, mailNickname: names.alias,
            description: names.marker, mailEnabled: false, securityEnabled: true, groupTypes: [],
            "owners@odata.bind": [`https://graph.microsoft.com/v1.0/users/${record.requesterId}`] }, signal));
        await verifyGroup(graph, record.resources[key]!, record, tier, signal);
      }
      for (const tier of tiers) {
        const names = resourceNames(record, tier);
        const key = tier === "standard" ? "standardPolicy" : "enhancedPolicy";
        await resource(key,
          () => graph.list(`${base}?$filter=${encodeURIComponent(`displayName eq '${names.policy}'`)}`, signal),
          () => graph.post(base, policyBody(record, tier), signal));
        const id = record.resources[key]!;
        const current = await policy(graph, id, signal);
        if (current.data.displayName !== names.policy || quantity(current.data.periodLimit) !== record[tier].credits ||
            JSON.stringify(templateSettings(current.data)) !== JSON.stringify({
              perUserCredits: record[tier].perUserCredits, billingMethods: record[tier].billingMethods,
            })) throw new Error("Recovered policy does not match the approved resource.");
        const services = object(await graph.get(`${base}/${id}/serviceSettings`, signal), "Policy services");
        if (!Array.isArray(services.value) || !services.value.some((value) => {
          const service = object(value, "Service setting"); return service.serviceId === "cowork" && service.state === "enabled";
        })) throw new Error("Recovered policy does not have Cowork enabled.");
        const assignment = `${base}/${id}/microsoft.graph.selectedGroupsSpendingPolicy/assignedGroups`;
        const groups = await graph.list(assignment, signal);
        const groupId = record.resources[tier === "standard" ? "standardGroup" : "enhancedGroup"];
        if (!groups.length) {
          if (current.data.status === "provisioning") throw new Error("Policy is provisioning. Resume once it is editable.");
          await graph.post(`${assignment}/$ref`, { "@odata.id": `https://graph.microsoft.com/beta/groups/${groupId}` }, signal, current.etag);
        } else if (groups.length !== 1 || object(groups[0], "Assigned group").id !== groupId) throw new Error("Recovered policy has a different audience; no assignments were changed.");
        await ready(graph, id, signal);
        await save({ action: "heartbeat" });
      }
    } else {
      const actual = await Promise.all(tiers.map(async (tier) => {
        const id = record.resources[tier === "standard" ? "standardPolicy" : "enhancedPolicy"];
        if (!id) throw new Error("Missing allocation policy ID.");
        const current = await ready(graph, id, signal);
        const groupId = record.resources[tier === "standard" ? "standardGroup" : "enhancedGroup"];
        if (!groupId) throw new Error("Missing allocation group ID.");
        await verifyAudience(graph, id, groupId, signal);
        const credits = quantity(current.data.periodLimit);
        if (credits === null) throw new Error("An unlimited policy cannot be reallocated under a finite approved ceiling.");
        return { tier, id, credits };
      }));
      actual.sort((a, b) => (record[a.tier].credits - a.credits) - (record[b.tier].credits - b.credits));
      const editedTier = record.perUserEdit?.tier;
      if (editedTier && actual.some((item) => item.tier !== editedTier && item.credits !== record[item.tier].credits)) {
        throw new Error("The other policy allocation changed. No external changes were overwritten; reconcile the saved edit before resuming.");
      }
      for (const item of actual) {
        const current = await policy(graph, item.id, signal);
        const other = actual.find((value) => value.tier !== item.tier)!;
        const otherCurrent = await policy(graph, other.id, signal);
        const otherCredits = quantity(otherCurrent.data.periodLimit);
        if (otherCredits === null) throw new Error("The other policy is unlimited; reconcile external changes first.");
        if (record.perUserEdit && other.tier !== record.perUserEdit.tier && otherCredits !== record[other.tier].credits) {
          throw new Error("The other policy allocation changed. Reconcile the saved edit; it was not overwritten.");
        }
        const currentCredits = quantity(current.data.periodLimit);
        if (currentCredits === null) throw new Error("An external change made this policy unlimited.");
        if (record[item.tier].credits >= currentCredits) checkAllocation(record[item.tier].credits, otherCredits, record.approvedCredits);
        const edit = record.perUserEdit?.tier === item.tier ? record.perUserEdit : undefined;
        const changePerUser = !!edit && quantity(current.data.perUserLimit, false) !== edit.credits;
        if (currentCredits !== record[item.tier].credits || changePerUser) {
          if (record.perUserEdit && record.perUserEdit.tier !== item.tier) {
            throw new Error("The other policy allocation changed. Reconcile the saved edit; it was not overwritten.");
          }
          await save({ action: "heartbeat" });
          await graph.patch(`${base}/${item.id}`, {
            ...(currentCredits !== record[item.tier].credits ? { periodLimit: limited(record[item.tier].credits) } : {}),
            ...(changePerUser && edit ? { perUserLimit: edit.credits === null ?
              { "@odata.type": "#microsoft.graph.unlimitedCostManagementQuantity", unit: "copilotCredit" } : limited(edit.credits) } : {}),
          }, signal, current.etag);
          await ready(graph, item.id, signal);
        }
      }
    }
    for (const tier of tiers) await verifyPolicy(graph, record, tier, signal);
    await save({ action: "complete", verified: true });
    return record;
  } catch (error) {
    const explanation = error instanceof Error ? error.message : String(error);
    try {
      const latest = await api.onboardingGet(record.id, AbortSignal.timeout(15000));
      if (latest.lease?.id === leaseId) await api.onboardingWork(latest.id, { action: "release", etag: latest.etag, leaseId,
        failure: "Provisioning was interrupted. An administrator or owner must refresh and resume or reconcile the recorded resources." }, AbortSignal.timeout(15000));
    } catch (releaseError) {
      throw new Error(`${explanation} The operation could not be released: ${releaseError instanceof Error ? releaseError.message : String(releaseError)}. Wait five minutes before resuming.`);
    }
    throw new Error(`${explanation} Progress was retained; refresh and resume rather than submitting a duplicate request.`);
  }
}
