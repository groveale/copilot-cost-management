import { RequestApi } from "./api";
import { guid, object, type BudgetConfig, type TierName } from "./config";
import { loadSpendingPolicy, serviceRows } from "./details";
import { GraphClient } from "./graph";
import { verifyAudience } from "./provisioning";

export interface DirectoryUser { id: string; displayName: string; userPrincipalName: string }

export async function directMember(graph: GraphClient, groupId: string, userId: string, signal: AbortSignal): Promise<boolean> {
  const members = await graph.list(`/v1.0/groups/${groupId}/members?$select=id`, signal);
  return members.map((member) => guid(object(member, "Group member").id, "Member ID")).includes(userId);
}

export async function transitiveGroups(graph: GraphClient, userId: string, groupIds: string[], signal: AbortSignal): Promise<string[]> {
  const result = object((await graph.post(`/v1.0/users/${userId}/checkMemberGroups`, { groupIds }, signal)).data, "Membership check");
  if (!Array.isArray(result.value)) throw new Error("Graph returned an invalid membership check.");
  return result.value.map((value) => {
    const id = guid(value, "Membership group ID");
    if (!groupIds.includes(id)) throw new Error("Graph returned an unexpected membership group.");
    return id;
  });
}

export async function verifyMembershipTargets(
  graph: GraphClient, tiers: BudgetConfig["tiers"], serviceId: string, signal: AbortSignal,
): Promise<void> {
  if (tiers.standard.groupId === tiers.enhanced.groupId) throw new Error("Standard and Enhanced must use different groups.");
  for (const tier of [tiers.standard, tiers.enhanced]) {
    const groupId = guid(tier.groupId, "Mapped group ID");
    const policy = await loadSpendingPolicy(graph, tier.policyId, signal);
    if (policy.fields["@odata.type"] !== "#microsoft.graph.selectedGroupsSpendingPolicy" || policy.fields.status !== "enabled") {
      throw new Error("Both mapped policies must be enabled selected-groups policies.");
    }
    await verifyAudience(graph, policy.id, groupId, signal);
    const services = await serviceRows(graph, `/beta/copilot/costManagement/spendingPolicies/${policy.id}/serviceSettings`, signal);
    if (!services.some((row) => row.serviceId === serviceId && row.state === "enabled")) throw new Error("The requested service is not enabled on both policies.");
    const group = object(await graph.get(`/v1.0/groups/${groupId}?$select=id,securityEnabled,mailEnabled,groupTypes,onPremisesSyncEnabled,isAssignableToRole`, signal), "Membership group");
    if (guid(group.id, "Group ID") !== groupId || group.securityEnabled !== true || group.mailEnabled !== false ||
        !Array.isArray(group.groupTypes) || group.groupTypes.length !== 0 ||
        (group.onPremisesSyncEnabled !== false && group.onPremisesSyncEnabled !== null) ||
        (group.isAssignableToRole !== false && group.isAssignableToRole !== null)) {
      throw new Error("Membership changes require cloud-managed, assigned, non-role-assignable security groups.");
    }
  }
}

export async function findDirectoryUser(graph: GraphClient, upn: string, signal: AbortSignal): Promise<DirectoryUser> {
  const username = upn.trim().toLowerCase();
  if (!/^[^@\s/?#]+@[^@\s/?#]+$/.test(username)) throw new Error("Enter the person's full sign-in email (UPN).");
  const user = object(await graph.get(`/v1.0/users/${encodeURIComponent(username)}?$select=id,displayName,userPrincipalName,accountEnabled`, signal), "Directory user");
  const id = guid(user.id, "User ID");
  if (typeof user.userPrincipalName !== "string" || user.userPrincipalName.toLowerCase() !== username) {
    throw new Error("Graph returned a different sign-in identity. No member was selected.");
  }
  if (user.accountEnabled !== true) throw new Error("Only an enabled directory user can be added.");
  return { id, userPrincipalName: user.userPrincipalName,
    displayName: typeof user.displayName === "string" && user.displayName ? user.displayName : user.userPrincipalName };
}

export async function addPolicyMember(
  api: RequestApi, graph: GraphClient, budget: BudgetConfig, tier: TierName, user: DirectoryUser,
  serviceId: string, signal: AbortSignal,
): Promise<void> {
  const userId = guid(user.id, "User ID");
  const target = budget.tiers[tier].groupId;
  const other = budget.tiers[tier === "standard" ? "enhanced" : "standard"].groupId;
  const authorize = () => api.member(budget.id, { action: "check", userId, tiers: budget.tiers }, signal);
  await authorize();
  await verifyMembershipTargets(graph, budget.tiers, serviceId, signal);
  const currentUser = object(await graph.get(`/v1.0/users/${userId}?$select=id,accountEnabled`, signal), "Directory user");
  if (guid(currentUser.id, "User ID") !== userId || currentUser.accountEnabled !== true) throw new Error("The selected user is no longer enabled or could not be verified.");
  const memberships = await transitiveGroups(graph, userId, [target, other], signal);
  if (memberships.includes(other) || await directMember(graph, other, userId, signal)) {
    throw new Error("This person is already in the other tier, directly or through a nested group. No membership was changed.");
  }
  const alreadyDirect = await directMember(graph, target, userId, signal);
  if (memberships.includes(target) && !alreadyDirect) throw new Error("This person inherits the selected tier through another group. Ask an administrator to reconcile it; no membership was changed.");
  await authorize();
  if (!alreadyDirect) {
    await graph.post(`/v1.0/groups/${target}/members/$ref`, { "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${userId}` }, signal);
  }
  if (!await directMember(graph, target, userId, signal) || await directMember(graph, other, userId, signal)) {
    throw new Error("The expected direct membership is not confirmed. Refresh and retry to reconcile it; no request routing was saved.");
  }
  const finalGroups = await transitiveGroups(graph, userId, [target, other], signal);
  if (!finalGroups.includes(target) || finalGroups.includes(other)) {
    throw new Error("Graph has not confirmed the expected tier membership. Allow for propagation, then retry; no request routing was saved.");
  }
  await api.member(budget.id, { action: "complete", userId, tiers: budget.tiers, verified: true }, signal);
}
