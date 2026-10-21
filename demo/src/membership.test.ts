import { describe, expect, it, vi } from "vitest";
import { RequestApi } from "./api";
import type { BudgetConfig } from "./config";
import { GraphClient } from "./graph";
import { addPolicyMember, findDirectoryUser, type DirectoryUser } from "./membership";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const budget: BudgetConfig = {
  id: "legal", displayName: "Legal", businessArea: "Corporate", costCenter: "CC100", owners: [id(1)],
  tiers: { standard: { policyId: id(3), groupId: id(4) }, enhanced: { policyId: id(5), groupId: id(6) } },
};
const user: DirectoryUser = { id: id(2), displayName: "Example person", userPrincipalName: "person@example.test" };
const response = (value: unknown, status = 200) => new Response(status === 204 ? null : JSON.stringify(value), { status });
interface Options {
  inStandard?: boolean; inEnhanced?: boolean; inheritedStandard?: boolean; inheritedEnhanced?: boolean;
  unsafe?: "dynamic" | "synced" | "role" | "mail"; disabledPolicy?: boolean; wrongAudience?: boolean;
  disabledUser?: boolean; differentUser?: boolean; unknownGroup?: boolean; denyMembers?: boolean;
  failPreflight?: boolean; failSecondCheck?: boolean; failCompletion?: boolean;
  lostResponse?: boolean; addDenied?: boolean; noFinalTransitive?: boolean;
}
function fixture(options: Options = {}) {
  const members = { standard: options.inStandard ?? false, enhanced: options.inEnhanced ?? false };
  const writes: string[] = [];
  const transport = vi.fn<typeof fetch>(async (input, init) => {
    const url = new URL(String(input));
    const path = decodeURIComponent(url.pathname);
    if (path.endsWith("/members/$ref")) {
      writes.push(path);
      expect(init?.method).toBe("POST");
      expect(JSON.parse(String(init?.body))).toEqual({ "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${user.id}` });
      if (options.addDenied) return response({ error: { code: "Denied", message: "Add denied" } }, 403);
      if (path.includes(id(4))) members.standard = true; else members.enhanced = true;
      if (options.lostResponse) { options.lostResponse = false; throw new Error("Add response lost"); }
      return response(null, 204);
    }
    if (path.endsWith("/checkMemberGroups")) return response({ value: options.unknownGroup ? [id(99)] :
      options.noFinalTransitive && writes.length ? [] : [
        ...(members.standard || options.inheritedStandard ? [id(4)] : []),
        ...(members.enhanced || options.inheritedEnhanced ? [id(6)] : []),
      ] });
    if (path.endsWith("/members")) {
      if (options.denyMembers) return response({ error: { code: "Denied", message: "Members denied" } }, 403);
      return response({ value: (path.includes(id(4)) ? members.standard : members.enhanced) ? [{ id: user.id }] : [] });
    }
    if (path.startsWith("/v1.0/users/")) return response({
      ...user, userPrincipalName: options.differentUser ? "other@example.test" : user.userPrincipalName,
      accountEnabled: !options.disabledUser,
    });
    if (path.endsWith("/assignedGroups")) return response({ value: [{
      id: options.wrongAudience ? id(99) : path.includes(id(3)) ? id(4) : id(6),
    }] });
    if (path.endsWith("/serviceSettings")) return response({ serviceId: "cowork", state: "enabled" });
    if (path.startsWith("/beta/copilot/costManagement/spendingPolicies/")) return response({
      id: path.split("/").at(-1), "@odata.type": "#microsoft.graph.selectedGroupsSpendingPolicy",
      status: options.disabledPolicy ? "disabled" : "enabled",
    });
    if (path.startsWith("/v1.0/groups/")) return response({
      id: path.split("/").at(-1), securityEnabled: true, mailEnabled: options.unsafe === "mail",
      groupTypes: options.unsafe === "dynamic" ? ["DynamicMembership"] : [],
      onPremisesSyncEnabled: options.unsafe === "synced" ? true : null,
      isAssignableToRole: options.unsafe === "role",
    });
    throw new Error(`Unexpected Graph path ${path}`);
  });
  const api = new RequestApi("https://example.test", async () => "synthetic-api-token");
  let checks = 0;
  const routing = vi.spyOn(api, "member").mockImplementation(async (_budgetId, input) => {
    if (input.action === "check") {
      checks++;
      if (options.failPreflight || (options.failSecondCheck && checks === 2)) throw new Error("Ownership or routing changed");
    } else if (options.failCompletion) throw new Error("Routing could not be saved");
    return { budget };
  });
  const graph = new GraphClient(async () => "synthetic-graph-token", transport);
  const add = (tier: "standard" | "enhanced" = "standard") =>
    addPolicyMember(api, graph, budget, tier, user, "cowork", new AbortController().signal);
  return { graph, transport, writes, members, routing, add };
}

describe("explicit policy member addition", () => {
  it.each(["standard", "enhanced"] as const)("adds a new %s member, verifies membership, then saves private routing", async (tier) => {
    const { add, writes, members, routing } = fixture();
    await add(tier);
    expect(writes).toEqual([`/v1.0/groups/${budget.tiers[tier].groupId}/members/$ref`]);
    expect(members[tier]).toBe(true);
    expect(routing.mock.calls.map(([, input]) => input.action)).toEqual(["check", "check", "complete"]);
    expect(routing.mock.calls.at(-1)?.[1]).toEqual({ action: "complete", userId: user.id, tiers: budget.tiers, verified: true });
  });
  it.each([{ inEnhanced: true }, { inheritedEnhanced: true }])("blocks direct or inherited other-tier membership: %j", async (options) => {
    const { add, writes, routing } = fixture(options);
    await expect(add()).rejects.toThrow("already in the other tier");
    expect(writes).toEqual([]);
    expect(routing.mock.calls.some(([, input]) => input.action === "complete")).toBe(false);
  });
  it("blocks inherited target membership rather than changing directory topology", async () => {
    const { add, writes } = fixture({ inheritedStandard: true });
    await expect(add()).rejects.toThrow("inherits the selected tier");
    expect(writes).toEqual([]);
  });
  it("reconciles an existing direct member without issuing another membership write", async () => {
    const { add, writes, routing } = fixture({ inStandard: true });
    await add();
    expect(writes).toEqual([]);
    expect(routing.mock.calls.at(-1)?.[1].action).toBe("complete");
  });
  it.each(["dynamic", "synced", "role", "mail"] as const)("refuses an unsafe %s group before any member write", async (unsafe) => {
    const { add, writes } = fixture({ unsafe });
    await expect(add()).rejects.toThrow("Membership changes require");
    expect(writes).toEqual([]);
  });
  it.each([
    [{ disabledPolicy: true }, "enabled selected-groups"],
    [{ wrongAudience: true }, "assignments differ"],
    [{ disabledUser: true }, "no longer enabled"],
    [{ unknownGroup: true }, "unexpected membership group"],
    [{ denyMembers: true }, "Members denied"],
    [{ failPreflight: true }, "Ownership or routing changed"],
    [{ failSecondCheck: true }, "Ownership or routing changed"],
  ] as const)("fails closed on unverifiable targets or authorization: %j", async (options, error) => {
    const { add, writes, routing } = fixture(options);
    await expect(add()).rejects.toThrow(error);
    expect(writes).toEqual([]);
    expect(routing.mock.calls.some(([, input]) => input.action === "complete")).toBe(false);
  });
  it("does not report routing completion when Graph rejects the add", async () => {
    const { add, routing, members } = fixture({ addDenied: true });
    await expect(add()).rejects.toThrow("Add denied");
    expect(members.standard).toBe(false);
    expect(routing.mock.calls.some(([, input]) => input.action === "complete")).toBe(false);
  });
  it("reconciles a lost add response on explicit retry instead of adding the same member twice", async () => {
    const { add, writes, routing } = fixture({ lostResponse: true });
    await expect(add()).rejects.toThrow("Add response lost");
    expect(routing.mock.calls.some(([, input]) => input.action === "complete")).toBe(false);
    await add();
    expect(writes).toHaveLength(1);
    expect(routing.mock.calls.at(-1)?.[1].action).toBe("complete");
  });
  it("retains an added membership but explicitly fails when routing cannot be saved, then repairs it on retry", async () => {
    const options = { failCompletion: true };
    const { add, writes, members } = fixture(options);
    await expect(add()).rejects.toThrow("Routing could not be saved");
    expect(members.standard).toBe(true);
    options.failCompletion = false;
    await add();
    expect(writes).toHaveLength(1);
  });
  it("does not save routing before final Graph membership confirmation", async () => {
    const { add, writes, routing } = fixture({ noFinalTransitive: true });
    await expect(add()).rejects.toThrow("has not confirmed");
    expect(writes).toHaveLength(1);
    expect(routing.mock.calls.some(([, input]) => input.action === "complete")).toBe(false);
  });
});

describe("exact directory-user lookup", () => {
  it("looks up only the submitted UPN and returns the identity for explicit confirmation", async () => {
    const { graph, transport, writes, routing } = fixture();
    expect(await findDirectoryUser(graph, " PERSON@example.test ", new AbortController().signal)).toEqual(user);
    const url = new URL(String(transport.mock.calls[0][0]));
    expect(decodeURIComponent(url.pathname)).toBe("/v1.0/users/person@example.test");
    expect(url.searchParams.get("$select")).toContain("accountEnabled");
    expect(writes).toEqual([]); expect(routing).not.toHaveBeenCalled();
  });
  it.each(["not-an-email", "person@example.test/members", "person@example.test?x=1"])("rejects invalid lookup input %s without calling Graph", async (upn) => {
    const { graph, transport } = fixture();
    await expect(findDirectoryUser(graph, upn, new AbortController().signal)).rejects.toThrow("full sign-in email");
    expect(transport).not.toHaveBeenCalled();
  });
  it.each([{ disabledUser: true }, { differentUser: true }])("does not select a disabled or mismatched identity: %j", async (options) => {
    const { graph } = fixture(options);
    await expect(findDirectoryUser(graph, user.userPrincipalName, new AbortController().signal)).rejects.toThrow();
  });
});
