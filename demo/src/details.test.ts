import { createElement } from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import { parseConfig } from "./config";
import { GraphClient } from "./graph";
import { BudgetDetails } from "./BudgetDetails";
import {
  formatConsumed, formatQuantity, initialSnapshot, loadBudgetDetails,
  memberPolicyNote, stopLoading, tierGroupIds, type BudgetSnapshot,
} from "./details";

const id = (n: number) => `10000000-0000-0000-0000-${String(n).padStart(12, "0")}`;
const config = parseConfig({
  version: 1, tenantId: id(1), clientId: id(2), serviceId: "cowork",
  budgets: [{
    id: "legal", displayName: "Legal", businessArea: "Legal", costCenter: "Not specified", owners: ["operator@example.test"],
    tiers: {
      standard: { policyId: id(3), groupId: id(4) },
      enhanced: { policyId: id(5), groupId: id(6) },
    },
  }],
});
const budget = config.budgets[0];
const limited = (value: number) => ({ "@odata.type": "#microsoft.graph.limitedCostManagementQuantity", value, unit: "copilotCredit" });
const unlimited = { "@odata.type": "#microsoft.graph.unlimitedCostManagementQuantity", unit: "copilotCredit" };
const user = (n: number) => ({ id: id(n), "@odata.type": "#microsoft.graph.user", displayName: `Test person ${n}`, userPrincipalName: `person${n}@example.test` });
const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { "request-id": "test-request" } });

interface Options {
  singleton?: boolean;
  overlap?: boolean;
  extraGroup?: boolean;
  failedPolicy?: boolean;
  failedServices?: boolean;
  failedAssignment?: boolean;
  noAssignment?: boolean;
  failedGroup?: boolean;
  failedMembers?: boolean;
  failedBalance?: boolean;
  noBalance?: boolean;
  missingConsumed?: boolean;
  wrongPolicy?: boolean;
  nestedGroup?: boolean;
  pages?: boolean;
  disabled?: boolean;
}

async function load(options: Options = {}) {
  const fetcher = vi.fn<typeof fetch>(async (input) => {
    const url = new URL(String(input));
    const path = url.pathname;
    if (path.endsWith("/serviceSettings")) {
      if (options.failedServices && path.includes(id(3))) return json({ error: { code: "Denied", message: "Service settings denied" } }, 403);
      const row = { serviceId: "cowork", state: "enabled" };
      return json(options.singleton ? row : { value: [row] });
    }
    if (path.endsWith("/assignedGroups")) {
      if (options.failedAssignment && path.includes(id(3))) return json({ error: { code: "Denied", message: "Assignment denied" } }, 403);
      return json({ value: options.noAssignment ? [] : [
        { id: path.includes(id(3)) ? id(4) : id(6) },
        ...(options.extraGroup && path.includes(id(3)) ? [{ id: id(8) }] : []),
      ] });
    }
    if (path.endsWith("/members")) {
      if (options.failedMembers && path.includes(id(4))) return json({ error: { code: "Denied", message: "Members denied" } }, 403);
      if (options.pages && url.searchParams.has("page")) return json({ value: [user(11)] });
      const rows = path.includes(id(4)) ? [user(7)] :
        path.includes(id(8)) || options.overlap ? [user(7)] : [user(10)];
      return json({
        value: [...rows, ...(options.nestedGroup && path.includes(id(4)) ? [{ id: id(9), "@odata.type": "#microsoft.graph.group", displayName: "Nested group" }] : [])],
        ...(options.pages && path.includes(id(4)) ? { "@odata.nextLink": `${url.href}?page=2` } : {}),
      });
    }
    if (path.includes("/groups/")) {
      if (options.failedGroup && path.includes(id(4))) return json({ error: { code: "Denied", message: "Group details denied" } }, 403);
      const groupId = path.split("/").at(-1);
      return json({ id: groupId, displayName: `Live group ${groupId}`, securityEnabled: true, mailEnabled: false, groupTypes: [] });
    }
    if (path.endsWith("/serviceBalances")) {
      if (options.failedBalance && path.includes(id(7))) return json({ error: { code: "NoBalancePermission", message: "Balance denied" } }, 403);
      if (options.noBalance) return json({ value: [] });
      if (options.pages && !url.searchParams.has("page")) return json({
        value: [{ serviceId: "other", consumedQuantity: 42 }],
        "@odata.nextLink": `${url.href}?page=2`,
      });
      const row = {
        serviceId: "cowork", spendingPolicyId: options.wrongPolicy || path.includes(id(10)) ? id(5) : id(3),
        ...(!options.missingConsumed ? { consumedQuantity: 2000 } : {}),
        unit: "copilotCredit", remainingQuantity: limited(8000), configuredPeriodLimit: limited(10000),
        configuredPerUserLimit: unlimited, effectiveCapacityQuantity: limited(10000),
        isConsumptionAllowed: true, accessDecisionReason: "allowed",
        dataStatus: { lastModifiedDateTime: "2026-10-20T10:00:00Z" },
      };
      return json(options.singleton ? row : { value: [row] });
    }
    if (options.failedPolicy && path.includes(id(3))) return json({ error: { code: "Denied", message: "Policy denied" } }, 403);
    return json({
      id: path.includes(id(3)) ? id(3) : id(5), displayName: path.includes(id(3)) ? "Legal Standard policy" : "Legal Enhanced policy",
      status: options.disabled ? "disabled" : "enabled", periodLimit: limited(10000), perUserLimit: unlimited,
      billingMethods: "payAsYouGo", "@odata.type": "#microsoft.graph.selectedGroupsSpendingPolicy",
    });
  });
  const snapshots: BudgetSnapshot[] = [];
  await loadBudgetDetails(budget, "cowork", new GraphClient(async () => "test-token", fetcher),
    new AbortController().signal, (snapshot) => snapshots.push(snapshot));
  const snapshot = snapshots.at(-1);
  if (!snapshot) throw new Error("No detail snapshot produced.");
  return { snapshot, snapshots, fetcher };
}

const markup = (snapshot: BudgetSnapshot) => renderToStaticMarkup(
  createElement(BudgetDetails, { budget, snapshot, serviceId: "cowork" }));

describe("live Graph detail loading", () => {
  it("loads policy config, services, assigned groups, members and consumption without inventory or writes", async () => {
    const { snapshot, snapshots, fetcher } = await load();
    expect(snapshot.tiers.standard.policy.state).toBe("ready");
    expect(snapshot.tiers.enhanced.services.state).toBe("ready");
    expect(snapshot.groups[id(4)].members.state).toBe("ready");
    expect(snapshot.balances[id(7)].state).toBe("ready");
    expect(snapshots[0].tiers.standard.policy.state).toBe("loading");
    expect(snapshots.at(-1)?.tiers.standard.policy.state).toBe("ready");
    expect(fetcher.mock.calls.every(([url, init]) => !String(url).endsWith("/spendingPolicies") && init?.method === undefined)).toBe(true);
    const html = markup(snapshot);
    for (const value of ["Legal Standard policy", "Legal Enhanced policy", "Unlimited Copilot credits", "payAsYouGo", "Assigned in Graph", "Test person 7", "2,000 Copilot credits", "8,000 Copilot credits", "allowed", "Source updated"]) {
      expect(html).toContain(value);
    }
    expect(html).not.toContain("checks passed");
  });
  it("supports the tenant's singleton service shape as well as collections", async () => {
    const { snapshot } = await load({ singleton: true });
    expect(snapshot.tiers.standard.services.state).toBe("ready");
    expect(snapshot.balances[id(7)].state).toBe("ready");
    expect(markup(snapshot)).toContain("2,000 Copilot credits");
  });
  it("does not hide disabled policy config behind a failed check", async () => {
    const { snapshot } = await load({ disabled: true });
    expect(snapshot.tiers.standard.policy.state).toBe("ready");
    expect(markup(snapshot)).toContain("disabled");
  });
  it("keeps directory members and balances when the policy GET fails", async () => {
    const { snapshot } = await load({ failedPolicy: true });
    expect(snapshot.tiers.standard.policy.state).toBe("error");
    expect(snapshot.groups[id(4)].members.state).toBe("ready");
    expect(markup(snapshot)).toContain("Policy denied");
    expect(markup(snapshot)).toContain("Test person 7");
    expect(markup(snapshot)).toContain("2,000 Copilot credits");
  });
  it("keeps other sections when service settings fail", async () => {
    const { snapshot } = await load({ failedServices: true });
    expect(snapshot.tiers.standard.services.state).toBe("error");
    expect(snapshot.tiers.standard.assignedGroups.state).toBe("ready");
    expect(snapshot.balances[id(7)].state).toBe("ready");
  });
  it("labels configured groups as unverified when the assignment read fails", async () => {
    const { snapshot } = await load({ failedAssignment: true });
    const html = markup(snapshot);
    expect(html).toContain("Assignment denied");
    expect(html).toContain("JSON mapping / assignment unverified");
    expect(html).toContain("Test person 7");
  });
  it("does not present a configured group as assigned when Graph returns none", async () => {
    const { snapshot } = await load({ noAssignment: true });
    expect(markup(snapshot)).toContain("Graph returned no assigned groups");
    expect(markup(snapshot)).toContain("JSON mapping only / not assigned");
    expect(markup(snapshot)).not.toContain(">Assigned in Graph<");
  });
  it("loads actual additional assigned groups but deduplicates user balance calls", async () => {
    const { snapshot, fetcher } = await load({ extraGroup: true });
    expect(tierGroupIds(budget, "standard", snapshot)).toEqual([id(4), id(8)]);
    expect(snapshot.groups[id(8)].members.state).toBe("ready");
    expect(fetcher.mock.calls.filter(([url]) => String(url).includes(`/userBalances/${id(7)}/`))).toHaveLength(1);
  });
  it("keeps memberships visible when directory group metadata fails", async () => {
    const { snapshot } = await load({ failedGroup: true });
    expect(snapshot.groups[id(4)].directory.state).toBe("error");
    expect(snapshot.groups[id(4)].members.state).toBe("ready");
    expect(markup(snapshot)).toContain("Group details denied");
    expect(markup(snapshot)).toContain("Test person 7");
  });
  it("marks membership coverage incomplete instead of claiming an empty group", async () => {
    const { snapshot } = await load({ failedMembers: true });
    const html = markup(snapshot);
    expect(html).toContain("Members denied");
    expect(html).toContain("membership coverage is incomplete");
    expect(snapshot.balances[id(7)]).toBeUndefined();
    expect(snapshot.balances[id(10)].state).toBe("ready");
  });
  it("shows a failed user's balance error without hiding another user's reading", async () => {
    const { snapshot } = await load({ failedBalance: true });
    const html = markup(snapshot);
    expect(html).toContain("Balance denied");
    expect(html).toContain("test-request");
    expect(html).toContain("1 of 2 discovered users");
    expect(html).toContain("2,000 Copilot credits");
  });
  it("shows missing service data and missing consumption as unavailable, not zero", async () => {
    const missingService = await load({ noBalance: true });
    expect(markup(missingService.snapshot)).toContain("Consumption is unavailable, not zero");
    const missingQuantity = await load({ missingConsumed: true });
    expect(markup(missingQuantity.snapshot)).toContain("Not reported");
    const balance = missingQuantity.snapshot.balances[id(7)];
    if (balance.state !== "ready" || !balance.data) throw new Error("Expected a current balance row.");
    expect(formatConsumed(balance.data)).toBe("Not reported");
  });
  it("shows actual readings alongside reported-policy and overlap warnings", async () => {
    const mismatched = await load({ wrongPolicy: true });
    expect(markup(mismatched.snapshot)).toContain("Reported policy differs");
    expect(markup(mismatched.snapshot)).toContain("2,000 Copilot credits");
    const overlapping = await load({ overlap: true });
    expect(markup(overlapping.snapshot)).toContain("Member appears in both tiers");
    expect(overlapping.fetcher.mock.calls.filter(([url]) => String(url).includes("/serviceBalances"))).toHaveLength(1);
  });
  it("lists non-user direct members without querying their consumption or expanding nested groups", async () => {
    const { snapshot, fetcher } = await load({ nestedGroup: true });
    expect(markup(snapshot)).toContain("Nested group");
    expect(markup(snapshot)).toContain("Consumption is queried for direct users only");
    expect(fetcher.mock.calls.some(([url]) => String(url).includes(`/userBalances/${id(9)}`))).toBe(false);
    expect(fetcher.mock.calls.some(([url]) => String(url).includes(`/groups/${id(9)}/members`))).toBe(false);
  });
  it("follows member and balance pages without refetching their first pages", async () => {
    const { snapshot, fetcher } = await load({ pages: true });
    expect(snapshot.balances[id(11)].state).toBe("ready");
    expect(markup(snapshot)).toContain("3 of 3 discovered users");
    expect(fetcher.mock.calls.filter(([url]) => String(url).endsWith(`/userBalances/${id(7)}/serviceBalances`))).toHaveLength(1);
    expect(fetcher.mock.calls.some(([url]) => String(url).endsWith(`/userBalances/${id(7)}/serviceBalances?page=2`))).toBe(true);
  });
  it("preserves completed data and explicitly stops pending fields when cancelled", async () => {
    const { snapshots } = await load();
    const partial = snapshots.find((snapshot) => snapshot.tiers.standard.policy.state === "ready");
    if (!partial) throw new Error("Expected a streamed partial snapshot.");
    const stopped = stopLoading(partial, "Loading cancelled");
    expect(stopped.tiers.standard.policy.state).toBe("ready");
    expect(stopped.tiers.enhanced.policy).toEqual({ state: "error", error: "Loading cancelled" });
    expect(partial.tiers.enhanced.policy.state).toBe("loading");
  });
  it("propagates cancellation instead of converting it into an ordinary route failure", async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(loadBudgetDetails(budget, "cowork", new GraphClient(async () => "t"),
      controller.signal, () => undefined)).rejects.toThrow();
  });
});

describe("current quantity and attribution presentation", () => {
  it("distinguishes finite, zero, unlimited, absent, unsupported and invalid quantities", () => {
    expect(formatQuantity(limited(1000))).toBe("1,000 Copilot credits");
    expect(formatQuantity(limited(0))).toBe("0 Copilot credits");
    expect(formatQuantity(unlimited)).toBe("Unlimited Copilot credits");
    expect(formatQuantity(undefined)).toBe("Not reported");
    expect(formatQuantity({ value: 5, unit: "copilotCredit" })).toContain("Unsupported");
    expect(formatQuantity({ ...limited(1), value: -1 })).toContain("Invalid");
    expect(formatQuantity({ ...limited(5), unit: "otherUnit" })).toBe("5 otherUnit");
  });
  it("never derives consumption or remaining credits from allowances", () => {
    expect(formatConsumed({ consumedQuantity: 0, unit: "copilotCredit" })).toBe("0 Copilot credits");
    expect(formatConsumed({ configuredPeriodLimit: limited(1000), remainingQuantity: limited(500) })).toBe("Not reported");
    expect(formatConsumed({ consumedQuantity: -5, unit: "copilotCredit" })).toContain("Invalid");
    expect(formatConsumed({ consumedQuantity: "2000", unit: "copilotCredit" })).toContain("Invalid");
  });
  it("flags missing policy and units without silently attributing a reading", async () => {
    const { snapshot } = await load();
    expect(memberPolicyNote(budget, snapshot, id(7), {})).toContain("not reported");
    expect(memberPolicyNote(budget, snapshot, id(7), { spendingPolicyId: id(3), unit: "other" })).toContain("unexpected");
    expect(initialSnapshot(budget).budgetId).toBe("legal");
  });
});
