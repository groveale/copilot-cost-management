import { beforeEach, expect, it, vi } from "vitest";
import type { OnboardingRecord } from "../../api/src/onboarding";
import { RequestApi } from "./api";
import { GraphClient } from "./graph";
import { object } from "./config";
import { resourceNames, runProvisioning, templateSettings } from "./provisioning";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const initial = (): OnboardingRecord => ({
  version: 1, kind: "onboarding", id: "a".repeat(64), tenantId: id(1), requesterId: id(2), operationId: id(3),
  name: "Finance BU", costCenter: "CC200", monthlyBudgetUsd: "20000", justification: "A governed monthly allocation for our team.",
  budgetId: `bu-${"a".repeat(24)}`, createdAt: "2026-10-21T09:00:00Z", status: "provisioning",
  approvedCredits: 2_000_000, standard: { credits: 1_000_000, perUserCredits: 5000, billingMethods: "prepaidCredits" },
  enhanced: { credits: 1_000_000, perUserCredits: 50000, billingMethods: "prepaidCredits" },
  resources: {}, intents: [], etag: "1",
});
let work: OnboardingRecord;
let api: RequestApi;
let graph: GraphClient;
let etag: number;
let sequence: number;
let ambiguousGroup: boolean;
let forbiddenGroup: boolean;
let stalePatch: boolean;
let groups: Map<string, Record<string, unknown>>;
let policies: Map<string, Record<string, unknown>>;
let audiences: Map<string, string[]>;
let patches: string[];
let groupPosts: number;
let mutationTotals: number[];
let actions: string[];
const response = (data: unknown, status = 200, tag?: string) =>
  new Response(status === 204 ? null : JSON.stringify(data), { status, headers: tag ? { ETag: tag } : {} });
beforeEach(() => {
  work = initial(); etag = 1; sequence = 10; ambiguousGroup = false; forbiddenGroup = false; stalePatch = false;
  groups = new Map(); policies = new Map(); audiences = new Map(); patches = []; groupPosts = 0; mutationTotals = []; actions = [];
  api = new RequestApi("https://example.test", async () => "synthetic-api-token");
  vi.spyOn(api, "onboardingGet").mockImplementation(async () => structuredClone(work));
  vi.spyOn(api, "onboardingWork").mockImplementation(async (_id, update) => {
    if (update.etag !== work.etag) throw new Error("Stale progress ETag.");
    if (update.action !== "claim" && update.leaseId !== work.lease?.id) throw new Error("Lease changed.");
    actions.push(update.action);
    if (update.action === "claim") work.lease = { id: id(4), actor: id(5), until: "2026-10-21T12:00:00Z" };
    if (update.action === "intent" && update.resource) work.intents.push(update.resource);
    if (update.action === "clear-intent") work.intents = work.intents.filter((key) => key !== update.resource);
    if (update.action === "resource" && update.resource && update.resourceId) work.resources[update.resource] = update.resourceId;
    if (update.action === "release") delete work.lease;
    if (update.action === "complete") { work.status = "completed"; delete work.lease; }
    work.etag = String(++etag);
    return structuredClone(work);
  });
  graph = new GraphClient(async () => "synthetic-graph-token", vi.fn<typeof fetch>(async (input, init) => {
    const url = new URL(String(input));
    const method = init?.method ?? "GET";
    const body = init?.body ? object(JSON.parse(String(init.body)), "body") : {};
    const filter = /eq '([^']+)'/.exec(url.searchParams.get("$filter") ?? "")?.[1];
    if (url.pathname === "/v1.0/groups") {
      if (method === "GET") return response({ value: [...groups.values()].filter((group) => group.mailNickname === filter) });
      groupPosts++;
      if (forbiddenGroup) return response({ error: { code: "Authorization_RequestDenied", message: "No group permission." } }, 403);
      const key = id(++sequence); const group = { ...body, id: key }; groups.set(key, group);
      if (ambiguousGroup) { ambiguousGroup = false; throw new TypeError("Response lost after create."); }
      return response(group, 201);
    }
    const groupMatch = /^\/v1.0\/groups\/([^/]+)(\/owners)?$/.exec(url.pathname);
    if (groupMatch) return groupMatch[2] ? response({ value: [{ id: work.requesterId }] }) : response(groups.get(groupMatch[1]));
    const base = "/beta/copilot/costManagement/spendingPolicies";
    if (url.pathname === base) {
      if (method === "GET") return response({ value: [...policies.values()].filter((policy) => policy.displayName === filter) });
      const key = id(++sequence);
      const policy = { ...body, id: key, status: "disabled", etag: 1 };
      policies.set(key, policy); audiences.set(key, []);
      return response(policy, 201, "1");
    }
    const match = /^\/beta\/copilot\/costManagement\/spendingPolicies\/([^/]+)(.*)$/.exec(url.pathname);
    if (!match || !policies.has(match[1])) throw new Error(`Unexpected Graph call ${method} ${url.pathname}`);
    const policy = policies.get(match[1])!;
    if (match[2] === "/serviceSettings") return response({ value: [{ serviceId: "cowork", state: "enabled" }] });
    if (match[2].endsWith("/assignedGroups")) return response({ value: (audiences.get(match[1]) ?? []).map((id) => ({ id })) });
    if (match[2].endsWith("/assignedGroups/$ref")) {
      if (new Headers(init?.headers).get("if-match") !== String(policy.etag)) return response({ error: { code: "preconditionFailed" } }, 412);
      audiences.set(match[1], [String(body["@odata.id"]).split("/").at(-1)!]);
      policy.status = "enabled"; policy.etag = Number(policy.etag) + 1;
      return response(null, 204, String(policy.etag));
    }
    if (method === "PATCH") {
      if (stalePatch || new Headers(init?.headers).get("if-match") !== String(policy.etag)) return response({ error: { code: "preconditionFailed", message: "Policy changed." } }, 412);
      patches.push(match[1]); Object.assign(policy, body); policy.etag = Number(policy.etag) + 1;
      mutationTotals.push([...policies.values()].reduce((total, item) => total + Number(object(item.periodLimit, "limit").value), 0));
      return response(policy, 200, String(policy.etag));
    }
    return response(policy, 200, String(policy.etag));
  }));
});
const run = () => runProvisioning(api, graph, work, new AbortController().signal, () => undefined);
it("provisions two groups owned by the requester, two one-million-credit policies, audiences and completion", async () => {
  const result = await run();
  expect(result.status).toBe("completed"); expect(groups.size).toBe(2); expect(policies.size).toBe(2);
  for (const group of groups.values()) expect(group["owners@odata.bind"]).toEqual([`https://graph.microsoft.com/v1.0/users/${work.requesterId}`]);
  for (const policy of policies.values()) {
    expect(policy.periodLimit).toMatchObject({ value: 1_000_000, unit: "copilotCredit" });
    expect(policy.status).toBe("enabled");
  }
  expect([...audiences.values()].every((value) => value.length === 1)).toBe(true);
  expect(actions.at(-1)).toBe("complete"); expect(actions).toContain("heartbeat");
});
it("reconciles a lost group-create response by finding the existing group without creating a duplicate", async () => {
  ambiguousGroup = true;
  await expect(run()).rejects.toThrow("Response lost");
  expect(work.status).toBe("provisioning"); expect(groups.size).toBe(1); expect(work.intents).toContain("standardGroup");
  await run();
  expect(groups.size).toBe(2); expect(groupPosts).toBe(2); expect(work.status).toBe("completed");
});
it("does not reissue an uncertain POST while the resource is not visible", async () => {
  work.intents.push("standardGroup");
  await expect(run()).rejects.toThrow("uncertain outcome");
  expect(groupPosts).toBe(0); expect(work.status).toBe("provisioning");
});
it("permits a safe explicit retry after Graph definitively rejects creation", async () => {
  forbiddenGroup = true;
  await expect(run()).rejects.toThrow("No group permission");
  expect(groupPosts).toBe(1); expect(work.intents).not.toContain("standardGroup");
  forbiddenGroup = false; await run();
  expect(groups.size).toBe(2); expect(groupPosts).toBe(3);
});
it("lowers Enhanced to 500K before raising Standard to 1.5M without exceeding the 2M ceiling", async () => {
  await run();
  work = { ...work, kind: "allocation", id: `allocation-${work.budgetId}`, status: "provisioning",
    standard: { ...work.standard, credits: 1_500_000 }, enhanced: { ...work.enhanced, credits: 500_000 } };
  await run();
  expect(patches).toEqual([work.resources.enhancedPolicy, work.resources.standardPolicy]);
  expect(mutationTotals).toEqual([1_500_000, 2_000_000]);
  expect(work.status).toBe("completed");
});
it("checks audience drift before mutating an allocation", async () => {
  await run();
  work = { ...work, kind: "allocation", status: "provisioning", standard: { ...work.standard, credits: 500_000 },
    enhanced: { ...work.enhanced, credits: 1_500_000 } };
  audiences.set(work.resources.standardPolicy!, [id(99)]);
  await expect(run()).rejects.toThrow("assignments differ");
  expect(patches).toEqual([]);
});
it("surfaces stale Graph ETags without retrying or reporting completion", async () => {
  await run();
  work = { ...work, kind: "allocation", status: "provisioning",
    standard: { ...work.standard, credits: 500_000 }, enhanced: { ...work.enhanced, credits: 1_500_000 } };
  stalePatch = true;
  await expect(run()).rejects.toThrow("preconditionFailed");
  expect(work.status).toBe("provisioning"); expect(patches).toEqual([]);
});
it("saves monthly and per-user limits on only the selected policy and leaves billing/other policy unchanged", async () => {
  await run();
  const other = structuredClone(policies.get(work.resources.enhancedPolicy!)!);
  work = { ...work, kind: "allocation", status: "provisioning",
    standard: { ...work.standard, credits: 900_000 }, perUserEdit: { tier: "standard", credits: 12000 } };
  await run();
  expect(patches).toEqual([work.resources.standardPolicy]);
  expect(policies.get(work.resources.standardPolicy!)!.perUserLimit).toMatchObject({ value: 12000, unit: "copilotCredit" });
  expect(policies.get(work.resources.standardPolicy!)!.billingMethods).toBe("prepaidCredits");
  expect(policies.get(work.resources.enhancedPolicy!)).toEqual(other);
  expect(work.status).toBe("completed");
});
it("preserves an explicit unlimited per-user edit and reconciles an already-applied saved edit without another patch", async () => {
  await run();
  work = { ...work, kind: "allocation", status: "provisioning", perUserEdit: { tier: "enhanced", credits: null } };
  await run();
  expect(policies.get(work.resources.enhancedPolicy!)!.perUserLimit).toEqual({
    "@odata.type": "#microsoft.graph.unlimitedCostManagementQuantity", unit: "copilotCredit",
  });
  expect(patches).toEqual([work.resources.enhancedPolicy]);
  work.status = "provisioning";
  await run();
  expect(patches).toEqual([work.resources.enhancedPolicy]);
  expect(work.status).toBe("completed");
});
it("does not restore a different policy's externally changed allocation", async () => {
  await run();
  work = { ...work, kind: "allocation", status: "provisioning",
    standard: { ...work.standard, credits: 900_000 }, perUserEdit: { tier: "standard", credits: 12000 } };
  policies.get(work.resources.enhancedPolicy!)!.periodLimit = {
    "@odata.type": "#microsoft.graph.limitedCostManagementQuantity", value: 800_000, unit: "copilotCredit",
  };
  await expect(run()).rejects.toThrow("other policy allocation changed");
  expect(patches).toEqual([]);
  expect(work.status).toBe("provisioning");
});
it("rejects wrong units, preserves explicit unlimited per-user limits and creates valid policy names", () => {
  expect(() => templateSettings({ billingMethods: "prepaidCredits", perUserLimit: { unit: "USD" } })).toThrow();
  expect(templateSettings({ billingMethods: "prepaidCredits",
    perUserLimit: { "@odata.type": "#microsoft.graph.unlimitedCostManagementQuantity", unit: "copilotCredit" } }).perUserCredits).toBeNull();
  const name = resourceNames({ ...work, name: "Finance & R&D / (West)" }, "enhanced").policy;
  expect(name).toMatch(/^[A-Za-z0-9 ]{10,64}$/);
});
