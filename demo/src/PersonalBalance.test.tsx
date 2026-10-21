import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { GraphClient } from "./graph";
import { loadPersonalBalance, PersonalBalanceReading, type PersonalReading } from "./PersonalBalance";

const userId = "10000000-0000-0000-0000-000000000001";
const policyId = "10000000-0000-0000-0000-000000000002";
const base = "https://graph.microsoft.com/beta/copilot/costManagement";
const limited = (value: number) => ({ "@odata.type": "#microsoft.graph.limitedCostManagementQuantity", value, unit: "copilotCredit" });
const unlimited = { "@odata.type": "#microsoft.graph.unlimitedCostManagementQuantity", unit: "copilotCredit" };
const balance = {
  serviceId: "cowork", spendingPolicyId: policyId, consumedQuantity: 2000, unit: "copilotCredit",
  remainingQuantity: limited(8000), configuredPeriodLimit: limited(10000),
  configuredPerUserLimit: limited(10000), effectiveCapacityQuantity: limited(10000),
  isConsumptionAllowed: true, accessDecisionReason: "allowed",
  dataStatus: { lastModifiedDateTime: "2026-10-20T10:00:00Z" },
};
const policy = { id: policyId, displayName: "Example Standard policy", status: "enabled", billingMethods: "payAsYouGo" };
const signal = () => new AbortController().signal;
const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { "request-id": "test-request" } });
const markup = (reading: PersonalReading) => renderToStaticMarkup(<PersonalBalanceReading reading={reading} serviceId="cowork" />);

function client(balanceResponse: unknown = { value: [balance] }, policyResponse: unknown = policy, policyStatus = 200) {
  const transport = vi.fn<typeof fetch>(async (input) => {
    if (String(input) === `${base}/userBalances/${userId}/serviceBalances`) return json(balanceResponse);
    if (String(input) === `${base}/spendingPolicies/${policyId}`) return json(policyResponse, policyStatus);
    throw new Error(`Unexpected request: ${input}`);
  });
  return { transport, graph: new GraphClient(async () => "balance-token", transport) };
}

describe("personal Graph balance reads", () => {
  it.each([balance, { value: [balance] }])("reads the signed-in object ID and only its effective policy", async (response) => {
    const { graph, transport } = client(response);
    const reading = await loadPersonalBalance(graph, userId.toUpperCase(), "cowork", signal());
    expect(reading.balance).toEqual(balance);
    expect(reading.policy).toMatchObject({ state: "ready", data: { id: policyId, fields: policy } });
    expect(reading.interactionRequired).toBe(false);
    expect(Number.isFinite(Date.parse(reading.fetchedAt))).toBe(true);
    expect(transport.mock.calls.map(([input]) => String(input))).toEqual([
      `${base}/userBalances/${userId}/serviceBalances`, `${base}/spendingPolicies/${policyId}`,
    ]);
    expect(transport.mock.calls.every(([, init]) => init?.method === undefined)).toBe(true);
    const html = markup(reading);
    for (const value of ["2,000 Copilot credits", "8,000 Copilot credits", "10,000 Copilot credits",
      "Configured per-user limit", "Effective capacity", "Example Standard policy", "payAsYouGo", "Source updated", "Fetched"]) {
      expect(html).toContain(value);
    }
    expect(html).not.toContain("USD");
  });
  it("filters unrelated services and follows all balance pages", async () => {
    const transport = vi.fn<typeof fetch>(async (input) => {
      if (String(input).endsWith("?page=2")) return json({ value: [balance] });
      if (String(input).endsWith("/serviceBalances")) return json({
        value: [{ serviceId: "another-service", consumedQuantity: 999999 }],
        "@odata.nextLink": `${base}/userBalances/${userId}/serviceBalances?page=2`,
      });
      return json(policy);
    });
    const reading = await loadPersonalBalance(new GraphClient(async () => "token", transport), userId, "cowork", signal());
    expect(reading.balance).toEqual(balance);
    expect(transport).toHaveBeenCalledTimes(3);
    expect(markup(reading)).not.toContain("999,999");
  });
  it.each([{ value: [] }, { value: [{ serviceId: "other", consumedQuantity: 0 }] }])("reports absent service balances as unavailable, never zero", async (response) => {
    const { graph, transport } = client(response);
    const reading = await loadPersonalBalance(graph, userId, "cowork", signal());
    expect(reading.balance).toBeNull();
    expect(reading.policy).toBeNull();
    expect(transport).toHaveBeenCalledTimes(1);
    expect(markup(reading)).toContain("Consumption is unavailable, not zero");
    expect(markup(reading)).not.toContain("0 Copilot credits");
  });
  it("rejects ambiguous duplicate service rows rather than choosing a route", async () => {
    const { graph, transport } = client({ value: [balance, balance] });
    await expect(loadPersonalBalance(graph, userId, "cowork", signal())).rejects.toThrow("multiple cowork rows");
    expect(transport).toHaveBeenCalledTimes(1);
  });
  it("rejects malformed balance responses instead of returning an empty balance", async () => {
    await expect(loadPersonalBalance(client({}).graph, userId, "cowork", signal())).rejects.toThrow("Unrecognized service response");
  });
  it("surfaces balance permission errors with the Graph request reference and no invented data", async () => {
    const transport = vi.fn<typeof fetch>(async () => json({ error: { code: "Denied", message: "Balance read denied" } }, 403));
    await expect(loadPersonalBalance(new GraphClient(async () => "token", transport), userId, "cowork", signal()))
      .rejects.toThrow("Denied (HTTP 403): Balance read denied [request test-request]");
    expect(transport).toHaveBeenCalledTimes(1);
  });
  it("keeps the balance visible when the effective policy cannot be read", async () => {
    const { graph } = client({ value: [balance] }, { error: { code: "Denied", message: "Policy denied" } }, 403);
    const reading = await loadPersonalBalance(graph, userId, "cowork", signal());
    expect(reading.balance).toEqual(balance);
    expect(reading.policy?.state).toBe("error");
    expect(reading.interactionRequired).toBe(false);
    expect(markup(reading)).toContain("2,000 Copilot credits");
    expect(markup(reading)).toContain("Policy denied");
    expect(markup(reading)).not.toContain("Example Standard policy");
  });
  it("checks that Graph returns the requested policy ID", async () => {
    const reading = await loadPersonalBalance(client({ value: [balance] }, { ...policy, id: userId }).graph, userId, "cowork", signal());
    expect(reading.policy).toMatchObject({ state: "error", error: expect.stringContaining("different ID") });
    expect(reading.balance).toEqual(balance);
  });
  it.each([undefined, null, ""])("does not invent a policy when spendingPolicyId is %s", async (spendingPolicyId) => {
    const { graph, transport } = client({ ...balance, spendingPolicyId });
    const reading = await loadPersonalBalance(graph, userId, "cowork", signal());
    expect(transport).toHaveBeenCalledTimes(1);
    expect(reading.policy).toBeNull();
    expect(markup(reading)).toContain("No policy or owner is inferred");
  });
  it("does not follow a malformed policy ID", async () => {
    const { graph, transport } = client({ ...balance, spendingPolicyId: "https://example.test/policy" });
    const reading = await loadPersonalBalance(graph, userId, "cowork", signal());
    expect(transport).toHaveBeenCalledTimes(1);
    expect(reading.policy).toMatchObject({ state: "error", error: expect.stringContaining("GUID") });
  });
  it("validates the user object ID before making any network request", async () => {
    const { graph, transport } = client();
    await expect(loadPersonalBalance(graph, "../other-user", "cowork", signal())).rejects.toThrow("GUID");
    expect(transport).not.toHaveBeenCalled();
  });
  it("propagates silent-auth interaction errors without issuing a request", async () => {
    const error = new InteractionRequiredAuthError("consent_required", "Read consent required.");
    const transport = vi.fn<typeof fetch>();
    const graph = new GraphClient(async () => { throw error; }, transport);
    await expect(loadPersonalBalance(graph, userId, "cowork", signal())).rejects.toBe(error);
    expect(transport).not.toHaveBeenCalled();
  });
  it("preserves the balance and exposes required interaction if authentication expires during policy loading", async () => {
    const error = new InteractionRequiredAuthError("interaction_required", "Renew Graph access.");
    const token = vi.fn().mockResolvedValueOnce("balance-token").mockRejectedValueOnce(error);
    const transport = vi.fn<typeof fetch>(async () => json(balance));
    const reading = await loadPersonalBalance(new GraphClient(token, transport), userId, "cowork", signal());
    expect(reading.balance).toEqual(balance);
    expect(reading.policy?.state).toBe("error");
    expect(reading.interactionRequired).toBe(true);
    expect(transport).toHaveBeenCalledTimes(1);
  });
  it("propagates cancellation, including when the policy request is in flight", async () => {
    const controller = new AbortController();
    const transport = vi.fn<typeof fetch>(async (input) => {
      if (String(input).includes("/spendingPolicies/")) { controller.abort(); throw controller.signal.reason; }
      return json(balance);
    });
    await expect(loadPersonalBalance(new GraphClient(async () => "token", transport), userId, "cowork", controller.signal)).rejects.toThrow();
    const aborted = new AbortController();
    aborted.abort();
    const { graph, transport: unused } = client();
    await expect(loadPersonalBalance(graph, userId, "cowork", aborted.signal)).rejects.toThrow();
    expect(unused).not.toHaveBeenCalled();
  });
});

describe("personal credit presentation", () => {
  const reading = (row: Record<string, unknown>): PersonalReading => ({
    balance: row, fetchedAt: "2026-10-20T10:01:00Z", policy: null, interactionRequired: false,
  });
  it("distinguishes zero, unlimited and unreported quantities without deriving usage", () => {
    const html = markup(reading({ ...balance, consumedQuantity: 0, remainingQuantity: unlimited, configuredPerUserLimit: limited(0) }));
    expect(html).toContain("0 Copilot credits");
    expect(html).toContain("Unlimited Copilot credits");
    const missing = markup(reading({ serviceId: "cowork", configuredPerUserLimit: limited(10000) }));
    expect(missing).toContain('<dt>Consumed</dt><dd class="quantity">Not reported</dd>');
    expect(missing).toContain("<dt>Remaining</dt><dd>Not reported</dd>");
    expect(missing).not.toContain(">0 Copilot credits<");
    expect(missing).not.toContain("Unlimited Copilot credits");
  });
  it("retains unknown units and exposes invalid quantities and source timestamps", () => {
    const html = markup(reading({
      ...balance, consumedQuantity: -1, remainingQuantity: { ...limited(500), unit: "otherUnit" },
      configuredPerUserLimit: { value: 10 }, dataStatus: "malformed",
    }));
    expect(html).toContain("Invalid consumed quantity");
    expect(html).toContain("500 otherUnit");
    expect(html).toContain("Unsupported quantity type");
    expect(html).toContain("Balance dataStatus must be an object");
  });
});
