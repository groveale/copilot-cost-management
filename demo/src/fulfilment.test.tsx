import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import type { FulfilmentUpdate, RequestRecord } from "../../api/src/contracts";
import { RequestApi } from "./api";
import { RequestCard } from "./HostedPortal";
import { applyMembershipMove, requestStatusLabel, runFulfilment } from "./fulfilment";
import { GraphClient } from "./graph";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const record: RequestRecord = {
  version: 1, id: "a".repeat(64), tenantId: id(1), requesterId: id(2), budgetId: "example",
  operationId: id(9), justification: "Additional credits for a customer demonstration.", grantKind: "permanent",
  createdAt: "2026-10-01T00:00:00Z", expiresAt: null, status: "approved", etag: "e1",
  fulfilment: { state: "pending", standard: { policyId: id(4), groupId: id(5) }, enhanced: { policyId: id(6), groupId: id(7) } },
};
const signal = () => new AbortController().signal;
const json = (data: unknown, status = 200) => new Response(JSON.stringify(data), { status });
interface Options {
  failAdd?: boolean; uncertainAdd?: boolean; failRemove?: boolean; uncertainRemove?: boolean;
  noConfirmation?: boolean; denyMembers?: boolean; pages?: boolean; badAudience?: boolean;
  disabledPolicy?: boolean; dynamic?: boolean; synced?: boolean; roleGroup?: boolean;
  disabledUser?: boolean; inheritedStandard?: boolean; inheritedEnhanced?: boolean; badCheck?: boolean;
}
function graphFixture(options: Options = {}, memberships = { standard: true, enhanced: false }) {
  const writes: string[] = [];
  const transport = vi.fn<typeof fetch>(async (input, init) => {
    const url = new URL(String(input));
    const path = url.pathname;
    if (path.endsWith("/checkMemberGroups")) return json({ value: options.badCheck ? [id(99)] : [
      ...(memberships.standard || options.inheritedStandard ? [id(5)] : []),
      ...(memberships.enhanced || options.inheritedEnhanced ? [id(7)] : []),
    ] });
    if (path.endsWith("/members/$ref")) {
      writes.push("add-enhanced");
      expect(path).toBe(`/v1.0/groups/${id(7)}/members/$ref`);
      expect(JSON.parse(String(init?.body))).toEqual({ "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${id(2)}` });
      if (options.failAdd) return json({ error: { code: "Denied", message: "Add denied" } }, 403);
      memberships.enhanced = true;
      if (options.uncertainAdd) throw new Error("Add response lost");
      return new Response(null, { status: 204 });
    }
    if (init?.method === "DELETE") {
      writes.push("remove-standard");
      expect(path).toBe(`/v1.0/groups/${id(5)}/members/${id(2)}/$ref`);
      expect(init.body).toBeUndefined();
      if (options.failRemove) return json({ error: { code: "Denied", message: "Remove denied" } }, 403);
      memberships.standard = false;
      if (options.uncertainRemove) throw new Error("Remove response lost");
      return new Response(null, { status: 204 });
    }
    if (path.endsWith("/members")) {
      if (options.denyMembers) return json({ error: { code: "Denied", message: "Membership read denied" } }, 403);
      if (options.pages && !url.searchParams.has("page")) return json({ value: [],
        "@odata.nextLink": `${url.href}&page=2` });
      const member = path.includes(id(5)) ? memberships.standard : memberships.enhanced && !options.noConfirmation;
      return json({ value: member ? [{ id: id(2) }] : [] });
    }
    if (path.endsWith("/assignedGroups")) return json({ value: [{ id: options.badAudience ? id(99) : path.includes(id(4)) ? id(5) : id(7) }] });
    if (path.endsWith("/serviceSettings")) return json({ serviceId: "cowork", state: "enabled" });
    if (path.includes("/spendingPolicies/")) return json({
      id: path.split("/").at(-1), "@odata.type": "#microsoft.graph.selectedGroupsSpendingPolicy",
      status: options.disabledPolicy ? "disabled" : "enabled",
    });
    if (path.includes("/groups/")) return json({
      id: path.split("/").at(-1), securityEnabled: true, mailEnabled: false,
      groupTypes: options.dynamic ? ["DynamicMembership"] : [],
      onPremisesSyncEnabled: options.synced ? true : null, isAssignableToRole: options.roleGroup ?? false,
    });
    if (path === `/v1.0/users/${id(2)}`) return json({ id: id(2), accountEnabled: !options.disabledUser });
    throw new Error(`Unexpected Graph request ${path}`);
  });
  return { graph: new GraphClient(async () => "synthetic-token", transport), transport, writes, memberships };
}

describe("verified Standard-to-Enhanced membership move", () => {
  it("adds/verifies Enhanced before removing Standard, using only the member $ref", async () => {
    const fixture = graphFixture();
    const authorize = vi.fn(async () => undefined);
    await applyMembershipMove(fixture.graph, record, "cowork", signal(), authorize);
    expect(fixture.writes).toEqual(["add-enhanced", "remove-standard"]);
    expect(fixture.memberships).toEqual({ standard: false, enhanced: true });
    expect(authorize).toHaveBeenCalledTimes(2);
    expect(fixture.transport.mock.calls.some(([, init]) => init?.method === "PATCH")).toBe(false);
  });
  it("follows complete direct-membership pages during verification", async () => {
    const fixture = graphFixture({ pages: true });
    await applyMembershipMove(fixture.graph, record, "cowork", signal(), async () => undefined);
    expect(fixture.writes).toEqual(["add-enhanced", "remove-standard"]);
    expect(fixture.transport.mock.calls.some(([url]) => String(url).includes("page=2"))).toBe(true);
  });
  it("skips already-confirmed writes when the user is already exclusively Enhanced", async () => {
    const fixture = graphFixture({}, { standard: false, enhanced: true });
    await applyMembershipMove(fixture.graph, record, "cowork", signal(), async () => undefined);
    expect(fixture.writes).toEqual([]);
  });
  it("does not remove Standard if Enhanced addition is denied or cannot be confirmed", async () => {
    for (const options of [{ failAdd: true }, { noConfirmation: true }]) {
      const fixture = graphFixture(options);
      await expect(applyMembershipMove(fixture.graph, record, "cowork", signal(), async () => undefined)).rejects.toThrow();
      expect(fixture.writes).toEqual(["add-enhanced"]);
      expect(fixture.memberships.standard).toBe(true);
    }
  });
  it.each(["uncertainAdd", "uncertainRemove"] as const)("reconciles %s through fresh reads rather than blindly repeating writes", async (failure) => {
    const options: Options = { [failure]: true };
    const fixture = graphFixture(options);
    await expect(applyMembershipMove(fixture.graph, record, "cowork", signal(), async () => undefined)).rejects.toThrow("response lost");
    options[failure] = false;
    await applyMembershipMove(fixture.graph, record, "cowork", signal(), async () => undefined);
    expect(fixture.writes).toEqual(["add-enhanced", "remove-standard"]);
    expect(fixture.memberships).toEqual({ standard: false, enhanced: true });
  });
  it.each([
    { badAudience: true }, { disabledPolicy: true }, { dynamic: true }, { synced: true },
    { roleGroup: true }, { disabledUser: true }, { denyMembers: true }, { badCheck: true },
  ])("refuses an unsafe/unverifiable preflight without membership writes: %j", async (options) => {
    const fixture = graphFixture(options);
    await expect(applyMembershipMove(fixture.graph, record, "cowork", signal(), async () => undefined)).rejects.toThrow();
    expect(fixture.writes).toEqual([]);
  });
  it("blocks inherited Standard membership before making changes", async () => {
    const fixture = graphFixture({ inheritedStandard: true }, { standard: false, enhanced: false });
    await expect(applyMembershipMove(fixture.graph, record, "cowork", signal(), async () => undefined)).rejects.toThrow("inherited");
    expect(fixture.writes).toEqual([]);
  });
  it("does not claim success when another inherited Standard path survives the direct removal", async () => {
    const fixture = graphFixture({ inheritedStandard: true });
    await expect(applyMembershipMove(fixture.graph, record, "cowork", signal(), async () => undefined)).rejects.toThrow("inherited tier membership");
    expect(fixture.writes).toEqual(["add-enhanced", "remove-standard"]);
  });
  it("reauthorizes each mutation and stops before any write when server routing/ownership changes", async () => {
    const fixture = graphFixture();
    await expect(applyMembershipMove(fixture.graph, record, "cowork", signal(), async () => { throw new Error("Routing changed"); }))
      .rejects.toThrow("Routing changed");
    expect(fixture.writes).toEqual([]);
  });
  it("does not grant expired temporary requests or run cancelled work", async () => {
    const fixture = graphFixture();
    await expect(applyMembershipMove(fixture.graph, { ...record, expiresAt: new Date(0).toISOString() }, "cowork", signal(), async () => undefined))
      .rejects.toThrow("expired");
    const controller = new AbortController();
    controller.abort();
    await expect(applyMembershipMove(fixture.graph, record, "cowork", controller.signal, async () => undefined)).rejects.toThrow();
    expect(fixture.transport).not.toHaveBeenCalled();
  });
});

function apiFixture(input: RequestRecord = record, failComplete = false) {
  let current = structuredClone(input);
  const actions: string[] = [];
  const transport = vi.fn<typeof fetch>(async (_url, init) => {
    const update: FulfilmentUpdate = JSON.parse(String(init?.body));
    actions.push(update.action);
    if (update.action === "complete" && failComplete) throw new Error("Progress save failed");
    if (update.etag !== current.etag) return json({ message: "Stale ETag" }, 409);
    current = { ...current, etag: `${current.etag}-next` };
    if (current.fulfilment?.state !== "applied") {
      current.fulfilment = { ...current.fulfilment!, state: update.action === "complete" ? "applied" :
        update.action === "release" ? "needs-attention" : "applying" };
      if (update.action === "complete" || update.action === "release") delete current.fulfilment.lease;
      else current.fulfilment.lease = { id: id(90), actor: id(3), until: new Date(Date.now() + 300000).toISOString() };
      if (update.action === "complete") { current.fulfilment.appliedAt = new Date().toISOString(); current.fulfilment.appliedBy = id(3); }
    }
    return json(current);
  });
  return { api: new RequestApi("https://synthetic.test", async () => "api-token", transport), actions };
}
describe("durable browser fulfilment coordinator", () => {
  it("claims, renews before both writes and completes only after final verification", async () => {
    const { api, actions } = apiFixture();
    const fixture = graphFixture();
    const progress: RequestRecord[] = [];
    const saved = await runFulfilment(api, fixture.graph, record, "cowork", signal(), (value) => progress.push(value));
    expect(actions).toEqual(["claim", "heartbeat", "heartbeat", "complete"]);
    expect(saved.fulfilment?.state).toBe("applied");
    expect(progress[0].fulfilment?.state).toBe("applying");
    expect(progress.at(-1)?.fulfilment?.state).toBe("applied");
  });
  it("records needs-attention after partial removal failure, never completion", async () => {
    const { api, actions } = apiFixture();
    const fixture = graphFixture({ failRemove: true });
    const progress: RequestRecord[] = [];
    await expect(runFulfilment(api, fixture.graph, record, "cowork", signal(), (value) => progress.push(value))).rejects.toThrow("Remove denied");
    expect(actions.at(-1)).toBe("release");
    expect(actions).not.toContain("complete");
    expect(progress.at(-1)?.fulfilment?.state).toBe("needs-attention");
    expect(fixture.memberships).toEqual({ standard: true, enhanced: true });
  });
  it("does not issue Graph calls for an already-applied idempotent result", async () => {
    const applied: RequestRecord = { ...record, fulfilment: { ...record.fulfilment!, state: "applied" } };
    const { api } = apiFixture(applied);
    const fixture = graphFixture();
    expect((await runFulfilment(api, fixture.graph, applied, "cowork", signal(), () => undefined)).fulfilment?.state).toBe("applied");
    expect(fixture.transport).not.toHaveBeenCalled();
  });
  it("surfaces an unconfirmed progress save even when the group move succeeded", async () => {
    const { api, actions } = apiFixture(record, true);
    const fixture = graphFixture();
    await expect(runFulfilment(api, fixture.graph, record, "cowork", signal(), () => undefined)).rejects.toThrow("Progress save failed");
    expect(fixture.memberships).toEqual({ standard: false, enhanced: true });
    expect(actions.at(-1)).toBe("release");
  });
});

describe("approval status and explicit recovery UI", () => {
  const html = (value: RequestRecord, inbox = true) => renderToStaticMarkup(
    <RequestCard record={value} api={apiFixture().api} inbox={inbox} onDecision={() => undefined} />);
  it("shows an explicit apply action for existing approved requests without running anything on render", () => {
    const legacy = { ...record, fulfilment: undefined };
    expect(html(legacy)).toContain("Apply Enhanced upgrade");
    expect(html(legacy)).not.toContain("manual fulfilment required");
    expect(html(legacy, false)).not.toContain("Apply Enhanced upgrade");
    expect(html({ ...record, status: "pending", fulfilment: undefined })).toContain("Approve and move to Enhanced");
  });
  it("offers retry for partial outcomes but disables active leased work", () => {
    expect(html({ ...record, fulfilment: { ...record.fulfilment!, state: "needs-attention" } })).toContain("Retry Enhanced upgrade");
    const active: RequestRecord = { ...record, fulfilment: { ...record.fulfilment!, state: "applying",
      lease: { id: id(90), actor: id(3), until: new Date(Date.now() + 300000).toISOString() } } };
    expect(html(active)).toContain('disabled="">Apply Enhanced upgrade');
  });
  it("labels manual temporary return without promising automatic expiry or applying expired requests", () => {
    const expired = { ...record, expiresAt: "2026-01-01T00:00:00Z", grantKind: "billingPeriod" as const };
    const applied = { ...expired, fulfilment: { ...record.fulfilment!, state: "applied" as const } };
    expect(requestStatusLabel(applied, Date.parse(expired.expiresAt))).toContain("manual return overdue");
    expect(requestStatusLabel(applied, Date.parse(expired.expiresAt) - 1)).toContain("manual return required");
    expect(html(applied)).toContain("No automatic reversion is scheduled");
    expect(html(expired)).not.toContain("Apply Enhanced upgrade");
    expect(html(applied)).not.toContain("Apply Enhanced upgrade");
  });
});
