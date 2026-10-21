import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { AccountInfo } from "@azure/msal-browser";
import { DemoAuth } from "./auth";
import { parseConfig } from "./config";
import { ExampleRequests } from "./ExampleRequests";
import { RequestPage } from "./RequestPage";
import { createRequestPreview, expiryLabel, periodEndUtc, requestExpiry, resolveRequestRoute } from "./requests";

const msal = vi.hoisted(() => ({ getActiveAccount: vi.fn(), getAllAccounts: vi.fn() }));
vi.mock("@azure/msal-browser", () => ({
  PublicClientApplication: class {
    getActiveAccount = msal.getActiveAccount;
    getAllAccounts = msal.getAllAccounts;
  },
}));

const id = (n: number) => `10000000-0000-0000-0000-${String(n).padStart(12, "0")}`;
const input = () => ({
  version: 1, tenantId: id(1), clientId: id(2), serviceId: "cowork",
  budgets: [{
    id: "legal", displayName: "Legal", businessArea: "Legal", costCenter: "0001", owners: ["owner@example.test"],
    tiers: {
      standard: { policyId: id(3), groupId: id(4) },
      enhanced: { policyId: id(5), groupId: id(6) },
    },
  }],
  requesterIndex: [{ username: "person@example.test", budgetId: "legal" }],
  exampleRequests: [{
    id: "DEMO-1", requesterUsername: "person@example.test", budgetId: "legal",
    justification: "I need additional capacity for ongoing contract review work.",
    grantKind: "permanent", requestedAt: "2026-10-20T10:00:00Z",
  }],
});
const config = parseConfig(input());
const requester: AccountInfo = {
  username: "person@example.test", tenantId: config.tenantId,
  homeAccountId: "test-home", localAccountId: "test-local", environment: "login.microsoftonline.com",
};
const owner = { ...requester, username: "owner@example.test" };

beforeEach(() => {
  vi.resetAllMocks();
  vi.stubGlobal("window", { location: { origin: "http://localhost:5173" } });
  msal.getActiveAccount.mockReturnValue(requester);
  msal.getAllAccounts.mockReturnValue([]);
});
afterEach(() => {
  vi.unstubAllGlobals();
  vi.useRealTimers();
});

describe("JSON request configuration", () => {
  it("accepts legacy configuration without the new optional data", () => {
    const { requesterIndex: _index, exampleRequests: _examples, ...legacy } = input();
    expect(parseConfig(legacy).requesterIndex).toEqual([]);
    expect(parseConfig(legacy).exampleRequests).toEqual([]);
  });
  it("normalizes index usernames and validates referenced budgets", () => {
    expect(parseConfig({ ...input(), exampleRequests: [], requesterIndex: [{ username: " PERSON@EXAMPLE.TEST ", budgetId: "legal" }] }).requesterIndex?.[0].username).toBe("person@example.test");
    expect(() => parseConfig({ ...input(), requesterIndex: [{ username: "person@example.test", budgetId: "unknown" }] })).toThrow("unknown budget");
  });
  it("rejects malformed lists, duplicate index entries and malformed usernames", () => {
    expect(() => parseConfig({ ...input(), requesterIndex: {} })).toThrow("must be an array");
    expect(() => parseConfig({ ...input(), exampleRequests: {} })).toThrow("must be an array");
    expect(() => parseConfig({ ...input(), requesterIndex: [...input().requesterIndex, ...input().requesterIndex] })).toThrow("Duplicate requester");
    expect(() => parseConfig({ ...input(), requesterIndex: [{ username: "person", budgetId: "legal" }] })).toThrow("UPN");
  });
  it("requires indexed requesters and valid sample fields", () => {
    const sample = input().exampleRequests[0];
    expect(() => parseConfig({ ...input(), exampleRequests: [{ ...sample, requesterUsername: "other@example.test" }] })).toThrow("matching requester index");
    expect(() => parseConfig({ ...input(), exampleRequests: [sample, sample] })).toThrow("Duplicate example");
    expect(() => parseConfig({ ...input(), exampleRequests: [{ ...sample, grantKind: "weekly" }] })).toThrow("grantKind");
    expect(() => parseConfig({ ...input(), exampleRequests: [{ ...sample, justification: "too short" }] })).toThrow("20-2000");
    expect(() => parseConfig({ ...input(), exampleRequests: [{ ...sample, requestedAt: "not-a-date" }] })).toThrow("UTC");
  });
});

describe("identity to budget to owner routing", () => {
  it("routes a requester without knowing their spend policy or calling Graph", () => {
    const fetcher = vi.fn();
    vi.stubGlobal("fetch", fetcher);
    const route = resolveRequestRoute(config, { ...requester, username: " PERSON@EXAMPLE.TEST " });
    expect(route.state).toBe("ready");
    if (route.state !== "ready") throw new Error("Expected owner routing.");
    expect(route.budget.id).toBe("legal");
    expect(route.approvers).toEqual(["owner@example.test"]);
    expect(fetcher).not.toHaveBeenCalled();
  });
  it("never guesses for unknown users, missing indexes or a different tenant", () => {
    expect(resolveRequestRoute(config, { ...requester, username: "unknown@example.test" }).state).toBe("needs-routing");
    expect(resolveRequestRoute({ ...config, requesterIndex: [] }, requester).state).toBe("needs-routing");
    expect(resolveRequestRoute(config, { ...requester, tenantId: id(99) }).state).toBe("needs-routing");
  });
  it("flags ambiguous budget mappings instead of choosing the first owner", () => {
    const ambiguous = {
      ...config,
      budgets: [...config.budgets, { ...config.budgets[0], id: "other" }],
      requesterIndex: [...config.requesterIndex ?? [], { username: requester.username, budgetId: "other" }],
    };
    const route = resolveRequestRoute(ambiguous, requester);
    expect(route.state).toBe("needs-routing");
    if (route.state !== "needs-routing") throw new Error("Expected triage.");
    expect(route.reason).toContain("More than one");
  });
  it("resolves current ownership, not a stale approver saved with the request", () => {
    const changed = { ...config, budgets: [{ ...config.budgets[0], owners: ["new.owner@example.test"] }] };
    const route = resolveRequestRoute(changed, requester);
    expect(route.state === "ready" ? route.approvers : []).toEqual(["new.owner@example.test"]);
  });
  it("requires valid owners and permits owner self-approval for the MVP", () => {
    for (const owners of [[], ["Demo owner"]]) {
      expect(resolveRequestRoute({ ...config, budgets: [{ ...config.budgets[0], owners }] }, requester).state).toBe("needs-routing");
    }
    const route = resolveRequestRoute({
      ...config, budgets: [{ ...config.budgets[0], owners: ["person@example.test", "OWNER@EXAMPLE.TEST", "owner@example.test"] }],
    }, requester);
    expect(route.state === "ready" ? route.approvers : []).toEqual(["person@example.test", "owner@example.test"]);
  });
});

describe("request preview and billing duration", () => {
  it("captures the UTC calendar-month boundary including leap years and December", () => {
    expect(periodEndUtc(new Date("2026-10-31T23:59:59Z"))).toBe("2026-11-01T00:00:00.000Z");
    expect(periodEndUtc(new Date("2026-12-31T23:59:59Z"))).toBe("2027-01-01T00:00:00.000Z");
    expect(periodEndUtc(new Date("2028-02-29T10:00:00Z"))).toBe("2028-03-01T00:00:00.000Z");
    expect(periodEndUtc(new Date("2026-11-01T00:00:00Z"))).toBe("2026-12-01T00:00:00.000Z");
    expect(() => periodEndUtc(new Date("invalid"))).toThrow("valid request timestamp");
    expect(requestExpiry("permanent", new Date())).toBeNull();
  });
  it("creates an honest temporary preview with the authenticated requester and no submission side effects", () => {
    const preview = createRequestPreview(config, requester, `  ${input().exampleRequests[0].justification}  `, "billingPeriod", new Date("2026-10-20T12:00:00Z"));
    expect(preview.requester).toBe(requester.username);
    expect(preview.justification).toBe(input().exampleRequests[0].justification);
    expect(preview.expiresAt).toBe("2026-11-01T00:00:00.000Z");
    expect(expiryLabel(preview.expiresAt)).toContain("UTC");
    expect(createRequestPreview(config, requester, preview.justification, "permanent").expiresAt).toBeNull();
    expect(config.exampleRequests).toHaveLength(1);
  });
  it("requires meaningful bounded justification", () => {
    for (const value of ["", " ".repeat(50), "short", "x".repeat(2001)]) {
      expect(() => createRequestPreview(config, requester, value, "permanent")).toThrow("20-2000");
    }
  });
});

describe("requester and example inbox rendering", () => {
  it("shows no form or routing metadata before sign-in", () => {
    msal.getActiveAccount.mockReturnValue(null);
    const html = renderToStaticMarkup(<RequestPage config={config} auth={new DemoAuth(config)} />);
    expect(html).toContain("Sign in with Microsoft");
    expect(html).not.toContain("Legal");
    expect(html).not.toContain("owner@example.test");
    expect(html).not.toContain("textarea");
  });
  it("shows the form and resolved owner without requiring a balance read", () => {
    const html = renderToStaticMarkup(<RequestPage config={config} auth={new DemoAuth(config)} />);
    for (const text of ["Business justification", "Rest of this billing period", "Permanent increase", "owner@example.test", "Preview request and routing", "Preview, not submission"]) expect(html).toContain(text);
    expect(html).not.toContain(config.budgets[0].tiers.standard.policyId);
    expect(html).not.toContain("consumedQuantity");
  });
  it("shows only examples belonging to the current owner's budget", () => {
    const html = renderToStaticMarkup(<ExampleRequests config={config} budget={config.budgets[0]} account={owner} />);
    expect(html).toContain("DEMO-1");
    expect(html).toContain("Preview approval");
    expect(html).toContain("synthetic examples");
    expect(html).toContain("?view=request");
    expect(renderToStaticMarkup(<ExampleRequests config={config} budget={config.budgets[0]} account={requester} />)).toBe("");
  });
  it("disables preview decisions on a temporary example from a previous billing period", () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-11-02T00:00:00Z"));
    const expired = { ...config, exampleRequests: config.exampleRequests?.map((request) => ({ ...request, grantKind: "billingPeriod" as const })) };
    const html = renderToStaticMarkup(<ExampleRequests config={expired} budget={config.budgets[0]} account={owner} />);
    expect(html).toContain("captured temporary period has ended");
    expect(html).toContain('type="button" disabled=""');
  });
  it("allows self-approval when the requester is the sole owner", () => {
    const self = { ...config, budgets: [{ ...config.budgets[0], owners: [requester.username] }] };
    const html = renderToStaticMarkup(<ExampleRequests config={self} budget={self.budgets[0]} account={requester} />);
    expect(html).not.toContain("cannot approve your own request");
    expect(html).toContain('type="button">Preview approval');
  });
});
