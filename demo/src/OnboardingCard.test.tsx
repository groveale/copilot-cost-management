import { renderToStaticMarkup } from "react-dom/server";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { OnboardingRecord } from "../../api/src/onboarding";
import { RequestApi } from "./api";
import { DemoAuth } from "./auth";
import { parseConfig } from "./config";
import { OnboardingCard } from "./Onboarding";

vi.mock("@azure/msal-browser", async (importOriginal) => ({
  ...await importOriginal<typeof import("@azure/msal-browser")>(),
  PublicClientApplication: class {},
}));
afterEach(() => vi.unstubAllGlobals());
const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const record: OnboardingRecord = {
  version: 1, kind: "onboarding", id: "a".repeat(64), tenantId: id(1), requesterId: id(3), operationId: id(4),
  name: "Legal BU", costCenter: "LEGAL-01", monthlyBudgetUsd: "20000", justification: "Monthly capacity for our contract review team.",
  budgetId: `bu-${"a".repeat(24)}`, createdAt: "2026-10-21T09:00:00Z", status: "pending",
  approvedCredits: 2000000,
  standard: { credits: 1000000, perUserCredits: 10000, billingMethods: "payAsYouGo" },
  enhanced: { credits: 1000000, perUserCredits: 25000, billingMethods: "payAsYouGo" },
  resources: {}, intents: [], etag: "specific-etag",
};
function render(value: OnboardingRecord, admin = false) {
  vi.stubGlobal("window", { location: { origin: "http://localhost:5173" } });
  const config = parseConfig({ version: 1, tenantId: id(1), clientId: id(2), serviceId: "cowork", budgets: [] });
  const perform = vi.fn(async () => undefined);
  const html = renderToStaticMarkup(<OnboardingCard record={value} admin={admin}
    api={new RequestApi("https://example.test", async () => "synthetic-token")} auth={new DemoAuth(config)}
    account={{ tenantId: id(1), localAccountId: id(3), homeAccountId: "home",
      environment: "login.microsoftonline.com", username: "owner@example.test" }}
    busy={false} defaults={{ standard: record.standard, enhanced: record.enhanced }}
    perform={perform} progress={() => undefined} onPublished={() => undefined} />);
  return { html, perform };
}

describe("compact BU request cards", () => {
  it.each([
    ["pending", "Awaiting approval"], ["provisioning", "Provisioning"],
    ["completed", "Completed"], ["rejected", "Rejected"],
  ] as const)("shows the %s status without expanding the record", (status, label) => {
    const { html, perform } = render({ ...record, status });
    expect(html).toContain(`class="badge request-status-badge request-status-${status}">${label}</span>`);
    expect(html.indexOf(`>${label}</span>`)).toBeLessThan(html.indexOf('class="request-card-details"'));
    expect(html).toContain("<h3>Legal BU</h3>");
    expect(html).toContain("LEGAL-01 / $20000 per month");
    expect(html).toContain('aria-expanded="false"');
    const controls = html.match(/aria-controls="([^"]+)"/)?.[1];
    expect(controls).toBeTruthy();
    expect(html).toContain(`id="${controls}" class="request-card-details" hidden=""`);
    expect(html).toContain('aria-label="Expand BU request Legal BU"');
    expect(perform).not.toHaveBeenCalled();
  });
  it("retains admin review controls in the hidden panel without exposing them to owners", () => {
    const admin = render(record, true).html;
    const owner = render(record).html;
    for (const text of ["Approve and provision", "Validation note / rejection reason", "validated this cost centre"]) {
      expect(admin).toContain(text);
      expect(owner).not.toContain(text);
    }
    expect(admin.indexOf('hidden=""')).toBeLessThan(admin.indexOf("Approve and provision"));
    expect(owner).toContain(record.justification);
  });
  it("keeps provisioning orange/incomplete and preserves the explicit admin recovery action", () => {
    const { html } = render({ ...record, status: "provisioning", resources: { standardGroup: id(5), enhancedGroup: id(6) } }, true);
    expect(html).toContain("request-status-provisioning");
    expect(html).not.toContain("request-status-completed");
    expect(html).toContain("Created resources: 2/4");
    expect(html).toContain("Resume provisioning");
    expect(html).toContain("Approved - provisioning not yet complete");
  });
  it("keeps recorded failures visible even when the details are collapsed", () => {
    const { html } = render({ ...record, status: "provisioning", lastError: "Synthetic provisioning failure" }, true);
    expect(html).toContain('hidden=""');
    expect(html).toContain('</div><p class="error" role="alert">Synthetic provisioning failure</p></article>');
  });
  it("preserves rejection notes and completed audit details without review controls", () => {
    const rejected = render({ ...record, status: "rejected", rejectionReason: "Cost centre needs finance review." }, true).html;
    expect(rejected).toContain("Rejection reason: Cost centre needs finance review.");
    expect(rejected).not.toContain("Approve and provision");
    const completed = render({ ...record, status: "completed", reviewNote: "Cost centre and cross-charge confirmed." }, true).html;
    expect(completed).toContain("Completed - owner workspace available");
    expect(completed).toContain("Finance review: Cost centre and cross-charge confirmed.");
    expect(completed).not.toContain("Resume provisioning");
  });
});
