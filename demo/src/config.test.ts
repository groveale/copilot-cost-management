import { describe, expect, it } from "vitest";
import { parseConfig } from "./config";

const id = (n: number) => `10000000-0000-0000-0000-${String(n).padStart(12, "0")}`;
const budget = () => ({
  id: "finance",
  displayName: "Finance",
  businessArea: "Finance",
  costCenter: "0001",
  owners: ["Demo owner"],
  tiers: {
    standard: { policyId: id(3), groupId: id(4) },
    enhanced: { policyId: id(5), groupId: id(6) },
  },
});
const config = () => ({
  version: 1, tenantId: id(1), clientId: id(2), serviceId: "cowork", budgets: [budget()],
});

describe("static JSON configuration", () => {
  it("preserves text cost centres and optional money metadata", () => {
    const result = parseConfig(config());
    expect(result.budgets[0].costCenter).toBe("0001");
    expect(result.budgets[0].plannedUsd).toBeUndefined();
  });
  it("allows an empty setup without inventing budget resources", () => {
    expect(parseConfig({ ...config(), budgets: [] }).budgets).toEqual([]);
  });
  it("rejects malformed and zero identifiers", () => {
    expect(() => parseConfig({ ...config(), tenantId: "not-a-guid" })).toThrow("GUID");
    expect(() => parseConfig({ ...config(), clientId: "00000000-0000-0000-0000-000000000000" })).toThrow("non-zero");
  });
  it("rejects duplicate policies between tiers", () => {
    const input = config();
    input.budgets[0].tiers.enhanced.policyId = input.budgets[0].tiers.standard.policyId;
    expect(() => parseConfig(input)).toThrow("mapped more than once");
  });
  it("rejects duplicate groups between budgets", () => {
    const first = budget();
    const second = { ...budget(), id: "second", tiers: {
      standard: { policyId: id(7), groupId: id(4) },
      enhanced: { policyId: id(8), groupId: id(9) },
    } };
    expect(() => parseConfig({ ...config(), budgets: [first, second] })).toThrow("Group");
  });
  it("requires exactly two named tiers", () => {
    const first = budget();
    expect(() => parseConfig({ ...config(), budgets: [{ ...first, tiers: { ...first.tiers, high: {} } }] })).toThrow("exactly");
  });
  it("rejects numeric cost centres and invalid descriptive USD", () => {
    expect(() => parseConfig({ ...config(), budgets: [{ ...budget(), costCenter: 1 }] })).toThrow("costCenter");
    expect(() => parseConfig({ ...config(), budgets: [{ ...budget(), plannedUsd: -1 }] })).toThrow("plannedUsd");
  });
  it("rejects duplicate budget keys", () => {
    expect(() => parseConfig({ ...config(), budgets: [budget(), budget()] })).toThrow("Duplicate budget");
  });
});
