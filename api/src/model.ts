import { z } from "zod";
import type { Metadata, RequestRecord } from "./contracts.js";

export const uuid = z.uuid().transform((value) => value.toLowerCase()).refine(
  (value) => value !== "00000000-0000-0000-0000-000000000000", "A non-zero UUID is required.",
);
const tier = z.object({ policyId: uuid, groupId: uuid }).strict();
export const metadataSchema = z.object({
  version: z.literal(1),
  tenantId: uuid,
  serviceId: z.string().trim().min(1),
  budgets: z.array(z.object({
    id: z.string().regex(/^[a-z0-9-]{1,64}$/),
    displayName: z.string().trim().min(1),
    businessArea: z.string().trim().min(1),
    costCenter: z.string().trim().min(1),
    owners: z.array(uuid).min(1),
    plannedUsd: z.number().nonnegative().optional(),
    approvedCreditCeiling: z.number().int().positive().max(Number.MAX_SAFE_INTEGER).optional(),
    periodLabel: z.string().trim().min(1).optional(),
    tiers: z.object({ standard: tier, enhanced: tier }).strict(),
  }).strict()),
}).strict().superRefine((data, ctx) => {
  for (const values of [
    data.budgets.map((b) => b.id),
    data.budgets.flatMap((b) => Object.values(b.tiers).map((t) => t.policyId)),
    data.budgets.flatMap((b) => Object.values(b.tiers).map((t) => t.groupId)),
  ]) {
    if (new Set(values).size !== values.length) ctx.addIssue({ code: "custom", message: "Duplicate budget, policy or group mapping." });
  }
});
export const submissionSchema = z.object({
  operationId: uuid,
  justification: z.string().trim().min(20).max(2000),
  grantKind: z.enum(["permanent", "billingPeriod"]),
}).strict();
export const decisionSchema = z.object({
  status: z.enum(["approved", "rejected"]),
  reason: z.string().trim().min(10).max(1000).optional(),
  etag: z.string().min(1).max(256).refine((value) => value !== "*", "A specific entity ETag is required."),
}).strict().refine((d) => d.status !== "rejected" || !!d.reason, "Rejection requires a reason.");
export const seedSchema = z.object({
  metadata: metadataSchema,
  requesterIndex: z.array(z.object({ objectId: uuid, budgetId: z.string() }).strict()),
}).strict().superRefine((seed, ctx) => {
  const seen = new Set<string>();
  for (const entry of seed.requesterIndex) {
    if (seen.has(entry.objectId) || !seed.metadata.budgets.some((b) => b.id === entry.budgetId)) {
      ctx.addIssue({ code: "custom", message: "Duplicate requester or unknown budget." });
    }
    seen.add(entry.objectId);
  }
});
export class ApiError extends Error {
  constructor(public readonly status: number, public readonly code: string, message: string) { super(message); }
}
export function periodEnd(now: Date): string {
  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + 1, 1)).toISOString();
}
export function validateStoredRequest(record: RequestRecord): RequestRecord {
  if (record.version !== 1) throw new Error("Unsupported request schema version.");
  return record;
}
export function parseMetadata(input: unknown, tenantId: string): Metadata {
  const parsed = metadataSchema.parse(input);
  if (parsed.tenantId !== tenantId) throw new Error("Metadata tenant does not match API tenant.");
  return parsed;
}
