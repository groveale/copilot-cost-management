import { z } from "zod";
import type { Metadata, Page } from "./contracts.js";
import { metadataSchema, policySettingsSchema as settingsSchema, uuid } from "./model.js";
import { usdToCredits } from "./budget-math.js";

export const resourceKeys = ["standardGroup", "enhancedGroup", "standardPolicy", "enhancedPolicy"] as const;
export type ResourceKey = typeof resourceKeys[number];
const credit = z.number().int().nonnegative().max(Number.MAX_SAFE_INTEGER);
export const buSubmissionSchema = z.object({
  operationId: uuid, name: z.string().trim().min(2).max(80),
  costCenter: z.string().trim().min(1).max(80),
  monthlyBudgetUsd: z.string().refine((value) => { try { usdToCredits(value); return true; } catch { return false; } },
    "A positive USD amount with at most two decimal places is required."),
  justification: z.string().trim().min(20).max(2000),
}).strict();
export const buDecisionSchema = z.discriminatedUnion("status", [
  z.object({ status: z.literal("approved"), etag: z.string().min(1).max(256).refine((value) => value !== "*"),
    financeConfirmed: z.literal(true), reviewNote: z.string().trim().min(10).max(1000),
    standard: settingsSchema, enhanced: settingsSchema }).strict(),
  z.object({ status: z.literal("rejected"), etag: z.string().min(1).max(256).refine((value) => value !== "*"),
    reason: z.string().trim().min(10).max(1000) }).strict(),
]);
export const workUpdateSchema = z.object({
  action: z.enum(["claim", "release", "heartbeat", "intent", "clear-intent", "resource", "complete"]),
  etag: z.string().min(1).max(256).refine((value) => value !== "*"),
  leaseId: uuid.optional(), resource: z.enum(resourceKeys).optional(), resourceId: uuid.optional(),
  verified: z.literal(true).optional(),
  failure: z.string().trim().min(1).max(500).optional(),
}).strict();
const perUserEditSchema = z.object({
  tier: z.enum(["standard", "enhanced"]), credits: z.number().nonnegative().max(Number.MAX_SAFE_INTEGER).nullable(),
}).strict();
export const allocationSchema = z.object({
  operationId: uuid, standard: credit, enhanced: credit, perUserEdit: perUserEditSchema.optional(),
}).strict();
export const memberUpdateSchema = z.discriminatedUnion("action", [
  z.object({ action: z.literal("check"), userId: uuid, tiers: metadataSchema.shape.budgets.element.shape.tiers }).strict(),
  z.object({ action: z.literal("complete"), userId: uuid, tiers: metadataSchema.shape.budgets.element.shape.tiers, verified: z.literal(true) }).strict(),
]);
export type AllocationInput = z.infer<typeof allocationSchema>;
export type MemberUpdate = z.infer<typeof memberUpdateSchema>;
export const onboardingRecordSchema = z.object({
  version: z.literal(1), kind: z.enum(["onboarding", "allocation"]), id: z.string(),
  tenantId: uuid, requesterId: uuid, operationId: uuid, name: z.string(), costCenter: z.string(),
  monthlyBudgetUsd: z.string(), justification: z.string(), createdAt: z.iso.datetime(),
  displayName: z.string().optional(), contactHint: z.string().optional(),
  budgetId: z.string(), status: z.enum(["pending", "provisioning", "rejected", "completed"]),
  approvedCredits: credit, standard: settingsSchema.extend({ credits: credit }),
  enhanced: settingsSchema.extend({ credits: credit }),
  perUserEdit: perUserEditSchema.optional(),
  resources: z.object({ standardGroup: uuid.optional(), enhancedGroup: uuid.optional(),
    standardPolicy: uuid.optional(), enhancedPolicy: uuid.optional() }).strict(),
  intents: z.array(z.enum(resourceKeys)), decisionBy: uuid.optional(), decisionAt: z.iso.datetime().optional(),
  reviewNote: z.string().optional(), rejectionReason: z.string().optional(), completedAt: z.iso.datetime().optional(),
  financeConfirmed: z.boolean().optional(), lastError: z.string().optional(),
  lease: z.object({ id: uuid, actor: uuid, until: z.iso.datetime() }).strict().optional(), etag: z.string(),
}).strict();
export type OnboardingRecord = z.infer<typeof onboardingRecordSchema>;
export type BuSubmission = z.infer<typeof buSubmissionSchema>;
export type BuDecision = z.infer<typeof buDecisionSchema>;
export type WorkUpdate = z.infer<typeof workUpdateSchema>;
export type TierSettings = z.infer<typeof settingsSchema>;
export interface MetadataSnapshot { metadata: Metadata | null; etag: string | null }
export interface OnboardingStore {
  get(id: string): Promise<OnboardingRecord | null>;
  create(record: OnboardingRecord): Promise<boolean>;
  replace(record: OnboardingRecord, etag: string): Promise<boolean>;
  list(requesterId?: string, cursor?: string): Promise<Page<OnboardingRecord>>;
  metadata(): Promise<MetadataSnapshot>;
  saveMetadata(metadata: Metadata, etag: string | null): Promise<void>;
  routing(objectId: string): Promise<string | null>;
  seedRouting(objectId: string, budgetId: string): Promise<void>;
}
