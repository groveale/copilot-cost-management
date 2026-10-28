import { randomUUID } from "node:crypto";
import { z } from "zod";
import type { Caller } from "./contracts.js";
import { ApiError, uuid } from "./model.js";
import type { SyncJob, SyncSettings, SyncState, SyncStore, SyncUpdate } from "./sync-contracts.js";
export type { StoredSync, SyncJob, SyncStore } from "./sync-contracts.js";
const concurrency = z.number().int().min(1).max(8);
const policySchema = z.object({
  id: uuid, displayName: z.string().trim().min(1).max(255), type: z.string().max(200),
}).strict();
export const storedSyncSchema: z.ZodType<SyncJob> = z.object({
  version: z.literal(1), revision: z.number().int().positive(), id: uuid, tenantId: uuid, initiatorId: uuid, executorId: uuid.optional(),
  actors: z.array(uuid).max(10000), policyId: uuid.optional(), concurrency,
  state: z.enum(["queued", "running", "paused", "access-required", "completed", "completed-with-exclusions",
    "partially-failed", "failed", "cancelled"]),
  phase: z.string().max(255), createdAt: z.iso.datetime(), updatedAt: z.iso.datetime(), finishedAt: z.iso.datetime().optional(),
  lease: z.object({ id: uuid, expiresAt: z.iso.datetime() }).strict().optional(),
  inventoryComplete: z.boolean(), mirrorPending: z.boolean(),
  policies: z.array(policySchema.extend({
    state: z.enum(["pending", "running", "completed", "excluded", "failed"]),
    users: z.number().int().nonnegative().optional(), groups: z.number().int().nonnegative().optional(),
    exclusions: z.number().int().nonnegative().optional(), error: z.string().max(160).optional(),
  }).strict()).max(10000),
}).strict();
export const syncMaintenanceSchema = z.object({ jobId: uuid, leaseId: uuid }).strict();
export const syncCreateSchema = z.object({ operationId: uuid, policyId: uuid.optional(), concurrency }).strict();
export const syncUpdateSchema = z.discriminatedUnion("action", [
  z.object({ action: z.literal("claim") }).strict(),
  z.object({ action: z.literal("claim-maintenance") }).strict(),
  z.object({ action: z.literal("release-maintenance"), leaseId: uuid }).strict(),
  z.object({ action: z.literal("cancel") }).strict(),
  z.object({ action: z.literal("retry") }).strict(),
  z.object({ action: z.literal("heartbeat"), leaseId: uuid }).strict(),
  z.object({ action: z.literal("inventory"), leaseId: uuid, policies: z.array(policySchema).max(10000) }).strict(),
  z.object({ action: z.literal("policy"), leaseId: uuid, policy: policySchema.extend({
    state: z.enum(["running", "completed", "excluded", "failed"]),
    users: z.number().int().nonnegative().optional(), groups: z.number().int().nonnegative().optional(),
    exclusions: z.number().int().nonnegative().optional(),
    error: z.string().regex(/^[A-Za-z0-9_.: -]{1,160}$/).optional(),
  }).strict() }).strict(),
  z.object({ action: z.literal("pause"), leaseId: uuid, accessRequired: z.boolean() }).strict(),
  z.object({ action: z.literal("finish"), leaseId: uuid }).strict(),
  z.object({ action: z.literal("mirrored"), revision: z.number().int().positive() }).strict(),
]);
export const syncTerminal = (state: SyncState) =>
  ["completed", "completed-with-exclusions", "partially-failed", "failed", "cancelled"].includes(state);

export function syncSettings(env: NodeJS.ProcessEnv = process.env): SyncSettings {
  const keys = ["SYNC_SITE_ID", "SYNC_POLICY_LIST_ID", "SYNC_BALANCE_LIST_ID", "SYNC_JOB_LIST_ID"] as const;
  const values = keys.map((key) => env[key]?.trim() ?? "");
  const numeric = (key: string, fallback: number, max: number) => {
    const raw = env[key] ?? String(fallback);
    if (!/^\d+$/.test(raw) || Number(raw) < 1 || Number(raw) > max) throw new Error(`Invalid ${key}.`);
    return Number(raw);
  };
  const maxConcurrency = numeric("SYNC_MAX_CONCURRENCY", 8, 8);
  const defaultConcurrency = numeric("SYNC_DEFAULT_CONCURRENCY", 3, maxConcurrency);
  if (values.some(Boolean) && !values.every(Boolean)) throw new Error("Complete all four SharePoint sync destination settings.");
  if (values.every(Boolean)) {
    if (!/^[a-z0-9.-]+\.sharepoint\.com,[a-f0-9-]{36},[a-f0-9-]{36}$/i.test(values[0])) throw new Error("Invalid SYNC_SITE_ID.");
    values.slice(1).forEach((value) => uuid.parse(value));
  }
  return { enabled: values.every(Boolean), siteId: values[0], policyListId: values[1], balanceListId: values[2],
    jobListId: values[3], defaultConcurrency, maxConcurrency, retentionDays: numeric("SYNC_JOB_RETENTION_DAYS", 30, 3650) };
}
export class SyncService {
  constructor(private readonly store: SyncStore, private readonly settings: SyncSettings,
    private readonly requireAdmin: (caller: Caller) => void, private readonly clock: () => Date = () => new Date()) {}
  configuration(caller: Caller) { this.requireAdmin(caller); return this.settings; }
  async current(caller: Caller) { this.requireAdmin(caller); return (await this.store.current())?.job ?? null; }
  async maintain(caller: Caller, input: z.infer<typeof syncMaintenanceSchema>) {
    this.requireAdmin(caller);
    const active = await this.store.current();
    if (!active || active.job.id !== input.jobId || active.job.lease?.id !== input.leaseId ||
        active.job.executorId !== caller.objectId || active.job.phase !== "Maintenance" ||
        Date.parse(active.job.lease.expiresAt) <= this.clock().getTime()) {
      throw new ApiError(409, "sync-active", "A live maintenance lease is required.");
    }
    return { removed: await this.store.prune(this.settings.retentionDays) };
  }
  async list(caller: Caller, cursor?: string) { this.requireAdmin(caller); return this.store.list(cursor); }
  async get(caller: Caller, id: string) {
    this.requireAdmin(caller);
    const job = await this.store.get(id);
    if (!job || job.tenantId !== caller.tenantId) throw new ApiError(404, "sync-not-found", "Sync job not found.");
    return job;
  }
  async create(caller: Caller, input: z.infer<typeof syncCreateSchema>) {
    this.requireAdmin(caller);
    if (!this.settings.enabled) throw new ApiError(409, "sync-setup-required", "Provision and configure the SharePoint sync destination first.");
    if (input.concurrency > this.settings.maxConcurrency) throw new ApiError(400, "sync-concurrency", "Concurrency exceeds the configured maximum.");
    const existing = await this.store.get(input.operationId);
    if (existing) {
      if (existing.tenantId !== caller.tenantId || existing.initiatorId !== caller.objectId ||
          existing.policyId !== input.policyId || existing.concurrency !== input.concurrency) {
        throw new ApiError(409, "sync-operation-conflict", "This operation ID belongs to another sync request.");
      }
      return existing;
    }
    const current = await this.store.current();
    if (current && (!syncTerminal(current.job.state) || (current.job.lease && Date.parse(current.job.lease.expiresAt) > this.clock().getTime()))) {
      throw new ApiError(409, "sync-active", "Another sync or maintenance task is active. Resume or cancel it first.");
    }
    const now = this.clock().toISOString();
    const job: SyncJob = { version: 1, revision: 1, id: input.operationId, tenantId: caller.tenantId, initiatorId: caller.objectId,
      actors: [caller.objectId], policyId: input.policyId, concurrency: input.concurrency, state: "queued",
      phase: "Awaiting access validation", createdAt: now, updatedAt: now, inventoryComplete: false, policies: [], mirrorPending: true };
    if (!await this.store.save(job, current?.etag)) throw new ApiError(409, "sync-conflict", "Sync execution changed. Refresh before continuing.");
    return job;
  }
  async update(caller: Caller, id: string, input: SyncUpdate) {
    this.requireAdmin(caller);
    if (input.action === "mirrored") {
      const saved = await this.store.acknowledge(id, input.revision);
      if (!saved || saved.tenantId !== caller.tenantId) throw new ApiError(404, "sync-not-found", "Sync job not found.");
      return saved;
    }
    for (let attempt = 0; attempt < 5; attempt++) {
      const current = await this.store.current();
      if (!current || current.job.id !== id || current.job.tenantId !== caller.tenantId) {
        throw new ApiError(409, "sync-not-current", "This is no longer the current execution job.");
      }
      const job = structuredClone(current.job);
      const now = this.clock();
      if (input.action === "claim" || input.action === "claim-maintenance") {
        const maintenance = input.action === "claim-maintenance";
        if (maintenance ? !syncTerminal(job.state) : syncTerminal(job.state)) {
          throw new ApiError(409, "sync-finished", maintenance ? "Finish or cancel collection before maintenance." : "Retry failed policies or start a new sync.");
        }
        if (job.lease && Date.parse(job.lease.expiresAt) > now.getTime()) {
          throw new ApiError(409, "sync-leased", "Another browser holds the live execution lease. Pause it or wait for lease expiry.");
        }
        job.executorId = caller.objectId;
        if (!job.actors.includes(caller.objectId)) job.actors.push(caller.objectId);
        job.lease = { id: randomUUID(), expiresAt: new Date(now.getTime() + 90_000).toISOString() };
        if (maintenance) job.phase = "Maintenance";
        else {
          job.state = "running"; job.phase = "Validating destination writes";
          job.policies = job.policies.map((policy) => policy.state === "running" ? { ...policy, state: "pending" } : policy);
        }
      } else if (input.action === "cancel") {
        if (syncTerminal(job.state)) throw new ApiError(409, "sync-finished", "This sync has already finished.");
        job.state = "cancelled"; job.phase = "Cancelled by administrator"; job.lease = undefined;
        job.finishedAt = now.toISOString();
      } else if (input.action === "retry") {
        if (job.lease && Date.parse(job.lease.expiresAt) > now.getTime()) throw new ApiError(409, "sync-leased", "Maintenance is active.");
        if (!["failed", "partially-failed"].includes(job.state)) throw new ApiError(409, "sync-not-failed", "Only failed policy work can be retried.");
        job.policies = job.policies.map((policy) => policy.state === "failed" ? { id: policy.id, displayName: policy.displayName, type: policy.type, state: "pending" } : policy);
        job.state = "paused"; job.phase = "Failed work ready to resume"; job.finishedAt = undefined; job.lease = undefined;
      } else {
        const maintenance = job.phase === "Maintenance" && ["heartbeat", "release-maintenance"].includes(input.action);
        if ((job.state !== "running" && !maintenance) || !job.lease || job.lease.id !== input.leaseId ||
            job.executorId !== caller.objectId || Date.parse(job.lease.expiresAt) <= now.getTime()) {
          throw new ApiError(409, "sync-lease-lost", "The execution lease expired or changed. Stop work and refresh.");
        }
        if (input.action === "heartbeat") job.lease.expiresAt = new Date(now.getTime() + 90_000).toISOString();
        else if (input.action === "release-maintenance") {
          if (!maintenance) throw new ApiError(409, "sync-not-maintenance", "This is not a maintenance lease.");
          job.lease = undefined; job.phase = "Finished; maintenance released";
        }
        else if (input.action === "inventory") {
          if (job.inventoryComplete) throw new ApiError(409, "sync-inventory-exists", "This job already has an inventory.");
          const ids = new Set(input.policies.map((policy) => policy.id));
          if (ids.size !== input.policies.length || (job.policyId && (ids.size !== 1 || !ids.has(job.policyId)))) {
            throw new ApiError(400, "sync-inventory", "Policy inventory does not match the requested scope.");
          }
          job.policies = input.policies.map((policy) => ({ ...policy, state: "pending" }));
          job.inventoryComplete = true; job.phase = "Collecting policy snapshots";
        } else if (input.action === "policy") {
          const index = job.policies.findIndex((policy) => policy.id === input.policy.id);
          if (index < 0 || !job.inventoryComplete) throw new ApiError(400, "sync-policy", "Policy is not in this job's inventory.");
          if (["completed", "excluded"].includes(job.policies[index].state)) throw new ApiError(409, "sync-policy-finished", "Policy has already been published.");
          job.policies[index] = input.policy;
          job.phase = `Policy ${input.policy.state}`;
        } else if (input.action === "pause") {
          job.state = input.accessRequired ? "access-required" : "paused";
          job.phase = input.accessRequired ? "Refresh access and validate before resuming" : "Paused";
          job.lease = undefined;
        } else {
          if (!job.inventoryComplete || job.policies.some((policy) => ["pending", "running"].includes(policy.state))) {
            throw new ApiError(409, "sync-incomplete", "Inventory and all policy outcomes must be recorded before finishing.");
          }
          const failures = job.policies.filter((policy) => policy.state === "failed").length;
          job.state = failures ? (failures === job.policies.length ? "failed" : "partially-failed") :
            job.policies.some((policy) => policy.state === "excluded" || policy.exclusions) ? "completed-with-exclusions" : "completed";
          job.phase = "Finished"; job.finishedAt = now.toISOString(); job.lease = undefined;
        }
      }
      job.updatedAt = now.toISOString(); job.revision++; job.mirrorPending = true;
      if (!await this.store.save(job, current.etag)) continue;
      return job;
    }
    throw new ApiError(409, "sync-conflict", "Sync progress changed concurrently. Refresh and retry.");
  }
}
