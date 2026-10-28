import type { Page } from "./contracts.js";

export interface SyncSettings {
  enabled: boolean;
  siteId: string;
  policyListId: string;
  balanceListId: string;
  jobListId: string;
  defaultConcurrency: number;
  maxConcurrency: number;
  retentionDays: number;
}
export type SyncState = "queued" | "running" | "paused" | "access-required" | "completed" |
  "completed-with-exclusions" | "partially-failed" | "failed" | "cancelled";
export interface SyncPolicy {
  id: string;
  displayName: string;
  type: string;
  state: "pending" | "running" | "completed" | "excluded" | "failed";
  users?: number;
  groups?: number;
  exclusions?: number;
  error?: string;
}
export type SyncPolicyUpdate = Omit<SyncPolicy, "state"> & { state: Exclude<SyncPolicy["state"], "pending"> };
export interface SyncJob {
  version: 1;
  revision: number;
  id: string;
  tenantId: string;
  initiatorId: string;
  executorId?: string;
  actors: string[];
  policyId?: string;
  concurrency: number;
  state: SyncState;
  phase: string;
  createdAt: string;
  updatedAt: string;
  finishedAt?: string;
  lease?: { id: string; expiresAt: string };
  inventoryComplete: boolean;
  policies: SyncPolicy[];
  mirrorPending: boolean;
}
export type SyncUpdate =
  | { action: "claim" }
  | { action: "cancel" }
  | { action: "retry" }
  | { action: "claim-maintenance" }
  | { action: "heartbeat"; leaseId: string }
  | { action: "finish"; leaseId: string }
  | { action: "release-maintenance"; leaseId: string }
  | { action: "inventory"; leaseId: string; policies: Pick<SyncPolicy, "id" | "displayName" | "type">[] }
  | { action: "policy"; leaseId: string; policy: SyncPolicyUpdate }
  | { action: "pause"; leaseId: string; accessRequired: boolean }
  | { action: "mirrored"; revision: number };
export interface StoredSync { job: SyncJob; etag: string }
export interface SyncStore {
  current(): Promise<StoredSync | null>;
  get(id: string): Promise<SyncJob | null>;
  save(job: SyncJob, etag?: string): Promise<boolean>;
  acknowledge(id: string, revision: number): Promise<SyncJob | null>;
  prune(retentionDays: number): Promise<number>;
  list(cursor?: string): Promise<Page<SyncJob>>;
}
