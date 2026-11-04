import { createHash } from "node:crypto";
import { z } from "zod";
import type { Caller, Metadata } from "./contracts.js";
import type { ReportingSettings } from "./reporting-contracts.js";
import type { SyncJob, SyncSettings } from "./sync-contracts.js";
import { ApiError, uuid } from "./model.js";
import type { PermissionDestination, ReportingPermissions } from "./reporting-sharepoint.js";

export const reportingSecuritySchema = z.object({ policyId: uuid.optional(), jobId: uuid.optional(), leaseId: uuid.optional() }).strict()
  .refine((value) => Boolean(value.jobId) === Boolean(value.leaseId));
export interface SecurityState {
  version: 1; ready: boolean; metadataEtag: string | null; ownershipHash: string; configurationHash: string; checkedAt: string;
}
export interface ReportingSnapshot { metadata: Metadata | null; etag: string | null }
export interface ReportingStore {
  metadata(): Promise<ReportingSnapshot>;
  state(): Promise<SecurityState | null>;
  save(state: SecurityState): Promise<void>;
  locked<T>(operation: (assertLock: () => void, metadataLeaseId?: string) => Promise<T>): Promise<T>;
}
export interface ReportingConfiguration {
  enabled: boolean; destination: PermissionDestination | null; tenantId: string; clientId: string; certificateId: string; admins: string[];
}
export function reportingConfiguration(sync: SyncSettings, env: NodeJS.ProcessEnv = process.env): ReportingConfiguration {
  const admins = [...new Set((env.ONBOARDING_ADMIN_OBJECT_IDS ?? "").split(",").filter((id) => id.trim()).map((id) => uuid.parse(id.trim())))].sort();
  const enabled = env.REPORTING_SECURITY_ENABLED === "true";
  const certificateId = env.REPORTING_CERTIFICATE_ID?.trim() ?? "";
  const siteUrl = env.REPORTING_SITE_URL?.trim().replace(/\/$/, "") ?? "";
  const groupId = env.REPORTING_ADMIN_GROUP_ID ?? "";
  const numeric = (name: string, fallback: number, max: number) => {
    const raw = env[name] ?? String(fallback);
    if (!/^\d+$/.test(raw) || Number(raw) < 1 || Number(raw) > max) throw new Error(`Invalid ${name}.`);
    return Number(raw);
  };
  let destination: PermissionDestination | null = null;
  if (enabled && sync.enabled && certificateId && siteUrl && /^[1-9]\d*$/.test(groupId) && Number.isSafeInteger(Number(groupId)) && admins.length) {
    if (!/^https:\/\/[a-z0-9-]+\.vault\.azure\.net\/certificates\/[a-zA-Z0-9-]+\/[a-f0-9]{32}$/.test(certificateId)) {
      throw new ApiError(409, "reporting-not-configured", "REPORTING_CERTIFICATE_ID must be a versioned Azure Key Vault public certificate URI.");
    }
    let uri: URL;
    try { uri = new URL(siteUrl); }
    catch { throw new ApiError(409, "reporting-not-configured", "REPORTING_SITE_URL must identify the configured SharePoint reporting site."); }
    if (uri.protocol !== "https:" || uri.hostname !== sync.siteId.split(",")[0] || uri.port || uri.username ||
        uri.password || uri.search || uri.hash || !/^\/(?:sites|teams)\/[a-zA-Z0-9._~-]+$/.test(uri.pathname)) {
      throw new ApiError(409, "reporting-not-configured", "REPORTING_SITE_URL must identify the configured SharePoint reporting site.");
    }
    destination = { ...sync, siteUrl, adminGroupId: Number(groupId),
      maxScopes: numeric("REPORTING_MAX_SCOPES", 5000, 5000), scopeReserve: numeric("REPORTING_SCOPE_RESERVE", 250, 4999) };
  }
  return { enabled, destination, tenantId: uuid.parse(env.TENANT_ID), clientId: uuid.parse(env.API_AUDIENCE), certificateId, admins };
}
export function policyOwners(metadata: Metadata | null): Map<string, string[]> {
  const owners = new Map<string, Set<string>>();
  for (const budget of metadata?.budgets ?? []) for (const tier of Object.values(budget.tiers)) {
    const values = owners.get(tier.policyId) ?? new Set<string>();
    budget.owners.forEach((id) => values.add(id));
    owners.set(tier.policyId, values);
  }
  return new Map([...owners].sort(([a], [b]) => a.localeCompare(b)).map(([id, values]) => [id, [...values].sort()]));
}
const hash = (value: unknown) => createHash("sha256").update(JSON.stringify(value)).digest("hex");
export class ReportingService {
  constructor(private readonly store: ReportingStore, private readonly configuration: ReportingConfiguration,
    private readonly syncEnabled: boolean,
    private readonly permissions: (assertLock: () => void) => ReportingPermissions,
    private readonly currentJob: () => Promise<SyncJob | null>,
    private readonly clock: () => Date = () => new Date()) {}
  private requireAdmin(caller: Caller) {
    if (caller.tenantId !== this.configuration.tenantId || !this.configuration.admins.includes(caller.objectId)) {
      throw new ApiError(403, "reporting-admin-required", "Only a configured administrator can reconcile reporting permissions.");
    }
  }
  private configured() {
    if (!this.configuration.enabled || !this.configuration.destination || !this.configuration.certificateId) {
      throw new ApiError(409, "reporting-not-configured",
        "Reporting permission provisioning is not configured. Explicitly enable it, configure the administrator group, certificate and SharePoint site, and grant delegated consent.");
    }
  }
  private state(snapshot: ReportingSnapshot, ready: boolean): SecurityState {
    return { version: 1, ready, metadataEtag: snapshot.etag, ownershipHash: hash([...policyOwners(snapshot.metadata)]),
      configurationHash: hash(this.configuration), checkedAt: this.clock().toISOString() };
  }
  async settings(caller: Caller): Promise<ReportingSettings> {
    if (caller.tenantId !== this.configuration.tenantId) throw new ApiError(403, "forbidden", "Reporting tenant does not match.");
    const snapshot = await this.store.metadata();
    const isAdmin = this.configuration.admins.includes(caller.objectId);
    const policyIds = [...policyOwners(snapshot.metadata)].filter(([, owners]) => owners.includes(caller.objectId)).map(([id]) => id);
    if (!isAdmin && !policyIds.length) throw new ApiError(403, "reporting-owner-required", "Reporting is available only to administrators and current policy owners.");
    const state = await this.store.state();
    const expected = this.state(snapshot, true);
    const ready = Boolean(this.configuration.enabled && this.configuration.destination && snapshot.metadata && state?.ready &&
      state.metadataEtag === expected.metadataEtag && state.ownershipHash === expected.ownershipHash &&
      state.configurationHash === expected.configurationHash);
    // A non-ready response contains no destination for owners: toggling a flag is
    // not evidence that SharePoint ACLs have actually been read back.
    const destination = this.configuration.destination;
    const disclose = isAdmin || ready;
    return { enabled: Boolean(this.configuration.enabled && destination),
      siteId: disclose ? destination?.siteId ?? "" : "", policyListId: disclose ? destination?.policyListId ?? "" : "",
      balanceListId: disclose ? destination?.balanceListId ?? "" : "", isAdmin, policyIds: isAdmin ? [] : policyIds, securityReady: ready };
  }
  private async lease(caller: Caller, input: z.infer<typeof reportingSecuritySchema>) {
    if (!input.jobId) return;
    const job = await this.currentJob();
    if (!job || job.tenantId !== caller.tenantId || job.id !== input.jobId || job.executorId !== caller.objectId ||
        !job.lease || job.lease.id !== input.leaseId || Date.parse(job.lease.expiresAt) <= this.clock().getTime() ||
        job.state !== "running" || (input.policyId && !job.policies.some((policy) => policy.id === input.policyId))) {
      throw new ApiError(409, "reporting-sync-lease", "An active tenant sync job and its current executor lease are required.");
    }
  }
  async reconcile(caller: Caller, input: z.infer<typeof reportingSecuritySchema>) {
    this.requireAdmin(caller); this.configured();
    await this.lease(caller, input);
    return this.store.locked(async (assertLock) => {
      const snapshot = await this.store.metadata();
      await this.store.save(this.state(snapshot, false));
      const counts = await this.permissions(assertLock).reconcile(policyOwners(snapshot.metadata));
      assertLock(); await this.lease(caller, input);
      const after = await this.store.metadata();
      if (after.etag !== snapshot.etag || hash(after.metadata) !== hash(snapshot.metadata)) throw new ApiError(409, "reporting-metadata-conflict", "Ownership metadata changed during verification. Retry.");
      const verified = this.state(after, true);
      assertLock(); await this.store.save(verified);
      return { ready: true as const, checkedAt: verified.checkedAt, ...counts };
    });
  }
  async mutateMetadata(next: Metadata, etag: string | null, write: (leaseId?: string) => Promise<void>) {
    // Even if configuration was removed, a persisted security marker prevents a
    // formerly secured deployment from bypassing revocation by disabling a flag.
    const previousState = await this.store.state();
    if (!this.syncEnabled && !previousState && !this.configuration.enabled) return write();
    return this.store.locked(async (assertLock, metadataLeaseId) => {
      const before = await this.store.metadata();
      if (before.etag !== etag) throw new ApiError(409, "metadata-conflict", "Ownership metadata changed. Refresh and retry.");
      if (hash([...policyOwners(before.metadata)]) === hash([...policyOwners(next)])) {
        assertLock(); await write(metadataLeaseId);
        // Metadata ETag still changes, so require a fresh verification.
        await this.store.save(this.state({ metadata: next, etag: null }, false));
        return;
      }
      this.configured();
      await this.store.save(this.state(before, false));
      await this.permissions(assertLock).reconcile(new Map(), true);
      assertLock();
      // Never commit a new ownership authority until every old row grant has
      // been revoked and read back. A partial failure leaves old metadata intact.
      await write(metadataLeaseId);
    });
  }
}
