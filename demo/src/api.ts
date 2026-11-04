import type { Context, Decision, FulfilmentUpdate, Page, RequestRecord, Submission } from "../../api/src/contracts";
import type { Budget, Metadata } from "../../api/src/contracts";
import type { AllocationInput, BuDecision, BuSubmission, MemberUpdate, OnboardingRecord, WorkUpdate } from "../../api/src/onboarding";
import type { SyncJob, SyncSettings, SyncUpdate } from "../../api/src/sync-contracts";
import type { ReportingSettings } from "../../api/src/reporting-contracts";
import { claimsChallenge } from "./graph";
export type { Context, Decision, Page, RequestRecord, Submission };
export interface OnboardingCapabilities { isAdmin: boolean; bootstrapEnabled: boolean; objectId: string; displayName?: string; contactHint?: string }
export class ApiRequestError extends Error {
  constructor(message: string, readonly status: number, readonly code: string, readonly claims?: string) { super(message); }
}

export class RequestApi {
  constructor(
    private readonly baseUrl: string,
    private readonly token: () => Promise<string>,
    private readonly transport: typeof fetch = fetch.bind(globalThis),
  ) {}
  private async call<T>(path: string, signal: AbortSignal, body?: unknown): Promise<T> {
    const token = await this.token();
    const response = await this.transport(`${this.baseUrl}/api/v1/${path}`, {
      method: body === undefined ? "GET" : "POST", signal,
      headers: { Authorization: `Bearer ${token}`, ...(body === undefined ? {} : { "Content-Type": "application/json" }) },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
    });
    if (!response.ok) {
      const error = await response.json();
      throw new ApiRequestError(`${error.message ?? `API request failed (HTTP ${response.status}).`}${error.correlationId ? ` Reference: ${error.correlationId}` : ""}`,
        response.status, typeof error.code === "string" ? error.code : "api-failed",
        claimsChallenge(response.headers.get("www-authenticate")));
    }
    return response.json();
  }
  context(signal: AbortSignal): Promise<Context> { return this.call("context", signal); }
  history(signal: AbortSignal, cursor?: string): Promise<Page<RequestRecord>> {
    return this.call(`requests${cursor ? `?cursor=${encodeURIComponent(cursor)}` : ""}`, signal);
  }
  inbox(signal: AbortSignal, cursor?: string): Promise<Page<RequestRecord>> {
    return this.call(`inbox${cursor ? `?cursor=${encodeURIComponent(cursor)}` : ""}`, signal);
  }
  submit(input: Submission, signal: AbortSignal): Promise<RequestRecord> { return this.call("requests", signal, input); }
  decide(id: string, input: Decision, signal: AbortSignal): Promise<RequestRecord> {
    return this.call(`requests/${id}/decision`, signal, input);
  }
  fulfilment(id: string, input: FulfilmentUpdate, signal: AbortSignal): Promise<RequestRecord> {
    return this.call(`requests/${id}/fulfilment`, signal, input);
  }
  onboardingCapabilities(signal: AbortSignal): Promise<OnboardingCapabilities> { return this.call("onboarding/capabilities", signal); }
  onboardingList(signal: AbortSignal, inbox = false, cursor?: string): Promise<Page<OnboardingRecord>> {
    return this.call(`onboarding${inbox ? "/inbox" : ""}${cursor ? `?cursor=${encodeURIComponent(cursor)}` : ""}`, signal);
  }
  onboardingSubmit(input: BuSubmission, signal: AbortSignal): Promise<OnboardingRecord> { return this.call("onboarding", signal, input); }
  onboardingGet(id: string, signal: AbortSignal): Promise<OnboardingRecord> { return this.call(`onboarding/${encodeURIComponent(id)}`, signal); }
  onboardingDecide(id: string, input: BuDecision, signal: AbortSignal): Promise<OnboardingRecord> {
    return this.call(`onboarding/${id}/decision`, signal, input);
  }
  onboardingWork(id: string, input: WorkUpdate, signal: AbortSignal): Promise<OnboardingRecord> {
    return this.call(`onboarding/${id}/work`, signal, input);
  }
  onboardingSetup(signal: AbortSignal): Promise<Metadata | null> { return this.call("onboarding/setup", signal); }
  bootstrap(input: unknown, signal: AbortSignal): Promise<{ budgets: number; mappings: number }> { return this.call("onboarding/bootstrap", signal, input); }
  member(budgetId: string, input: MemberUpdate, signal: AbortSignal): Promise<{ budget: Budget }> {
    return this.call(`budgets/${encodeURIComponent(budgetId)}/members`, signal, input);
  }
  allocation(budgetId: string, signal: AbortSignal, input?: AllocationInput): Promise<OnboardingRecord | null> {
    return this.call(`budgets/${encodeURIComponent(budgetId)}/allocation`, signal, input);
  }
  syncSettings(signal: AbortSignal): Promise<SyncSettings> { return this.call("admin/sync/settings", signal); }
  reportingSettings(signal: AbortSignal): Promise<ReportingSettings> { return this.call("reporting/settings", signal); }
  reportingReconcile(signal: AbortSignal, input: { policyId?: string; jobId?: string; leaseId?: string } = {}):
    Promise<{ ready: true; checkedAt: string; policies: number; rows: number }> {
    return this.call("admin/reporting/security", signal, input);
  }
  syncCurrent(signal: AbortSignal): Promise<SyncJob | null> { return this.call("admin/sync/current", signal); }
  syncMaintain(jobId: string, leaseId: string, signal: AbortSignal): Promise<{ removed: number }> {
    return this.call("admin/sync/maintenance", signal, { jobId, leaseId });
  }
  syncJobs(signal: AbortSignal, cursor?: string): Promise<Page<SyncJob>> {
    return this.call(`admin/sync/jobs${cursor ? `?cursor=${encodeURIComponent(cursor)}` : ""}`, signal);
  }
  syncJob(id: string, signal: AbortSignal): Promise<SyncJob> { return this.call(`admin/sync/jobs/${id}`, signal); }
  syncCreate(input: { operationId: string; policyId?: string; concurrency: number }, signal: AbortSignal): Promise<SyncJob> {
    return this.call("admin/sync/jobs", signal, input);
  }
  syncUpdate(id: string, input: SyncUpdate, signal: AbortSignal): Promise<SyncJob> {
    return this.call(`admin/sync/jobs/${id}`, signal, input);
  }
}
