import type { Context, Decision, FulfilmentUpdate, Page, RequestRecord, Submission } from "../../api/src/contracts";
import type { Budget, Metadata } from "../../api/src/contracts";
import type { AllocationInput, BuDecision, BuSubmission, MemberUpdate, OnboardingRecord, WorkUpdate } from "../../api/src/onboarding";
export type { Context, Decision, Page, RequestRecord, Submission };
export interface OnboardingCapabilities { isAdmin: boolean; bootstrapEnabled: boolean; objectId: string; displayName?: string; contactHint?: string }

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
      throw new Error(`${error.message ?? `API request failed (HTTP ${response.status}).`}${error.correlationId ? ` Reference: ${error.correlationId}` : ""}`);
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
}
