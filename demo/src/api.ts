import type { Context, Decision, Page, RequestRecord, Submission } from "../../api/src/contracts";
export type { Context, Decision, Page, RequestRecord, Submission };

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
}
