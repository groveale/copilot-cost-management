import { afterEach, describe, expect, it, vi } from "vitest";
import { GraphClient, GraphError, retryAfterMs } from "./graph";

const signal = () => new AbortController().signal;
const json = (body: unknown, status = 200, headers = {}) =>
  new Response(JSON.stringify(body), { status, headers });

afterEach(() => {
  vi.useRealTimers();
  vi.unstubAllGlobals();
});

describe("direct Graph requests", () => {
  it("keeps the browser fetch receiver instead of calling it as a GraphClient method", async () => {
    const browserFetch = vi.fn<typeof fetch>(function (this: unknown) {
      expect(this).toBe(globalThis);
      return Promise.resolve(json({ value: [] }));
    });
    vi.stubGlobal("fetch", browserFetch);
    expect(await new GraphClient(async () => "t").get("/beta/data", signal())).toEqual({ value: [] });
    expect(browserFetch).toHaveBeenCalledTimes(1);
  });
  it("follows all continuation pages unchanged", async () => {
    const next = "https://graph.microsoft.com/beta/copilot/costManagement/spendingPolicies?$skiptoken=exact";
    const fetcher = vi.fn<typeof fetch>()
      .mockResolvedValueOnce(json({ value: [1], "@odata.nextLink": next }))
      .mockResolvedValueOnce(json({ value: [2] }));
    const client = new GraphClient(async () => "test-token", fetcher);
    expect(await client.list("/beta/copilot/costManagement/spendingPolicies", signal())).toEqual([1, 2]);
    expect(fetcher.mock.calls[1][0]).toBe(next);
    expect(fetcher.mock.calls[0][1]?.redirect).toBe("error");
  });
  it("never sends a token to an untrusted continuation host", async () => {
    const fetcher = vi.fn<typeof fetch>().mockResolvedValueOnce(json({ value: [], "@odata.nextLink": "https://example.com/beta/data" }));
    const token = vi.fn(async () => "test-token");
    const client = new GraphClient(token, fetcher);
    await expect(client.list("/beta/copilot/costManagement/spendingPolicies", signal())).rejects.toThrow("outside");
    expect(fetcher).toHaveBeenCalledTimes(1);
    expect(token).toHaveBeenCalledTimes(1);
  });
  it("rejects repeated pagination and malformed collection shapes", async () => {
    const fetcher = vi.fn<typeof fetch>().mockResolvedValue(json({ value: [], "@odata.nextLink": "/beta/data" }));
    await expect(new GraphClient(async () => "t", fetcher).list("/beta/data", signal())).rejects.toThrow("repeated");
    fetcher.mockResolvedValue(json({ serviceId: "cowork" }));
    await expect(new GraphClient(async () => "t", fetcher).list("/beta/data", signal())).rejects.toThrow("value array");
  });
  it("serializes overlapping requests and permits a later request after failure", async () => {
    let release: ((response: Response) => void) | undefined;
    const fetcher = vi.fn<typeof fetch>()
      .mockImplementationOnce(() => new Promise((resolve) => { release = resolve; }))
      .mockResolvedValueOnce(json({ ok: true }));
    const client = new GraphClient(async () => "t", fetcher);
    const first = client.get("/beta/first", signal());
    const second = client.get("/beta/second", signal());
    await vi.waitFor(() => expect(fetcher).toHaveBeenCalledTimes(1));
    release?.(json({ error: { code: "Denied", message: "Consent required" } }, 403));
    await expect(first).rejects.toBeInstanceOf(GraphError);
    await expect(second).resolves.toEqual({ ok: true });
    expect(fetcher).toHaveBeenCalledTimes(2);
  });
  it("honors Retry-After and does not return success after retry exhaustion", async () => {
    vi.useFakeTimers();
    const fetcher = vi.fn<typeof fetch>().mockResolvedValue(json({ error: { code: "tooManyRequests", message: "Throttled" } }, 429, { "Retry-After": "2" }));
    const result = new GraphClient(async () => "t", fetcher).get("/beta/data", signal());
    const rejection = expect(result).rejects.toThrow("tooManyRequests");
    await vi.advanceTimersByTimeAsync(5999);
    expect(fetcher).toHaveBeenCalledTimes(3);
    await vi.advanceTimersByTimeAsync(1);
    await rejection;
    expect(fetcher).toHaveBeenCalledTimes(4);
  });
  it("reports long throttling delays instead of ignoring Retry-After", async () => {
    const fetcher = vi.fn<typeof fetch>().mockResolvedValue(json({}, 429, { "Retry-After": "120" }));
    await expect(new GraphClient(async () => "t", fetcher).get("/beta/data", signal())).rejects.toThrow("120 seconds");
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
  it("reports Graph code and request ID, and malformed successful JSON", async () => {
    const fetcher = vi.fn<typeof fetch>()
      .mockResolvedValueOnce(json({ error: { code: "Forbidden", message: "Role required" } }, 403, { "request-id": "request-123" }))
      .mockResolvedValueOnce(new Response("not JSON", { status: 200 }));
    const client = new GraphClient(async () => "t", fetcher);
    await expect(client.get("/beta/data", signal())).rejects.toThrow("request-123");
    await expect(client.get("/beta/data", signal())).rejects.toThrow("not valid JSON");
  });
  it("cancels queued work without acquiring a token", async () => {
    const controller = new AbortController();
    controller.abort();
    const token = vi.fn(async () => "t");
    await expect(new GraphClient(token).get("/beta/data", controller.signal)).rejects.toThrow();
    expect(token).not.toHaveBeenCalled();
  });
  it("deletes only an explicit group membership $ref and never the directory object", async () => {
    const fetcher = vi.fn<typeof fetch>(async () => new Response(null, { status: 204 }));
    const token = vi.fn(async () => "t");
    const client = new GraphClient(token, fetcher);
    const path = "/v1.0/groups/10000000-0000-4000-8000-000000000001/members/10000000-0000-4000-8000-000000000002";
    expect(() => client.deleteReference(path, signal())).toThrow("$ref");
    expect(() => client.deleteReference(`${path}/$ref?unexpected=1`, signal())).toThrow("$ref");
    expect(token).not.toHaveBeenCalled();
    await client.deleteReference(`${path}/$ref`, signal());
    expect(fetcher.mock.calls[0][1]).toMatchObject({ method: "DELETE" });
    expect(fetcher.mock.calls[0][1]?.body).toBeUndefined();
    expect(fetcher).toHaveBeenCalledTimes(1);
  });
  it("parses both forms of Retry-After and falls back visibly through bounded retries", () => {
    expect(retryAfterMs("3", 0)).toBe(3000);
    expect(retryAfterMs("Thu, 01 Jan 2026 00:00:05 GMT", 0, Date.parse("2026-01-01T00:00:00Z"))).toBe(5000);
    expect(retryAfterMs(null, 2)).toBe(4000);
  });
});
