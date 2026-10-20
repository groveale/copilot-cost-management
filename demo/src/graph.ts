import { object } from "./config";

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code: string,
    readonly requestId: string | null,
  ) {
    super(`${code} (HTTP ${status}): ${message}${requestId ? ` [request ${requestId}]` : ""}`);
    this.name = "GraphError";
  }
}

function graphUrl(path: string): string {
  const url = new URL(path, "https://graph.microsoft.com");
  if (url.origin !== "https://graph.microsoft.com" ||
      !/^\/(beta|v1\.0)\//.test(url.pathname) || url.username || url.password || url.hash) {
    throw new Error("Refusing a Graph request outside the expected HTTPS Graph API routes.");
  }
  return url.href;
}

export function retryAfterMs(value: string | null, attempt: number, now = Date.now()): number {
  if (value !== null) {
    if (/^\d+$/.test(value.trim())) return Number(value) * 1000;
    const time = Date.parse(value);
    if (Number.isFinite(time)) return Math.max(0, time - now);
  }
  return 1000 * 2 ** attempt;
}

export async function wait(ms: number, signal: AbortSignal): Promise<void> {
  signal.throwIfAborted();
  await new Promise<void>((resolve, reject) => {
    const onAbort = () => {
      clearTimeout(timer);
      reject(signal.reason);
    };
    const timer = setTimeout(() => {
      signal.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    signal.addEventListener("abort", onAbort, { once: true });
  });
}

export class GraphClient {
  private tail: Promise<void> = Promise.resolve();

  constructor(
    private readonly token: () => Promise<string>,
    private readonly fetcher: typeof fetch = fetch.bind(globalThis),
  ) {}

  get(path: string, signal: AbortSignal): Promise<unknown> {
    const result = this.tail.then(() => this.read(path, signal));
    this.tail = result.then(() => undefined, () => undefined);
    return result;
  }

  private async read(path: string, signal: AbortSignal): Promise<unknown> {
    const url = graphUrl(path);
    for (let attempt = 0; ; attempt++) {
      signal.throwIfAborted();
      const token = await this.token();
      const response = await this.fetcher(url, {
        headers: { Authorization: `Bearer ${token}`, Accept: "application/json" },
        signal: AbortSignal.any([signal, AbortSignal.timeout(30_000)]),
        redirect: "error",
        cache: "no-store",
      });
      if (response.status === 429 && attempt < 3) {
        const delay = retryAfterMs(response.headers.get("Retry-After"), attempt);
        if (delay > 60_000) {
          throw new GraphError(`Retry after ${Math.ceil(delay / 1000)} seconds, then refresh the details.`,
            429, "Throttled", response.headers.get("request-id"));
        }
        await wait(delay, signal);
        continue;
      }
      const raw = await response.text();
      if (!response.ok) {
        let message = raw.slice(0, 500) || response.statusText;
        let code = "GraphRequestFailed";
        try {
          const data = object(JSON.parse(raw), "Graph error");
          if (data.error && typeof data.error === "object" && !Array.isArray(data.error)) {
            const error = object(data.error, "Graph error details");
            if (typeof error.message === "string") message = error.message;
            if (typeof error.code === "string") code = error.code;
          }
        } catch {
          message = `Non-JSON or malformed error response: ${message}`;
        }
        throw new GraphError(message, response.status, code, response.headers.get("request-id"));
      }
      try {
        return JSON.parse(raw);
      } catch {
        throw new Error(`Graph returned HTTP ${response.status} but not valid JSON for ${new URL(url).pathname}.`);
      }
    }
  }

  async list(path: string, signal: AbortSignal): Promise<unknown[]> {
    const values: unknown[] = [];
    let next: string | undefined = path;
    const visited = new Set<string>();
    while (next) {
      const url = graphUrl(next);
      if (visited.has(url)) throw new Error("Graph returned a repeated continuation link.");
      visited.add(url);
      const page = object(await this.get(url, signal), "Graph collection");
      if (!Array.isArray(page.value)) throw new Error("Graph collection did not contain a value array.");
      values.push(...page.value);
      const link = page["@odata.nextLink"];
      if (link !== undefined && link !== null && typeof link !== "string") {
        throw new Error("Graph returned an invalid continuation link.");
      }
      next = link || undefined;
    }
    return values;
  }
}
