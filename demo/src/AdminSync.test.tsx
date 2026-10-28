import { renderToStaticMarkup } from "react-dom/server";
import { expect, it } from "vitest";
import type { SyncJob } from "../../api/src/sync-contracts";
import { SyncJobSummary } from "./AdminSync";

it("renders authoritative task state, coverage and reporting-reconciliation warnings", () => {
  const job: SyncJob = { version: 1, revision: 2, id: "job", tenantId: "tenant", initiatorId: "first-admin", executorId: "second-admin",
    actors: ["first-admin", "second-admin"], concurrency: 4, state: "partially-failed", phase: "Finished",
    createdAt: "2026-10-28T13:00:00Z", updatedAt: "2026-10-28T13:01:00Z", inventoryComplete: true, mirrorPending: true,
    policies: [{ id: "policy", displayName: "Finance", type: "selected", state: "failed", error: "Graph HTTP 404" }] };
  const html = renderToStaticMarkup(<SyncJobSummary job={job} />);
  for (const value of ["partially-failed", "second-admin", "2 administrators", "reconciliation", "Graph HTTP 404", "Not collected", "Concurrency 4"]) {
    expect(html).toContain(value);
  }
});
