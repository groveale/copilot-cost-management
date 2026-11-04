import { beforeEach, expect, it, vi } from "vitest";
import { createHash } from "node:crypto";
import { SharePointPermissions, type PermissionDestination } from "../src/reporting-sharepoint.js";

const id = (n: number) => `10000000-0000-4000-8000-${String(n).padStart(12, "0")}`;
const d: PermissionDestination = { siteId: `example.sharepoint.com,${id(5)},${id(6)}`, siteUrl: "https://example.sharepoint.com/sites/reporting",
  policyListId: id(7), balanceListId: id(8), jobListId: id(9), adminGroupId: 20, maxScopes: 5000, scopeReserve: 250 };
type Item = { Id: number; TenantId: string; PolicyId?: string; PolicyKey?: string; FileSystemObjectType: number;
  SchemaVersion?: number; SnapshotId?: string; SnapshotKey?: string; UserId?: string; ServiceId?: string; ObservationKey?: string };
let acls: Map<string, Map<number, number[]>>;
let unique: Set<string>;
let lists: Map<string, Item[]>;
let calls: string[];
let failDelete: boolean;
let externalNext: boolean;
let readbackLie: boolean;
let members: number[];
let principalMismatch: boolean;
let unsafeRead: boolean;
const listPath = (n: number) => `web/lists(guid'${id(n)}')`;
const row = (n: number, policy = id(11)): Item => {
  const SnapshotKey = `${id(1)}:${policy}:${id(30)}`, UserId = id(40 + n), ServiceId = "cowork";
  return { Id: n, TenantId: id(1), PolicyId: policy, PolicyKey: `${id(1)}:${policy}`, FileSystemObjectType: 0, SchemaVersion: 1,
    SnapshotId: id(30), SnapshotKey, UserId, ServiceId,
    ObservationKey: createHash("sha256").update(JSON.stringify([SnapshotKey, UserId, ServiceId])).digest("hex") };
};
const response = (value: unknown, status = 200) => new Response(JSON.stringify(value), { status });
let fetcher: typeof fetch;
beforeEach(() => {
  acls = new Map(); unique = new Set(); calls = []; members = [102, 199]; failDelete = false; externalNext = false; readbackLie = false; principalMismatch = false; unsafeRead = false;
  lists = new Map([[listPath(7), [row(1), row(2, id(99))]], [listPath(8), [row(1), row(2), row(3), row(4, id(99)), { ...row(5), PolicyKey: "bad-key" }]],
    [listPath(9), [{ Id: 1, TenantId: id(1), FileSystemObjectType: 0 }]]]);
  for (const [path, rows] of lists) {
    acls.set(path, new Map([[199, [1073741826]]]));
    for (const item of rows) { const key = `${path}/items(${item.Id})`; acls.set(key, new Map([[198, [1073741826]]])); unique.add(key); }
  }
  fetcher = vi.fn(async (input, init) => {
    const url = new URL(String(input)), method = init?.method ?? "GET";
    const path = decodeURIComponent(url.pathname.split("/_api/")[1] ?? url.pathname.replace("/v1.0/", ""));
    calls.push(`${method} ${url.origin} ${path}`);
    expect(init?.redirect).toBe("error");
    if (path === "site") return response({ Id: id(5) });
    if (path === "web") return response({ Id: id(6) });
    if (path === "web/roledefinitions") return response({ value: [{ Id: 1073741829, RoleTypeKind: 5 },
      { Id: 1073741826, RoleTypeKind: 2, BasePermissions: { High: "176", Low: String(138612833 | (unsafeRead ? 4 : 0)) } },
      { Id: 1073741825, RoleTypeKind: 1, BasePermissions: { High: "48", Low: "134287360" } }] });
    if (path.startsWith("users/")) {
      const objectId = path.split("/")[1];
      return response({ id: objectId, userPrincipalName: `user${Number(objectId.slice(-12))}@example.test` });
    }
    if (path === "web/ensureuser") {
      const login = JSON.parse(String(init!.body)).logonName;
      return response({ Id: 100 + Number(login.match(/user(\d+)/)[1]), LoginName: login, PrincipalType: 1 });
    }
    if (path.startsWith("web/getuserbyid(")) {
      const number = Number(path.match(/getuserbyid\((\d+)/)![1]);
      return response({ Id: number, LoginName: `i:0#.f|membership|user${number - 100}@example.test`,
        PrincipalType: 1, AadObjectId: { NameId: id(principalMismatch ? 99 : number - 100) } });
    }
    if (path === "web/sitegroups/getbyid(20)") return response({
      Id: 20, Description: "CopilotCostSync/v1/administrators", AllowMembersEditMembership: false,
      AllowRequestToJoinLeave: false, AutoAcceptRequestToJoinLeave: false, Owner: { Id: 102 },
    });
    if (path === "web/sitegroups/getbyid(20)/users") return response({ value: members.map((Id) => ({ Id })) });
    if (path.startsWith("web/sitegroups/getbyid(20)/users/removebyid(")) { members = members.filter((n) => n !== Number(path.match(/removebyid\((\d+)/)![1])); return response({}); }
    const list = [...lists.keys()].find((key) => path.startsWith(key));
    if (list) {
      const base = path.split(/\/(?:roleassignments|breakroleinheritance)/)[0];
      if (path.includes("/breakroleinheritance")) { unique.add(base); acls.set(base, new Map([[102, [1073741829]]])); return response({}); }
      if (path.includes("/roleassignments/getbyprincipalid(")) {
        if (failDelete) return response({}, 403);
        const principal = Number(path.match(/getbyprincipalid\((\d+)/)![1]);
        acls.get(base)!.delete(principal); return response({});
      }
      if (path.includes("/roleassignments/addroleassignment(")) {
        const [, principal, role] = path.match(/principalid=(\d+),roledefid=(\d+)/)!;
        if (!readbackLie) acls.get(base)!.set(Number(principal), [Number(role)]);
        return response({});
      }
      if (path.endsWith("/roleassignments")) return response({ value: [...acls.get(base)!].map(([PrincipalId, roles]) =>
        ({ PrincipalId, RoleDefinitionBindings: roles.map((Id) => ({ Id })) })) });
      if (path.endsWith("/items")) return response({ value: lists.get(list), ...(externalNext ? { "@odata.nextLink": "https://untrusted.example/data" } : {}) });
      if (path === list) return response({ Id: list.match(/guid'([^']+)/)![1], Description: `CopilotCostSync/v1/${list === listPath(7) ? "policyListId" : list === listPath(8) ? "balanceListId" : "jobListId"}`,
        BaseTemplate: 100, HasUniqueRoleAssignments: unique.has(path), ItemCount: lists.get(list)!.length });
      if (/\/items\(\d+\)$/.test(path)) return response({ HasUniqueRoleAssignments: unique.has(path) });
    }
    throw new Error(`Unexpected mock request: ${method} ${url}`);
  }) as typeof fetch;
});
const permissions = (destination = d) => new SharePointPermissions(destination, id(1), [id(2)],
  { token: async (resource) => `synthetic-${resource}` }, () => {}, fetcher);
it("locks all lists, verifies admin membership, and reconciles every stale/staging balance row", async () => {
  expect(await permissions().reconcile(new Map([[id(11), [id(3)]]]))).toEqual({ policies: 1, rows: 8 });
  expect(members).toEqual([102]);
  for (const path of lists.keys()) expect([...acls.get(path)!]).toEqual([[20, [1073741829]]]);
  for (const n of [1, 2, 3]) expect(acls.get(`${listPath(8)}/items(${n})`)).toEqual(new Map([[20, [1073741829]], [103, [1073741826]]]));
  for (const key of [`${listPath(7)}/items(2)`, `${listPath(8)}/items(4)`, `${listPath(8)}/items(5)`, `${listPath(9)}/items(1)`]) {
    expect([...acls.get(key)!]).toEqual([[20, [1073741829]]]);
  }
  expect(calls.some((call) => /PATCH|MERGE/.test(call))).toBe(false);
  await permissions().reconcile(new Map(), true);
  for (const acl of acls.values()) expect([...acl]).toEqual([[20, [1073741829]]]);
});
it("refuses failed revocation and does not grant replacement owners on that item", async () => {
  unique.add(listPath(7)); acls.set(listPath(7), new Map([[20, [1073741829]]]));
  failDelete = true;
  await expect(permissions().reconcile(new Map([[id(11), [id(3)]]]))).rejects.toMatchObject({ code: "reporting-permissions" });
  expect(calls.some((call) => call.includes("principalid=103"))).toBe(false);
});
it("requires role readback, not just successful permission POSTs", async () => {
  readbackLie = true;
  await expect(permissions().reconcile(new Map())).rejects.toThrow("readback");
});
it("checks existing rows plus staging headroom and never silently increases the 5,000 ceiling", async () => {
  await expect(permissions({ ...d, maxScopes: 255, scopeReserve: 250 }).reconcile(new Map())).rejects.toMatchObject({ code: "reporting-scope-limit" });
  expect(calls.some((call) => call.includes("/items(1)/roleassignments"))).toBe(false);
});
it("rejects pagination outside the pinned site before sending a bearer token", async () => {
  externalNext = true;
  await expect(permissions().reconcile(new Map())).rejects.toThrow("untrusted");
  expect(calls.some((call) => call.includes("untrusted.example"))).toBe(false);
});
it("preserves only implicit Limited Access navigation grants on parents, never broad read", async () => {
  unique.add(listPath(7));
  acls.set(listPath(7), new Map([[20, [1073741829]], [103, [1073741825]], [199, [1073741826]]]));
  await permissions().reconcile(new Map([[id(11), [id(3)]]]));
  expect(acls.get(listPath(7))).toEqual(new Map([[20, [1073741829]], [103, [1073741825]]]));
});
it("refuses a stale SharePoint login even when the current UPN resolves successfully", async () => {
  principalMismatch = true;
  await expect(permissions().reconcile(new Map([[id(11), [id(3)]]]))).rejects.toMatchObject({ code: "reporting-principal-mismatch" });
  expect(calls.some((call) => call.startsWith("POST") && call.includes("/roleassignments/"))).toBe(false);
});
it("returns validated SharePoint Conditional Access challenges for both unauthorized and forbidden responses", async () => {
  const claims = Buffer.from(JSON.stringify({ access_token: { acrs: { value: "c1" } } })).toString("base64");
  for (const status of [401, 403]) {
    const fetcher = vi.fn(async () => new Response("denied", { status, headers: { "WWW-Authenticate": `Bearer error="insufficient_claims", claims="${claims}"` } })) as typeof fetch;
    const service = new SharePointPermissions(d, id(1), [id(2)], { token: async () => "synthetic" }, () => {}, fetcher);
    await expect(service.reconcile(new Map())).rejects.toMatchObject({ code: "reporting-interaction-required", challenge: expect.stringContaining(claims) });
  }
});
it("does not mistake a customized editable Read permission level for owner read-only access", async () => {
  unsafeRead = true;
  await expect(permissions().reconcile(new Map([[id(11), [id(3)]]]))).rejects.toMatchObject({ code: "reporting-permission-level" });
  expect(calls.some((call) => call.startsWith("POST"))).toBe(false);
});
it("ensures first-time owner principals before looking up their SharePoint numeric identity", async () => {
  await permissions().reconcile(new Map([[id(11), [id(3)]]]));
  const graphUser = calls.findIndex((call) => call.includes(`users/${id(3)}`));
  const ensure = calls.findIndex((call, index) => index > graphUser && call.endsWith("web/ensureuser"));
  const verified = calls.findIndex((call) => call.endsWith("web/getuserbyid(103)"));
  expect(graphUser).toBeGreaterThan(-1);
  expect(ensure).toBeGreaterThan(graphUser);
  expect(verified).toBeGreaterThan(ensure);
  expect(acls.get(`${listPath(8)}/items(1)`)?.get(103)).toEqual([1073741826]);
});
it.each(["directory", "ensureuser"])("fails closed when the %s principal resolver cannot materialize the current owner", async (failure) => {
  const original = fetcher;
  fetcher = vi.fn(async (input, init) => {
    const url = new URL(String(input));
    if ((failure === "directory" && url.hostname === "graph.microsoft.com" && url.pathname.endsWith(`/users/${id(3)}`)) ||
        (failure === "ensureuser" && url.pathname.endsWith("/web/ensureuser") && String(init?.body).includes("user3@example.test"))) {
      return response({}, failure === "directory" ? 404 : 403);
    }
    return original(input, init);
  }) as typeof fetch;
  await expect(permissions().reconcile(new Map([[id(11), [id(3)]]]))).rejects.toMatchObject({ code: "reporting-permissions" });
  expect(calls.some((call) => call.startsWith("POST") && call.includes("/roleassignments/"))).toBe(false);
});
