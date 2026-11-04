import { createHash } from "node:crypto";
import { ApiError, uuid } from "./model.js";
import { ReportingInteractionRequired, validatedClaims, type ReportingAuth } from "./reporting-auth.js";

export interface ReportingRow { id: number; policyId?: string }
export interface PermissionDestination {
  siteId: string; siteUrl: string; policyListId: string; balanceListId: string; jobListId: string;
  adminGroupId: number; maxScopes: number; scopeReserve: number;
}
export interface ReportingPermissions {
  reconcile(owners: Map<string, string[]>, revokeOnly?: boolean): Promise<{ policies: number; rows: number }>;
}
type Json = Record<string, any>;
const permissionBits = (kinds: number[]) => kinds.reduce((value, kind) => value | (1n << BigInt(kind - 1)), 0n);
function permissionMask(role: Json): bigint {
  const { High, Low } = role.BasePermissions ?? {};
  if (!/^\d+$/.test(String(High)) || !/^\d+$/.test(String(Low)) || BigInt(High) > 0xffffffffn || BigInt(Low) > 0xffffffffn) {
    throw new Error("SharePoint permission-level mask could not be verified.");
  }
  return (BigInt(High) << 32n) | BigInt(Low);
}
export class SharePointPermissions implements ReportingPermissions {
  constructor(private readonly destination: PermissionDestination, private readonly tenantId: string,
    private readonly admins: string[], private readonly auth: ReportingAuth,
    private readonly assertLock: () => void = () => {}, private readonly fetcher: typeof fetch = fetch) {}
  private async request(path: string, method = "GET", body?: unknown, graph = false): Promise<Json> {
    this.assertLock();
    const base = graph ? "https://graph.microsoft.com/v1.0/" : `${this.destination.siteUrl.replace(/\/$/, "")}/_api/`;
    const url = new URL(path, base);
    if (!url.href.startsWith(base) || url.username || url.password || url.hash) throw new Error("Refusing an untrusted permission-service URL.");
    const token = await this.auth.token(graph ? "graph" : "sharepoint");
    this.assertLock();
    const response = await this.fetcher(url, {
      method, redirect: "error", signal: AbortSignal.timeout(30000),
      headers: { Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata", ...(body === undefined ? {} : { "Content-Type": "application/json;odata=nometadata" }) },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
    });
    if (!response.ok) {
      const challenge = response.headers.get("WWW-Authenticate")?.match(/(?:^|[, ])claims="([^"]+)"/i)?.[1];
      if ([401, 403].includes(response.status) && challenge && validatedClaims(challenge)) {
        throw new ReportingInteractionRequired(this.tenantId, validatedClaims(challenge));
      }
      throw new ApiError(503, "reporting-permissions", "SharePoint permission reconciliation failed. Owner reads/publication remain blocked; verify delegated rights and rerun verification.");
    }
    if (response.status === 204 || response.headers.get("content-length") === "0") return {};
    const text = await response.text();
    if (!text) return {};
    const value = JSON.parse(text);
    return value.d ?? value;
  }
  private async pages(path: string): Promise<Json[]> {
    const rows: Json[] = [], seen = new Set<string>();
    while (path) {
      if (seen.has(path)) throw new Error("Repeated SharePoint continuation.");
      seen.add(path);
      const page = await this.request(path);
      const values = page.value ?? page.results;
      if (!Array.isArray(values)) throw new Error("Invalid SharePoint permission collection.");
      rows.push(...values);
      if (rows.length > 5000) throw new ApiError(409, "reporting-scope-limit", "Reporting exceeds the 5,000-scope isolation design. Use a separately approved partitioning design.");
      path = page["@odata.nextLink"] ?? page["odata.nextLink"] ?? page.__next ?? "";
      if (typeof path !== "string") throw new Error("Invalid permission continuation.");
    }
    return rows;
  }
  private async principal(objectId: string): Promise<number> {
    const user = await this.request(`users/${uuid.parse(objectId)}?$select=id,userPrincipalName`, "GET", undefined, true);
    if (typeof user.id !== "string" || user.id.toLowerCase() !== objectId || typeof user.userPrincipalName !== "string" ||
        !/^[^@\s|]+@[^@\s|]+$/.test(user.userPrincipalName)) throw new Error("Directory user identity could not be verified.");
    const login = `i:0#.f|membership|${user.userPrincipalName.toLowerCase()}`;
    const ensured = await this.request("web/ensureuser", "POST", { logonName: login });
    if (!Number.isSafeInteger(ensured.Id) || ensured.Id < 1 || ensured.LoginName?.toLowerCase() !== login || ensured.PrincipalType !== 1) {
      throw new Error("SharePoint did not resolve the expected user principal.");
    }
    const verified = await this.request(`web/getuserbyid(${ensured.Id})?$select=Id,LoginName,PrincipalType,AadObjectId`);
    if (verified.Id !== ensured.Id || verified.LoginName?.toLowerCase() !== login || verified.PrincipalType !== 1 ||
        verified.AadObjectId?.NameId?.toLowerCase() !== objectId) {
      throw new ApiError(409, "reporting-principal-mismatch",
        "SharePoint's resolved directory object ID does not match the current owner. Repair stale site-user identity mappings before retrying.");
    }
    return ensured.Id;
  }
  private async acl(path: string): Promise<Map<number, number[]>> {
    const rows = await this.pages(`${path}/roleassignments?$expand=RoleDefinitionBindings&$select=PrincipalId,RoleDefinitionBindings/Id`);
    const result = new Map<number, number[]>();
    for (const row of rows) {
      const roles = row.RoleDefinitionBindings?.results ?? row.RoleDefinitionBindings;
      if (!Number.isSafeInteger(row.PrincipalId) || !Array.isArray(roles) ||
          roles.some((role: Json) => !Number.isSafeInteger(role.Id)) || result.has(row.PrincipalId)) throw new Error("Invalid permission readback.");
      result.set(row.PrincipalId, roles.map((role: Json) => role.Id).sort((a: number, b: number) => a - b));
    }
    return result;
  }
  private async exact(path: string, expected: Map<number, number[]>, limitedAccess?: number, knownUnique = false) {
    if (!knownUnique) {
      const item = await this.request(`${path}?$select=HasUniqueRoleAssignments`);
      if (item.HasUniqueRoleAssignments !== true) throw new Error("SharePoint permission inheritance was not broken.");
    }
    const trimNavigation = (acl: Map<number, number[]>) => new Map([...acl].filter(([, roles]) =>
      limitedAccess === undefined || !roles.length || roles.some((role) => role !== limitedAccess)));
    // SharePoint automatically creates Limited Access on a parent when an item
    // is shared. It grants navigation, not list data access; removing it can
    // invalidate otherwise correct child-item permissions.
    const current = trimNavigation(await this.acl(path));
    const administrator = this.destination.adminGroupId;
    const adminRoles = expected.get(administrator)!;
    let changed = false;
    // Keep repair access throughout: add Full Control before stripping an
    // implicit caller grant. Never remove the administrator's Full Control.
    for (const role of adminRoles) if (!current.get(administrator)?.includes(role)) {
      await this.request(`${path}/roleassignments/addroleassignment(principalid=${administrator},roledefid=${role})`, "POST");
      changed = true;
    }
    for (const role of current.get(administrator) ?? []) if (!adminRoles.includes(role)) {
      await this.request(`${path}/roleassignments/removeroleassignment(principalid=${administrator},roledefid=${role})`, "POST");
      changed = true;
    }
    // Revoke first. In particular never grant a replacement owner before removing
    // the previous owner. The durable ready marker is already invalidated.
    for (const [principal, roles] of current) {
      if (principal === administrator) continue;
      if (JSON.stringify(roles) !== JSON.stringify(expected.get(principal))) {
        await this.request(`${path}/roleassignments/getbyprincipalid(${principal})`, "DELETE");
        changed = true;
      }
    }
    for (const [principal, roles] of expected) {
      if (principal === administrator) continue;
      if (JSON.stringify(roles) !== JSON.stringify(current.get(principal))) {
        for (const role of roles) await this.request(`${path}/roleassignments/addroleassignment(principalid=${principal},roledefid=${role})`, "POST");
        changed = true;
      }
    }
    if (!changed) return;
    const actual = trimNavigation(await this.acl(path));
    if (actual.size !== expected.size || [...expected].some(([id, roles]) => JSON.stringify(actual.get(id)) !== JSON.stringify(roles))) {
      throw new Error("SharePoint role-assignment readback did not match.");
    }
  }
  async reconcile(owners: Map<string, string[]>, revokeOnly = false) {
    const d = this.destination;
    const [, siteId, webId] = d.siteId.split(",");
    const site = await this.request("site?$select=Id");
    const web = await this.request("web?$select=Id");
    if (site.Id?.toLowerCase() !== siteId || web.Id?.toLowerCase() !== webId) throw new Error("Configured SharePoint web does not match the selected site ID.");
    const roles = await this.pages("web/roledefinitions?$select=Id,RoleTypeKind,BasePermissions");
    const role = (kind: number) => {
      const matches = roles.filter((value) => value.RoleTypeKind === kind && Number.isSafeInteger(value.Id));
      if (matches.length !== 1) throw new Error("SharePoint built-in role is ambiguous.");
      return matches[0].Id as number;
    };
    const full = role(5), read = role(2), limitedAccess = role(1);
    const readMask = permissionMask(roles.find((value) => value.Id === read)!);
    const writeMask = permissionBits([2, 3, 4, 5, 8, 9, 10, 12, 19, 20, 21, 24, 25, 26, 29, 30, 31, 39, 41, 63]);
    const navigationMask = permissionBits([13, 17, 18, 28, 37, 38]);
    if (!(readMask & 1n) || (readMask & writeMask) || (permissionMask(roles.find((value) => value.Id === limitedAccess)!) & ~navigationMask)) {
      throw new ApiError(409, "reporting-permission-level", "SharePoint Read/Limited Access levels were customized with unsafe rights. Restore read-only/navigation permissions before verification.");
    }
    const principals = new Map<string, number>();
    for (const id of [...new Set([...this.admins, ...(!revokeOnly ? [...owners.values()].flat() : [])])]) {
      principals.set(id, await this.principal(id));
    }
    const adminIds = new Set(this.admins.map((id) => principals.get(id)!));
    const groupPath = `web/sitegroups/getbyid(${d.adminGroupId})`;
    const group = await this.request(`${groupPath}?$expand=Owner&$select=Id,Description,AllowMembersEditMembership,AllowRequestToJoinLeave,AutoAcceptRequestToJoinLeave,Owner/Id`);
    if (group.Id !== d.adminGroupId || group.Description !== "CopilotCostSync/v1/administrators" ||
        group.AllowMembersEditMembership !== false || group.AllowRequestToJoinLeave !== false ||
        group.AutoAcceptRequestToJoinLeave !== false || !adminIds.has(group.Owner?.Id)) {
      throw new ApiError(409, "reporting-admin-group", "Rerun the explicit reporting lockdown setup: the reporting group must be owned by a current configured administrator and disallow member edits.");
    }
    let members = await this.pages(`${groupPath}/users?$select=Id`);
    for (const member of members) if (!adminIds.has(member.Id)) await this.request(`${groupPath}/users/removebyid(${Number(member.Id)})`, "POST");
    for (const id of this.admins) if (!members.some((member) => member.Id === principals.get(id))) {
      const user = await this.request(`web/getuserbyid(${principals.get(id)})?$select=LoginName`);
      await this.request(`${groupPath}/users`, "POST", { LoginName: user.LoginName });
    }
    members = await this.pages(`${groupPath}/users?$select=Id`);
    if (members.length !== adminIds.size || members.some((member) => !adminIds.has(member.Id))) throw new Error("Administrator group membership verification failed.");
    const adminAcl = new Map([[d.adminGroupId, [full]]]);
    const definitions = [
      [d.policyListId, "policyListId"], [d.balanceListId, "balanceListId"], [d.jobListId, "jobListId"],
    ] as const;
    const inventory: { path: string; rows: ReportingRow[] }[] = [];
    for (const [id, key] of definitions) {
      const path = `web/lists(guid'${uuid.parse(id)}')`;
      const list = await this.request(`${path}?$select=Id,Description,BaseTemplate,HasUniqueRoleAssignments,ItemCount`);
      if (list.Id?.toLowerCase() !== id || list.BaseTemplate !== 100 || list.Description !== `CopilotCostSync/v1/${key}`) {
        throw new Error("Refusing permissions on an unmanaged SharePoint list.");
      }
      if (!list.HasUniqueRoleAssignments) await this.request(`${path}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)`, "POST");
      await this.exact(path, adminAcl, limitedAccess);
      const fields = key === "jobListId" ? "Id,TenantId,FileSystemObjectType" :
        `Id,TenantId,PolicyId,PolicyKey,FileSystemObjectType,SchemaVersion${key === "balanceListId" ? ",SnapshotId,SnapshotKey,UserId,ServiceId,ObservationKey" : ""}`;
      const values = await this.pages(`${path}/items?$select=${fields}&$top=500`);
      if (values.length + 1 + d.scopeReserve > d.maxScopes) throw new ApiError(409, "reporting-scope-limit",
        `Reporting list needs ${values.length + 1} scopes plus ${d.scopeReserve} reserved. Maximum ${d.maxScopes}; remove obsolete generations administratively or adopt an approved partitioning design.`);
      const rows = values.map((value) => {
        let valid = value.TenantId?.toLowerCase() === this.tenantId && value.FileSystemObjectType === 0 &&
          value.SchemaVersion === 1 && uuid.safeParse(value.PolicyId).success && value.PolicyKey === `${this.tenantId}:${value.PolicyId}`;
        if (key === "balanceListId") {
          valid &&= uuid.safeParse(value.SnapshotId).success && uuid.safeParse(value.UserId).success &&
            typeof value.ServiceId === "string" && value.ServiceId.length > 0 && value.ServiceId.length <= 255 &&
            value.SnapshotKey === `${value.PolicyKey}:${value.SnapshotId}` &&
            value.ObservationKey === createHash("sha256").update(JSON.stringify([value.SnapshotKey, value.UserId, value.ServiceId])).digest("hex");
        }
        if (!Number.isSafeInteger(value.Id) || value.Id < 1) throw new Error("Invalid managed item ID.");
        // Malformed/foreign and unmapped rows are never owner-readable.
        return { id: value.Id, ...(key !== "jobListId" && valid ? { policyId: value.PolicyId.toLowerCase() as string } : {}) };
      });
      inventory.push({ path, rows });
    }
    let count = 0;
    const work = inventory.flatMap(({ path, rows }) => rows.map((row) => ({ path, row })));
    let next = 0;
    let failure: unknown;
    await Promise.all(Array.from({ length: Math.min(8, work.length) }, async () => {
      while (next < work.length && !failure) {
        const { path, row } = work[next++];
        try {
          const item = `${path}/items(${row.id})`;
          const current = await this.request(`${item}?$select=HasUniqueRoleAssignments`);
          if (!current.HasUniqueRoleAssignments) await this.request(`${item}/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=false)`, "POST");
          const expected = new Map(adminAcl);
          if (!revokeOnly && row.policyId) for (const owner of owners.get(row.policyId) ?? []) {
            if (!adminIds.has(principals.get(owner)!)) expected.set(principals.get(owner)!, [read]);
          }
          // Every generation is covered, not just the current snapshot.
          await this.exact(item, expected, undefined, current.HasUniqueRoleAssignments === true);
          count++;
        } catch (error) { failure ??= error; }
      }
    }));
    if (failure) throw failure;
    return { policies: owners.size, rows: count };
  }
}
