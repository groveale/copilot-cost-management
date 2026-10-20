import type { BudgetConfig, TierName } from "./config";
import {
  formatConsumed, formatField, formatQuantity, formatTime, memberPolicyNote,
  sourceTimestamp, tierGroupIds, tierNames,
  type BudgetSnapshot, type Fields, type Member, type Resource,
} from "./details";

function ResourceStatus({ resource }: { resource: Resource<unknown> | undefined }) {
  if (!resource || resource.state === "loading") return <p className="small muted" role="status">Loading from Graph...</p>;
  if (resource.state === "error") return <p className="error small" role="alert">{resource.error}</p>;
  return null;
}

function JsonDetails({ title, value }: { title: string; value: unknown }) {
  return <details className="json-details"><summary>{title}</summary><pre>{JSON.stringify(value, null, 2)}</pre></details>;
}

function BalanceDetails({ row, fetchedAt, policyName, note }: {
  row: Fields; fetchedAt: string; policyName: string; note: string | null;
}) {
  let source: string;
  try {
    source = formatTime(sourceTimestamp(row));
  } catch (error) {
    source = error instanceof Error ? error.message : String(error);
  }
  return <>
    <dl className="balance-fields">
      <dt>Consumed</dt><dd className="quantity">{formatConsumed(row)}</dd>
      <dt>Remaining</dt><dd>{formatQuantity(row.remainingQuantity)}</dd>
      <dt>Consumption allowed</dt><dd>{formatField(row.isConsumptionAllowed)}</dd>
      <dt>Access decision</dt><dd>{formatField(row.accessDecisionReason)}</dd>
      <dt>Reported policy</dt><dd>{policyName}<span className="mono block">{formatField(row.spendingPolicyId)}</span></dd>
      <dt>Source updated</dt><dd>{source}</dd>
      <dt>Fetched</dt><dd>{formatTime(fetchedAt)}</dd>
    </dl>
    {note && <p className="reading-note">{note}</p>}
    <details className="balance-more"><summary>Balance limits and Graph response</summary>
      <dl>
        <dt>Configured period limit</dt><dd>{formatQuantity(row.configuredPeriodLimit)}</dd>
        <dt>Configured per-user limit</dt><dd>{formatQuantity(row.configuredPerUserLimit)}</dd>
        <dt>Effective capacity</dt><dd>{formatQuantity(row.effectiveCapacityQuantity)}</dd>
      </dl>
      <pre>{JSON.stringify(row, null, 2)}</pre>
    </details>
  </>;
}

function MemberRow({ member, groups, budget, snapshot, serviceId }: {
  member: Member; groups: string[]; budget: BudgetConfig; snapshot: BudgetSnapshot; serviceId: string;
}) {
  const balance = snapshot.balances[member.id];
  const row = balance?.state === "ready" ? balance.data : null;
  const reportedPolicy = typeof row?.spendingPolicyId === "string" ? row.spendingPolicyId.toLowerCase() : null;
  const policyTier = tierNames.find((name) => budget.tiers[name].policyId === reportedPolicy);
  const policy = policyTier ? snapshot.tiers[policyTier].policy : undefined;
  const policyName = policy?.state === "ready" ? formatField(policy.data.fields.displayName) :
    policyTier ? `${policyTier === "standard" ? "Standard" : "Enhanced"} policy` : "Other / unknown policy";
  return <article className="member-row">
    <div className="member-identity">
      <h4>{typeof member.fields.displayName === "string" && member.fields.displayName ? member.fields.displayName : member.id}</h4>
      <p className="small muted">{formatField(member.fields.userPrincipalName ?? member.fields.mail)}</p>
      <span className="badge">{member.kind}</span>
      <p className="small muted member-groups">Direct member of: {groups.map((id) => {
        const directory = snapshot.groups[id].directory;
        return directory.state === "ready" && typeof directory.data.fields.displayName === "string" ?
          directory.data.fields.displayName : id;
      }).join(", ")}</p>
      <details><summary className="small">Directory member details</summary><pre>{JSON.stringify(member.fields, null, 2)}</pre></details>
    </div>
    <div className="member-reading">
      {member.kind !== "user" ? <p className="muted">Consumption is queried for direct users only. Nested groups are not expanded.</p> : <>
        <span className="eyebrow">{serviceId} / CURRENT READING</span>
        <ResourceStatus resource={balance} />
        {balance?.state === "ready" && (row ?
          <BalanceDetails row={row} fetchedAt={balance.fetchedAt} policyName={policyName}
            note={memberPolicyNote(budget, snapshot, member.id, row)} /> :
          <p className="reading-note">No {serviceId} balance was returned. Consumption is unavailable, not zero.</p>)}
      </>}
    </div>
  </article>;
}

function TierPanel({ name, budget, snapshot, serviceId }: {
  name: TierName; budget: BudgetConfig; snapshot: BudgetSnapshot; serviceId: string;
}) {
  const tier = budget.tiers[name];
  const details = snapshot.tiers[name];
  const policy = details.policy.state === "ready" ? details.policy.data.fields : null;
  const groupIds = tierGroupIds(budget, name, snapshot);
  const assigned = details.assignedGroups;
  const members = new Map<string, { member: Member; groups: string[] }>();
  for (const id of groupIds) {
    const result = snapshot.groups[id]?.members;
    if (result?.state === "ready") for (const member of result.data) {
      const existing = members.get(member.id);
      members.set(member.id, { member, groups: [...(existing?.groups ?? []), id] });
    }
  }
  const complete = groupIds.every((id) => snapshot.groups[id]?.members.state === "ready");
  return <section className="card tier-panel" aria-labelledby={`${name}-heading`}>
    <div className="section-heading">
      <div><span className="eyebrow">{name.toUpperCase()} / SPENDING POLICY</span>
        <h2 id={`${name}-heading`}>{policy && typeof policy.displayName === "string" ? policy.displayName : `${name === "standard" ? "Standard" : "Enhanced"} policy`}</h2></div>
      {policy && <span className="badge">{formatField(policy.status)}</span>}
    </div>
    <p className="mono muted">{tier.policyId}</p>
    <ResourceStatus resource={details.policy} />
    {policy && <>
      <div className="allowances">
        <div><span className="small muted">Policy period limit</span><strong>{formatQuantity(policy.periodLimit)}</strong></div>
        <div><span className="small muted">Per-user limit</span><strong>{formatQuantity(policy.perUserLimit)}</strong></div>
      </div>
      <dl className="policy-fields">
        <dt>Billing methods</dt><dd>{formatField(policy.billingMethods)}</dd>
        <dt>Policy type</dt><dd>{formatField(policy["@odata.type"]).replace("#microsoft.graph.", "")}</dd>
        <dt>Last modified</dt><dd>{formatTime(policy.lastModifiedDateTime)}</dd>
      </dl>
      <JsonDetails title="Full spending policy configuration" value={policy} />
    </>}
    <h3>Service configuration</h3>
    <ResourceStatus resource={details.services} />
    {details.services.state === "ready" && <>
      {details.services.data.length === 0 ? <p className="muted">No service settings returned. Omitted services are not assumed disabled.</p> :
        <ul className="service-list">{details.services.data.map((row, i) => <li key={`${formatField(row.serviceId)}-${i}`}>
          <strong>{formatField(row.serviceId)}</strong><span className="badge">{formatField(row.state)}</span>
        </li>)}</ul>}
      <JsonDetails title="Full service settings" value={details.services.data} />
    </>}
    <h3>Assigned groups</h3>
    <ResourceStatus resource={assigned} />
    {assigned.state === "ready" && assigned.data.length === 0 && <p className="reading-note">Graph returned no assigned groups for this policy.</p>}
    <p className="small muted">Assignment comes from Graph. The JSON-mapped group is also shown for directory visibility, explicitly labelled if its policy binding is unverified or absent.</p>
    <div className="group-list">{groupIds.map((id) => {
      const group = snapshot.groups[id];
      const isAssigned = assigned.state === "ready" && assigned.data.some((item) => item.id === id);
      const fields = group.directory.state === "ready" ? group.directory.data.fields : null;
      return <div className="group-card" key={id}>
        <div className="section-heading"><strong>{fields && typeof fields.displayName === "string" ? fields.displayName : id}</strong>
          <span className={`badge ${isAssigned ? "" : "warning"}`}>{isAssigned ? "Assigned in Graph" : assigned.state === "ready" ? "JSON mapping only / not assigned" : "JSON mapping / assignment unverified"}</span></div>
        <p className="mono muted">{id}</p>
        <ResourceStatus resource={group.directory} />
        {fields && <>
          <p className="small muted">Security enabled: {formatField(fields.securityEnabled)} · Mail enabled: {formatField(fields.mailEnabled)} · Group types: {Array.isArray(fields.groupTypes) && fields.groupTypes.length === 0 ? "Static" : formatField(fields.groupTypes)}</p>
          <JsonDetails title="Directory group configuration" value={fields} />
        </>}
        <ResourceStatus resource={group.members} />
        {group.members.state === "ready" && <p className="small">{group.members.data.length} direct members loaded from Graph.</p>}
      </div>;
    })}</div>
    <h3>Members and consumption</h3>
    <p className="small muted">{members.size} distinct direct members loaded{complete ? "" : " so far; membership coverage is incomplete"}. Current {serviceId} user readings are not historical policy totals.</p>
    {complete && members.size === 0 && <p className="muted">No direct members returned for the displayed groups.</p>}
    {[...members.values()].map(({ member, groups }) =>
      <MemberRow key={member.id} member={member} groups={groups} budget={budget} snapshot={snapshot} serviceId={serviceId} />)}
  </section>;
}

export function BudgetDetails({ budget, snapshot, serviceId }: {
  budget: BudgetConfig; snapshot: BudgetSnapshot; serviceId: string;
}) {
  const balances = Object.values(snapshot.balances);
  const discovered = new Set(Object.values(snapshot.groups).flatMap((group) =>
    group.members.state === "ready" ? group.members.data.filter((member) => member.kind === "user").map((member) => member.id) : []));
  const available = balances.filter((balance) => balance.state === "ready" && balance.data !== null).length;
  const failed = balances.filter((balance) => balance.state === "error").length;
  return <>
    <p className="small muted" role="status" aria-live="polite">
      Current readings available for {available} of {discovered.size} discovered users{failed > 0 ? `; ${failed} could not be loaded` : ""}.
      {" "}Refresh started {formatTime(snapshot.requestedAt)}. Source timestamps are shown per reading.
    </p>
    <div className="tier-grid">{tierNames.map((name) =>
      <TierPanel key={name} name={name} budget={budget} snapshot={snapshot} serviceId={serviceId} />)}</div>
  </>;
}
