import { useEffect, useRef, useState } from "react";
import type { AccountInfo } from "@azure/msal-browser";
import type { DemoAuth } from "./auth";
import type { DemoConfig } from "./config";
import type { Appearance, ReportingSource, usePreferences } from "./preferences";

export function identityDiagnostics(account: AccountInfo) {
  const claims: Record<string, unknown> = {};
  for (const key of ["iss", "aud", "tid", "oid", "name", "preferred_username", "iat", "nbf", "exp", "auth_time", "amr", "acrs"]) {
    const value = account.idTokenClaims?.[key];
    if (typeof value === "string" || typeof value === "number" || typeof value === "boolean" ||
        (Array.isArray(value) && value.every((item) => typeof item === "string"))) claims[key] = value;
  }
  return claims;
}

export function Settings({ preferences, account, auth, config, locked, onAuthenticate }: {
  preferences: ReturnType<typeof usePreferences>; account: AccountInfo; auth: Pick<DemoAuth, "diagnostics">; config: DemoConfig;
  locked: boolean; onAuthenticate: (signOut?: boolean) => void;
}) {
  const dialog = useRef<HTMLDialogElement>(null);
  const button = useRef<HTMLButtonElement>(null);
  const [open, setOpen] = useState(false);
  const [diagnostics, setDiagnostics] = useState(false);
  useEffect(() => { setDiagnostics(false); }, [account]);
  useEffect(() => {
    if (open) dialog.current?.showModal();
    else dialog.current?.close();
  }, [open]);
  return <>
    <button type="button" className="secondary settings-cog" ref={button} aria-label="Settings"
      aria-haspopup="dialog" onClick={() => setOpen(true)}>
      <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" strokeWidth="1.8" aria-hidden="true">
        <path d="m9 3-.6 2.2-1.6.9-2.2-.6-3 5.2 1.6 1.6v1.8l-1.6 1.6 3 5.2 2.2-.6 1.6.9L9 23h6l.6-2.2 1.6-.9 2.2.6 3-5.2-1.6-1.6v-1.8l1.6-1.6-3-5.2-2.2.6-1.6-.9L15 3Z" transform="translate(1 -1) scale(.92)" />
        <circle cx="12" cy="12" r="3" />
      </svg> Settings
    </button>
    <dialog ref={dialog} className="settings-dialog" aria-labelledby="settings-heading"
      onCancel={(event) => { event.preventDefault(); setOpen(false); }}
      onKeyDown={(event) => { if (event.key === "Escape") { event.preventDefault(); setOpen(false); } }}
      onClose={() => { setOpen(false); setDiagnostics(false); button.current?.focus(); }}>
      <div className="section-heading"><h2 id="settings-heading">Settings</h2>
        <button type="button" className="secondary" autoFocus onClick={() => setOpen(false)}>Close settings</button></div>
      <label>Reporting data source
        <select value={preferences.value.source} disabled={locked}
          onChange={(event) => preferences.save("source", event.target.value as ReportingSource)}>
          <option value="realtime">Real time - call the APIs</option>
          <option value="sync">Sync - read SharePoint snapshots</option>
        </select>
      </label>
      <p className="small muted">Remembered in this browser. Actions and approvals always use live authorization.
        Sync reports do not update until another sync completes.</p>
      {locked && <p className="notice" role="status">Wait for the current action to finish, or pause sync, before switching data source or account.</p>}
      <label>Appearance
        <select value={preferences.value.appearance}
          onChange={(event) => preferences.save("appearance", event.target.value as Appearance)}>
          <option value="system">System</option><option value="light">Light</option><option value="dark">Dark</option>
        </select>
      </label>
      {preferences.error && <p className="error" role="alert">{preferences.error}</p>}
      <h3>Sign-in details</h3>
      <dl className="metadata-grid">
        <div><dt>Name</dt><dd>{account.name ?? "Not reported"}</dd></div>
        <div><dt>Account</dt><dd>{account.username}</dd></div>
        <div><dt>Tenant</dt><dd className="mono">{account.tenantId}</dd></div>
        <div><dt>Object ID</dt><dd className="mono">{account.localAccountId}</dd></div>
        <div><dt>Application ID</dt><dd className="mono">{config.clientId}</dd></div>
        <div><dt>Authority</dt><dd>https://login.microsoftonline.com/{config.tenantId}</dd></div>
      </dl>
      <details onToggle={(event) => setDiagnostics(event.currentTarget.open)}>
        <summary>Decoded sign-in claims and token diagnostics</summary>
        <p className="small muted">Local diagnostics only, not proof of authorization or PIM activation.
          Access tokens are opaque and are never displayed here. Opening Settings does not request new permissions.</p>
        {diagnostics && <pre>{JSON.stringify({ signInClaims: identityDiagnostics(account),
          acquisitions: auth.diagnostics(account) }, null, 2)}</pre>}
      </details>
      <div className="actions">
        <button className="secondary" disabled={locked} onClick={() => onAuthenticate()}>Switch account / renew sign-in</button>
        <button className="secondary" disabled={locked} onClick={() => onAuthenticate(true)}>Sign out</button>
      </div>
    </dialog>
  </>;
}
