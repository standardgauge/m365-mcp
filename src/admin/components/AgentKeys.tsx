import React, { useCallback, useEffect, useState } from 'react';

const SERVICE_LABELS: Record<string, string> = {
  mail: 'Mail', sharepoint: 'SharePoint', calendar: 'Calendar', onedrive: 'OneDrive',
  onenote: 'OneNote', contacts: 'Contacts', teams: 'Teams',
};
const SERVICE_ORDER = ['mail', 'sharepoint', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'];

interface SitePick { id: string; name: string }
interface KeyPolicy { services: string[]; readOnlyServices: string[]; allowedSites: SitePick[] }
interface AgentKey {
  keyId: string; label: string; userId: string; userEmail: string; policy: KeyPolicy;
  createdAt: number; expiresAt: number; lastUsedAt: number | null; revokedAt: number | null; revokedBy: string | null;
  reminderSet?: boolean;
}
interface TenantPolicy { allowKeys: boolean; disabledServices: string[]; readOnlyServices: string[]; maxKeyLifetimeDays: number }
type ReminderResult = { eventId: string } | { skipped: string } | { error: string };

interface Props {
  userId: string;
  /** Services the caller can use: tenant-enabled minus their own opt-outs. */
  grantableServices: string[];
  /** The tenant's SharePoint allow-list, when it has one; offered as site picks. */
  allowedSites?: SitePick[];
}

function when(ms: number | null | undefined): string {
  return ms ? new Date(ms).toLocaleString() : 'never';
}

/**
 * The caller's own agent API keys: a credential for an agent or app that is
 * narrower than their session. Minted here, shown once, revocable here.
 */
export default function AgentKeys({ userId, grantableServices, allowedSites }: Props) {
  const [keys, setKeys] = useState<AgentKey[] | null>(null);
  const [tenantPolicy, setTenantPolicy] = useState<TenantPolicy | null>(null);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [label, setLabel] = useState('');
  const [services, setServices] = useState<string[]>([]);
  const [readOnly, setReadOnly] = useState<string[]>([]);
  const [sitePicks, setSitePicks] = useState<string[]>([]);
  const [expiresInDays, setExpiresInDays] = useState<string>('90');
  const [wantReminder, setWantReminder] = useState(true);
  const [saving, setSaving] = useState(false);
  const [minted, setMinted] = useState<{ key: string; label: string; keyId: string; reminder?: ReminderResult } | null>(null);
  const [copied, setCopied] = useState(false);
  const [confirmRevoke, setConfirmRevoke] = useState<AgentKey | null>(null);

  const headers = { 'x-user-id': userId, 'Content-Type': 'application/json' };

  const flash = (kind: 'ok' | 'error', text: string) => {
    setStatusMsg({ kind, text });
    setTimeout(() => setStatusMsg(null), kind === 'ok' ? 2500 : 5000);
  };

  const load = useCallback(async () => {
    try {
      const [k, p] = await Promise.all([
        fetch('/api/manage/agent-keys', { credentials: 'same-origin', headers: { 'x-user-id': userId } }),
        fetch('/api/manage/agent-key-policy', { credentials: 'same-origin', headers: { 'x-user-id': userId } }),
      ]);
      if (k.ok) setKeys(((await k.json()) as { keys: AgentKey[] }).keys);
      else setKeys([]);
      if (p.ok) setTenantPolicy((await p.json()) as TenantPolicy);
    } catch {
      setKeys([]);
    }
  }, [userId]);

  useEffect(() => { load(); }, [load]);

  const grantable = SERVICE_ORDER.filter((s) =>
    grantableServices.includes(s) && !(tenantPolicy?.disabledServices ?? []).includes(s));
  const keysAllowed = tenantPolicy?.allowKeys !== false;
  const maxDays = Math.min(365, tenantPolicy?.maxKeyLifetimeDays ?? 365);
  const daysValue = Number(expiresInDays);
  const daysValid = Number.isInteger(daysValue) && daysValue >= 1 && daysValue <= maxDays;

  const toggleIn = (list: string[], v: string, on: boolean) => (on ? [...new Set([...list, v])] : list.filter((x) => x !== v));

  const resetForm = () => {
    setLabel(''); setServices([]); setReadOnly([]); setSitePicks([]); setExpiresInDays(String(Math.min(90, maxDays))); setWantReminder(true); setShowForm(false);
  };

  const handleCreate = async (e: React.FormEvent) => {
    e.preventDefault();
    setSaving(true);
    try {
      const body: Record<string, unknown> = {
        label: label.trim(),
        services,
        readOnlyServices: readOnly.filter((s) => services.includes(s)),
      };
      if (services.includes('sharepoint') && sitePicks.length > 0 && allowedSites) {
        body.allowedSites = allowedSites.filter((s) => sitePicks.includes(s.id));
      }
      body.expiresInDays = daysValue;
      body.reminder = wantReminder;
      const res = await fetch('/api/manage/agent-keys', { method: 'POST', credentials: 'same-origin', headers, body: JSON.stringify(body) });
      const data = (await res.json()) as { key?: string; keyId?: string; label?: string; error?: string; reminder?: ReminderResult };
      if (!res.ok || !data.key) throw new Error(data.error ?? `HTTP ${res.status}`);
      setMinted({ key: data.key, label: data.label ?? label, keyId: data.keyId ?? '', reminder: data.reminder });
      setCopied(false);
      resetForm();
      await load();
    } catch (err: unknown) {
      flash('error', err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const handleRevoke = async (k: AgentKey) => {
    setConfirmRevoke(null);
    try {
      const res = await fetch(`/api/manage/agent-keys?keyId=${encodeURIComponent(k.keyId)}`, { method: 'DELETE', credentials: 'same-origin', headers });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      flash('ok', `Revoked “${k.label}”.`);
      await load();
    } catch (err: unknown) {
      flash('error', err instanceof Error ? err.message : String(err));
    }
  };

  const copyKey = async () => {
    if (!minted) return;
    try {
      await navigator.clipboard.writeText(minted.key);
      setCopied(true);
    } catch {
      setCopied(false);
    }
  };

  const live = (keys ?? []).filter((k) => !k.revokedAt);
  const revoked = (keys ?? []).filter((k) => k.revokedAt);

  return (
    <>
      <p className="section-hint">
        An agent key lets an app or agent you run call this server as you, with less than you have.
        Pick the services it may use, which of those it may only read, and for SharePoint which sites it
        may see. The key can never reach anything you cannot, and your admin can tighten every key at once.
        Every key has an expiry you choose, up to {maxDays} days, and by default a reminder lands on your
        calendar a few days before it. Signing out everywhere ends your keys too, and each one stops working
        the moment you revoke it here. The key value is shown once.
      </p>

      {!keysAllowed && (
        <p className="error-msg">Your administrator has turned agent keys off for this organization. Existing keys stop working while this is set.</p>
      )}

      {minted && (
        <div className="confirm-dialog" role="dialog" aria-label="New agent key">
          <p><strong>Key “{minted.label}” created.</strong> Copy it now. It is not shown again.</p>
          {minted.reminder && 'eventId' in minted.reminder && (
            <p className="section-hint">A reminder is on your calendar a few days before it expires.</p>
          )}
          {minted.reminder && 'error' in minted.reminder && (
            <p className="error-msg">{minted.reminder.error} Note the expiry date yourself.</p>
          )}
          <pre style={{ userSelect: 'all', whiteSpace: 'pre-wrap', wordBreak: 'break-all', background: '#f6f6f6', padding: 8, borderRadius: 4, fontSize: 12 }}>{minted.key}</pre>
          <p className="section-hint" style={{ marginTop: 6 }}>
            Send it as <code>Authorization: Bearer &lt;key&gt;</code> to <code>/api/mcp</code>, or hand it to the
            shim with <code>--store-token</code> like a session token. Keep it in a secret store, never in a repo.
          </p>
          <div className="confirm-actions">
            <button type="button" className="confirm-btn" onClick={copyKey}>{copied ? 'Copied' : 'Copy key'}</button>
            <button type="button" className="confirm-btn" onClick={() => setMinted(null)}>Done</button>
          </div>
        </div>
      )}

      {keys === null ? (
        <p style={{ color: '#666', fontSize: 13 }}>Loading keys…</p>
      ) : live.length === 0 ? (
        <p className="empty-state">No agent keys.</p>
      ) : (
        <table className="deny-table">
          <thead>
            <tr><th>Label</th><th>Id</th><th>May use</th><th>Read-only</th><th>Sites</th><th>Last used</th><th>Expires</th><th>Reminder</th><th></th></tr>
          </thead>
          <tbody>
            {live.map((k) => (
              <tr key={k.keyId}>
                <td>{k.label}</td>
                <td><code>{k.keyId}</code></td>
                <td>{k.policy.services.map((s) => SERVICE_LABELS[s] ?? s).join(', ')}</td>
                <td>{k.policy.readOnlyServices.length ? k.policy.readOnlyServices.map((s) => SERVICE_LABELS[s] ?? s).join(', ') : '—'}</td>
                <td>{k.policy.allowedSites.length ? k.policy.allowedSites.map((s) => s.name).join(', ') : 'as tenant'}</td>
                <td>{when(k.lastUsedAt)}</td>
                <td>{when(k.expiresAt)}</td>
                <td>{k.reminderSet ? 'on calendar' : '—'}</td>
                <td><button type="button" className="remove-btn" onClick={() => setConfirmRevoke(k)}>Revoke</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {confirmRevoke && (
        <div className="confirm-dialog" role="dialog" aria-label="Revoke agent key">
          <p>Revoke “{confirmRevoke.label}”? Anything using it stops working within about half a minute.</p>
          <div className="confirm-actions">
            <button type="button" className="confirm-btn confirm-btn--danger" onClick={() => handleRevoke(confirmRevoke)}>Revoke</button>
            <button type="button" className="confirm-btn" onClick={() => setConfirmRevoke(null)}>Cancel</button>
          </div>
        </div>
      )}

      {revoked.length > 0 && (
        <details style={{ marginTop: 10 }}>
          <summary style={{ fontSize: 13, color: '#666', cursor: 'pointer' }}>{revoked.length} revoked</summary>
          <table className="deny-table" style={{ marginTop: 6 }}>
            <thead><tr><th>Label</th><th>Id</th><th>Revoked</th></tr></thead>
            <tbody>
              {revoked.map((k) => (
                <tr key={k.keyId}><td>{k.label}</td><td><code>{k.keyId}</code></td><td>{when(k.revokedAt)}</td></tr>
              ))}
            </tbody>
          </table>
        </details>
      )}

      {keysAllowed && !showForm && (
        <p style={{ marginTop: 12 }}>
          <button type="button" className="confirm-btn" onClick={() => setShowForm(true)} disabled={grantable.length === 0}>New agent key</button>
          {grantable.length === 0 && <span className="section-hint" style={{ marginLeft: 8 }}>No service is available to grant.</span>}
        </p>
      )}

      {keysAllowed && showForm && (
        <form onSubmit={handleCreate} className="card" style={{ marginTop: 12 }}>
          <label style={{ display: 'block', marginBottom: 10 }}>
            <span className="service-name">Label</span><br />
            <input type="text" value={label} onChange={(e) => setLabel(e.target.value)} maxLength={64} required placeholder="e.g. nightly-report" style={{ width: '100%', maxWidth: 360 }} />
          </label>

          <div className="service-name" style={{ marginBottom: 4 }}>Services the key may use</div>
          <div className="services-list">
            {grantable.map((s) => (
              <div key={s} className="service-row">
                <label className="service-toggle">
                  <input type="checkbox" checked={services.includes(s)} onChange={(e) => {
                    setServices(toggleIn(services, s, e.target.checked));
                    if (!e.target.checked) setReadOnly(readOnly.filter((x) => x !== s));
                  }} />
                  <span className="service-info"><span className="service-name">{SERVICE_LABELS[s] ?? s}</span></span>
                </label>
                {services.includes(s) && (
                  <label className="service-toggle" style={{ marginLeft: 'auto' }}>
                    <input type="checkbox" checked={readOnly.includes(s) || (tenantPolicy?.readOnlyServices ?? []).includes(s)}
                      disabled={(tenantPolicy?.readOnlyServices ?? []).includes(s)}
                      onChange={(e) => setReadOnly(toggleIn(readOnly, s, e.target.checked))} />
                    <span className="service-desc">read-only{(tenantPolicy?.readOnlyServices ?? []).includes(s) ? ' (set by admin)' : ''}</span>
                  </label>
                )}
              </div>
            ))}
          </div>

          {services.includes('sharepoint') && allowedSites && allowedSites.length > 0 && (
            <>
              <div className="service-name" style={{ margin: '10px 0 4px' }}>SharePoint sites (leave empty for all the tenant allows)</div>
              <div className="services-list">
                {allowedSites.map((site) => (
                  <div key={site.id} className="service-row">
                    <label className="service-toggle">
                      <input type="checkbox" checked={sitePicks.includes(site.id)} onChange={(e) => setSitePicks(toggleIn(sitePicks, site.id, e.target.checked))} />
                      <span className="service-info"><span className="service-name">{site.name}</span></span>
                    </label>
                  </div>
                ))}
              </div>
            </>
          )}

          <label style={{ display: 'block', margin: '10px 0' }}>
            <span className="service-name">Expires in days</span>{' '}
            <input type="number" min={1} max={maxDays} required value={expiresInDays} onChange={(e) => setExpiresInDays(e.target.value)} style={{ width: 80 }} />
            <span className="service-desc" style={{ marginLeft: 8 }}>1 to {maxDays}. {maxDays < 365 ? 'The maximum is set by your admin.' : 'Up to a year.'}</span>
          </label>

          <label className="service-toggle" style={{ display: 'flex', margin: '0 0 10px' }}>
            <input type="checkbox" checked={wantReminder} onChange={(e) => setWantReminder(e.target.checked)} />
            <span className="service-info">
              <span className="service-name">Calendar reminder</span>
              <span className="service-desc">Put a short event on my calendar before the key expires (a week before for keys of two weeks or more, two days before for shorter ones).</span>
            </span>
          </label>

          <div className="confirm-actions">
            <button type="submit" className="confirm-btn" disabled={saving || !label.trim() || services.length === 0 || !daysValid}>{saving ? 'Creating…' : 'Create key'}</button>
            <button type="button" className="confirm-btn" onClick={resetForm} disabled={saving}>Cancel</button>
          </div>
        </form>
      )}

      {statusMsg && <p className={`status-msg status-msg--${statusMsg.kind}`}>{statusMsg.text}</p>}
    </>
  );
}
