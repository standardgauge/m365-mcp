import React, { useCallback, useEffect, useState } from 'react';

const SERVICE_LABELS: Record<string, string> = {
  mail: 'Mail', sharepoint: 'SharePoint', calendar: 'Calendar', onedrive: 'OneDrive',
  onenote: 'OneNote', contacts: 'Contacts', teams: 'Teams',
};
const SERVICE_ORDER = ['mail', 'sharepoint', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'];

interface TenantPolicy { allowKeys: boolean; disabledServices: string[]; readOnlyServices: string[]; maxKeyLifetimeDays: number }
interface AgentKey {
  keyId: string; label: string; userId: string; userEmail: string;
  policy: { services: string[]; readOnlyServices: string[]; allowedSites: Array<{ id: string; name: string }> };
  createdAt: number; expiresAt: number | null; lastUsedAt: number | null; revokedAt: number | null;
}

interface Props { userId: string; enabledServices: string[] }

/**
 * Tenant-wide agent key policy and the list of every key in the tenant.
 * Global Administrator only (the routes refuse anyone else).
 */
export default function AgentKeyPolicy({ userId, enabledServices }: Props) {
  const [policy, setPolicy] = useState<TenantPolicy | null>(null);
  const [keys, setKeys] = useState<AgentKey[] | null>(null);
  const [saving, setSaving] = useState(false);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);
  const [confirmRevoke, setConfirmRevoke] = useState<AgentKey | null>(null);
  const [maxDaysDraft, setMaxDaysDraft] = useState<string>('');

  const headers = { 'x-user-id': userId, 'Content-Type': 'application/json' };
  const flash = (kind: 'ok' | 'error', text: string) => {
    setStatusMsg({ kind, text });
    setTimeout(() => setStatusMsg(null), kind === 'ok' ? 2500 : 5000);
  };

  const load = useCallback(async () => {
    try {
      const [p, k] = await Promise.all([
        fetch('/api/manage/agent-key-policy', { credentials: 'same-origin', headers: { 'x-user-id': userId } }),
        fetch('/api/manage/agent-keys?all=true', { credentials: 'same-origin', headers: { 'x-user-id': userId } }),
      ]);
      if (p.ok) {
        const tp = (await p.json()) as TenantPolicy;
        setPolicy(tp);
        setMaxDaysDraft(String(tp.maxKeyLifetimeDays));
      }
      if (k.ok) setKeys(((await k.json()) as { keys: AgentKey[] }).keys); else setKeys([]);
    } catch (err: unknown) {
      flash('error', err instanceof Error ? err.message : String(err));
    }
  }, [userId]);

  useEffect(() => { load(); }, [load]);

  const save = async (patch: Partial<TenantPolicy>) => {
    setSaving(true);
    try {
      const res = await fetch('/api/manage/agent-key-policy', { method: 'POST', credentials: 'same-origin', headers, body: JSON.stringify(patch) });
      const data = (await res.json()) as TenantPolicy & { error?: string };
      if (!res.ok) throw new Error(data.error ?? `HTTP ${res.status}`);
      setPolicy({ allowKeys: data.allowKeys, disabledServices: data.disabledServices, readOnlyServices: data.readOnlyServices, maxKeyLifetimeDays: data.maxKeyLifetimeDays });
      setMaxDaysDraft(String(data.maxKeyLifetimeDays));
      flash('ok', 'Saved.');
    } catch (err: unknown) {
      flash('error', err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const toggleIn = (list: string[], v: string, on: boolean) => (on ? [...new Set([...list, v])] : list.filter((x) => x !== v));

  const handleRevoke = async (k: AgentKey) => {
    setConfirmRevoke(null);
    try {
      const res = await fetch(`/api/manage/agent-keys?keyId=${encodeURIComponent(k.keyId)}`, { method: 'DELETE', credentials: 'same-origin', headers });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      flash('ok', `Revoked “${k.label}” (${k.userEmail}).`);
      await load();
    } catch (err: unknown) {
      flash('error', err instanceof Error ? err.message : String(err));
    }
  };

  const live = (keys ?? []).filter((k) => !k.revokedAt);

  return (
    <>
      <p className="section-hint">
        Users mint agent keys under My Settings for apps and agents they run. A key is bound to the
        user's own delegated session and can only be narrower than it. The settings here apply to
        every key in the organization, existing ones included, on the next request: turn keys off
        entirely, take services away from all keys, make a service read-only for all keys even
        where it is writable for the user, or cap how long a key may live. A key's lifetime is its
        own, chosen by the owner within that cap, and is not tied to the 30-day sign-in cap.
      </p>
      {policy === null ? (
        <p style={{ color: '#666', fontSize: 13 }}>Loading policy…</p>
      ) : (
        <>
          <div className="services-list">
            <div className="service-row">
              <label className="service-toggle">
                <input type="checkbox" checked={policy.allowKeys} disabled={saving} onChange={(e) => save({ allowKeys: e.target.checked })} />
                <span className="service-info">
                  <span className="service-name">Allow agent keys</span>
                  <span className="service-desc">{policy.allowKeys ? 'Users can mint keys and existing keys work.' : 'No key authenticates and none can be minted.'}</span>
                </span>
              </label>
              {!policy.allowKeys && <span className="override-badge">Off</span>}
            </div>
          </div>

          <div className="service-name" style={{ margin: '12px 0 4px' }}>Longest key lifetime</div>
          <p style={{ fontSize: 13, margin: '0 0 8px' }}>
            <input type="number" min={1} max={365} value={maxDaysDraft} disabled={saving}
              onChange={(e) => setMaxDaysDraft(e.target.value)} style={{ width: 80 }} />{' '}
            days (1 to 365){' '}
            <button type="button" className="confirm-btn" disabled={saving || Number(maxDaysDraft) === policy.maxKeyLifetimeDays || !Number.isInteger(Number(maxDaysDraft)) || Number(maxDaysDraft) < 1 || Number(maxDaysDraft) > 365}
              onClick={() => save({ maxKeyLifetimeDays: Number(maxDaysDraft) })}>Save</button>
            <span className="service-desc" style={{ marginLeft: 8 }}>Applies to keys minted from now on; existing keys keep the lifetime they were given.</span>
          </p>

          <div className="service-name" style={{ margin: '12px 0 4px' }}>Per service, for every key</div>
          <table className="deny-table">
            <thead><tr><th>Service</th><th>Keys may use</th><th>Read-only for keys</th></tr></thead>
            <tbody>
              {SERVICE_ORDER.map((s) => {
                const tenantOn = enabledServices.includes(s);
                const allowed = !policy.disabledServices.includes(s);
                return (
                  <tr key={s} className={tenantOn ? undefined : 'service-row--tenant-disabled'}>
                    <td>{SERVICE_LABELS[s] ?? s}{!tenantOn && <span className="service-desc"> (not enabled for the tenant)</span>}</td>
                    <td>
                      <input type="checkbox" checked={allowed} disabled={saving}
                        onChange={(e) => save({ disabledServices: toggleIn(policy.disabledServices, s, !e.target.checked) })} />
                    </td>
                    <td>
                      <input type="checkbox" checked={policy.readOnlyServices.includes(s)} disabled={saving || !allowed}
                        onChange={(e) => save({ readOnlyServices: toggleIn(policy.readOnlyServices, s, e.target.checked) })} />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </>
      )}

      <div className="service-name" style={{ margin: '14px 0 4px' }}>Keys in this organization</div>
      {keys === null ? (
        <p style={{ color: '#666', fontSize: 13 }}>Loading keys…</p>
      ) : live.length === 0 ? (
        <p className="empty-state">No live agent keys.</p>
      ) : (
        <table className="deny-table">
          <thead><tr><th>Owner</th><th>Label</th><th>Id</th><th>May use</th><th>Read-only</th><th>Last used</th><th>Expires</th><th></th></tr></thead>
          <tbody>
            {live.map((k) => (
              <tr key={k.keyId}>
                <td>{k.userEmail}</td>
                <td>{k.label}</td>
                <td><code>{k.keyId}</code></td>
                <td>{k.policy.services.map((s) => SERVICE_LABELS[s] ?? s).join(', ')}</td>
                <td>{k.policy.readOnlyServices.length ? k.policy.readOnlyServices.map((s) => SERVICE_LABELS[s] ?? s).join(', ') : '—'}</td>
                <td>{k.lastUsedAt ? new Date(k.lastUsedAt).toLocaleString() : 'never'}</td>
                <td>{k.expiresAt ? new Date(k.expiresAt).toLocaleString() : '—'}</td>
                <td><button type="button" className="remove-btn" onClick={() => setConfirmRevoke(k)}>Revoke</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {confirmRevoke && (
        <div className="confirm-dialog" role="dialog" aria-label="Revoke agent key">
          <p>Revoke “{confirmRevoke.label}” belonging to {confirmRevoke.userEmail}? Anything using it stops working within about half a minute.</p>
          <div className="confirm-actions">
            <button type="button" className="confirm-btn confirm-btn--danger" onClick={() => handleRevoke(confirmRevoke)}>Revoke</button>
            <button type="button" className="confirm-btn" onClick={() => setConfirmRevoke(null)}>Cancel</button>
          </div>
        </div>
      )}

      {statusMsg && <p className={`status-msg status-msg--${statusMsg.kind}`}>{statusMsg.text}</p>}
    </>
  );
}
