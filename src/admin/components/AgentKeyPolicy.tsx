import React, { useCallback, useEffect, useState } from 'react';
import KeyScopeFields, { SERVICE_LABELS, SERVICE_ORDER, isNarrowingOf, serviceNames } from './KeyScopeFields';
import type { KeyPolicy, SitePick } from './KeyScopeFields';

interface TenantPolicy { allowKeys: boolean; disabledServices: string[]; readOnlyServices: string[]; maxKeyLifetimeDays: number; ownersMayWiden: boolean }
interface AgentKey {
  keyId: string; label: string; userId: string; userEmail: string; policy: KeyPolicy;
  createdAt: number; expiresAt: number; lastUsedAt: number | null; revokedAt: number | null; revokedBy: string | null;
  reminderSet?: boolean; policyUpdatedAt?: number | null; policyUpdatedBy?: string | null;
}

interface Props { userId: string; enabledServices: string[]; allowedSites?: SitePick[] }

function when(ms: number | null | undefined): string {
  return ms ? new Date(ms).toLocaleString() : 'never';
}

/**
 * Tenant-wide agent key policy and every key in the tenant. Global
 * Administrator only (the routes refuse anyone else). An administrator can
 * revoke any key and narrow any key; widening is the owner's decision.
 */
export default function AgentKeyPolicy({ userId, enabledServices, allowedSites }: Props) {
  const [policy, setPolicy] = useState<TenantPolicy | null>(null);
  const [keys, setKeys] = useState<AgentKey[] | null>(null);
  const [saving, setSaving] = useState(false);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);
  const [confirmRevoke, setConfirmRevoke] = useState<AgentKey | null>(null);
  const [maxDaysDraft, setMaxDaysDraft] = useState<string>('');
  const [showRevoked, setShowRevoked] = useState(false);
  const [editing, setEditing] = useState<AgentKey | null>(null);
  const [editScope, setEditScope] = useState<KeyPolicy>({ services: [], readOnlyServices: [], allowedSites: [] });

  const headers = { 'x-user-id': userId, 'Content-Type': 'application/json' };
  const flash = (kind: 'ok' | 'error', text: string) => {
    setStatusMsg({ kind, text });
    setTimeout(() => setStatusMsg(null), kind === 'ok' ? 2500 : 6000);
  };
  const parseError = async (res: Response) => {
    const data = (await res.json().catch(() => ({}))) as { error?: string };
    return data.error ?? `HTTP ${res.status}`;
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

  useEffect(() => { void load(); }, [load]);

  const save = async (patch: Partial<TenantPolicy>) => {
    setSaving(true);
    try {
      const res = await fetch('/api/manage/agent-key-policy', { method: 'POST', credentials: 'same-origin', headers, body: JSON.stringify(patch) });
      if (!res.ok) throw new Error(await parseError(res));
      const data = (await res.json()) as TenantPolicy;
      setPolicy(data);
      setMaxDaysDraft(String(data.maxKeyLifetimeDays));
      flash('ok', 'Saved. Applies to every key on its next call.');
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
      if (!res.ok) throw new Error(await parseError(res));
      flash('ok', `Revoked “${k.label}” (${k.userEmail}).`);
      await load();
    } catch (err: unknown) {
      flash('error', err instanceof Error ? err.message : String(err));
    }
  };

  const startEdit = (k: AgentKey) => {
    setEditing(k);
    setEditScope({ services: [...k.policy.services], readOnlyServices: [...k.policy.readOnlyServices], allowedSites: k.policy.allowedSites.map((s) => ({ ...s })) });
  };

  const handleEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editing) return;
    setSaving(true);
    try {
      const res = await fetch(`/api/manage/agent-keys?keyId=${encodeURIComponent(editing.keyId)}`, {
        method: 'PATCH', credentials: 'same-origin', headers,
        body: JSON.stringify({
          services: editScope.services,
          readOnlyServices: editScope.readOnlyServices.filter((s) => editScope.services.includes(s)),
          allowedSites: editScope.services.includes('sharepoint') ? editScope.allowedSites : [],
        }),
      });
      if (!res.ok) throw new Error(await parseError(res));
      flash('ok', `Narrowed “${editing.label}” (${editing.userEmail}). The owner's app has the new limits from its next call.`);
      setEditing(null);
      await load();
    } catch (err: unknown) {
      flash('error', err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  // An administrator may only narrow; the server refuses anything else, and so does the form.
  const editWidens = editing ? !isNarrowingOf(editScope, editing.policy) : false;

  const live = (keys ?? []).filter((k) => !k.revokedAt);
  const revoked = (keys ?? []).filter((k) => k.revokedAt);
  const allReadOnly = policy ? SERVICE_ORDER.every((s) => policy.readOnlyServices.includes(s) || policy.disabledServices.includes(s)) : false;

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
      <p className="section-hint">
        <strong>Who may change a key after it is minted.</strong> Owners edit their own keys. You can
        <strong> narrow</strong> any key here (remove a service, make it read-only, remove a site) and
        <strong> revoke</strong> any key. You cannot widen another person's key: a key acts on its owner's
        data, so widening is theirs to decide. Whether owners may widen their own keys at all is the
        switch below. Every edit is recorded in the audit log with before and after.
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
                  <span className="service-desc">{policy.allowKeys ? 'Users can mint keys and existing keys work.' : 'No key authenticates and none can be minted. Existing keys are kept and resume if this is turned back on.'}</span>
                </span>
              </label>
              {!policy.allowKeys && <span className="override-badge">Off</span>}
            </div>
            <div className="service-row">
              <label className="service-toggle">
                <input type="checkbox" checked={policy.ownersMayWiden} disabled={saving} onChange={(e) => save({ ownersMayWiden: e.target.checked })} />
                <span className="service-info">
                  <span className="service-name">Owners may widen their own keys</span>
                  <span className="service-desc">
                    {policy.ownersMayWiden
                      ? 'On: an owner can add a service, clear read-only, or add sites to a key they already minted. The app holding the key gains the access on its next call without the key changing hands. Each widening is audited.'
                      : 'Off: an owner can only narrow a key after minting it. Wider access means a new key and retiring the old one, so every grant is a fresh, deliberate issue. Administrators are narrow-only either way.'}
                  </span>
                </span>
              </label>
              {!policy.ownersMayWiden && <span className="override-badge">Narrow-only</span>}
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
          <p className="section-hint" style={{ marginTop: 0 }}>
            These override whatever an owner chose or chooses later. Untick "Keys may use" to keep a service out of
            every key; tick "Read-only for keys" to refuse writes from every key even where the owner may write.
            {allReadOnly ? ' Every service is currently read-only or unavailable for keys, so no key in this organization can write anything.' : ' To make all keys read-only, tick the read-only column on every service.'}
          </p>
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

      <div className="service-name" style={{ margin: '14px 0 4px' }}>
        Keys in this organization
        {revoked.length > 0 && (
          <button type="button" className="confirm-btn" style={{ marginLeft: 10, fontSize: 12, padding: '2px 8px' }} onClick={() => setShowRevoked((v) => !v)}>
            {showRevoked ? 'Hide revoked' : `Show ${revoked.length} revoked`}
          </button>
        )}
      </div>
      {keys === null ? (
        <p style={{ color: '#666', fontSize: 13 }}>Loading keys…</p>
      ) : live.length === 0 && !showRevoked ? (
        <p className="empty-state">No live agent keys.</p>
      ) : (
        <table className="deny-table">
          <thead><tr><th>Owner</th><th>Label</th><th>Id</th><th>May use</th><th>Read-only</th><th>Sites</th><th>Last used</th><th>Expires</th><th></th></tr></thead>
          <tbody>
            {(showRevoked ? [...live, ...revoked] : live).map((k) => (
              <tr key={k.keyId} style={k.revokedAt ? { color: '#888' } : undefined}>
                <td>{k.userEmail}</td>
                <td>{k.label}{k.revokedAt ? <span className="service-desc"><br />revoked {when(k.revokedAt)}</span> : k.policyUpdatedAt ? <span className="service-desc"><br />edited {when(k.policyUpdatedAt)}</span> : null}</td>
                <td><code>{k.keyId}</code></td>
                <td>{serviceNames(k.policy.services)}</td>
                <td>{serviceNames(k.policy.readOnlyServices)}</td>
                <td>{k.policy.allowedSites.length ? k.policy.allowedSites.map((s) => s.name).join(', ') : 'as tenant'}</td>
                <td>{when(k.lastUsedAt)}</td>
                <td>{when(k.expiresAt)}</td>
                <td style={{ whiteSpace: 'nowrap' }}>
                  {!k.revokedAt && (
                    <>
                      <button type="button" className="confirm-btn" style={{ marginRight: 6 }} onClick={() => startEdit(k)}>Narrow</button>
                      <button type="button" className="remove-btn" onClick={() => setConfirmRevoke(k)}>Revoke</button>
                    </>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {editing && (
        <form onSubmit={handleEdit} className="card" style={{ marginTop: 12 }}>
          <p><strong>Narrow “{editing.label}”</strong> belonging to {editing.userEmail} <code>{editing.keyId}</code></p>
          <p className="section-hint">
            As an administrator you can <strong>only narrow</strong> this key: remove a service, make a service read-only,
            or remove a site. Controls that would widen it are locked and say so. The change takes effect on the key's
            next call and is recorded with before and after; the owner sees it as "edited" in their list. If the key
            needs more access, the owner mints or widens it themselves.
          </p>
          <KeyScopeFields
            value={editScope}
            onChange={setEditScope}
            grantable={editing.policy.services}
            tenantReadOnly={policy?.readOnlyServices ?? []}
            siteOptions={editing.policy.allowedSites.length ? editing.policy.allowedSites : allowedSites}
            narrowOnlyFrom={editing.policy}
            disabled={saving}
          />
          <div className="confirm-actions">
            <button type="submit" className="confirm-btn" disabled={saving || editScope.services.length === 0 || editWidens}>{saving ? 'Saving…' : 'Save'}</button>
            <button type="button" className="confirm-btn" onClick={() => setEditing(null)} disabled={saving}>Cancel</button>
          </div>
        </form>
      )}

      {confirmRevoke && (
        <div className="confirm-dialog" role="dialog" aria-label="Revoke agent key">
          <p>Revoke “{confirmRevoke.label}” belonging to {confirmRevoke.userEmail}? Anything using it stops working within about half a minute. Their calendar reminder for it is removed.</p>
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
