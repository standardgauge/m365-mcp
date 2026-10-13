import React, { useCallback, useEffect, useState } from 'react';
import KeyScopeFields, { isNarrowingOf, serviceNames } from './KeyScopeFields';
import AgentKeyDenyList from './AgentKeyDenyList';
import type { KeyPolicy, SitePick } from './KeyScopeFields';

interface AgentKey {
  keyId: string; label: string; userId: string; userEmail: string; policy: KeyPolicy;
  createdAt: number; expiresAt: number; lastUsedAt: number | null; revokedAt: number | null; revokedBy: string | null;
  reminderSet?: boolean; policyUpdatedAt?: number | null;
}
interface TenantPolicy { allowKeys: boolean; disabledServices: string[]; readOnlyServices: string[]; maxKeyLifetimeDays: number; ownersMayWiden: boolean }
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
const EMPTY: KeyPolicy = { services: [], readOnlyServices: [], allowedSites: [] };

/**
 * The caller's own agent API keys: a credential for an agent or app that is
 * narrower than their session. Minted, edited and revoked here; shown once.
 */
export default function AgentKeys({ userId, grantableServices, allowedSites }: Props) {
  const [keys, setKeys] = useState<AgentKey[] | null>(null);
  const [tenantPolicy, setTenantPolicy] = useState<TenantPolicy | null>(null);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);
  const [showForm, setShowForm] = useState(false);
  const [label, setLabel] = useState('');
  const [scope, setScope] = useState<KeyPolicy>(EMPTY);
  const [expiresInDays, setExpiresInDays] = useState<string>('90');
  const [wantReminder, setWantReminder] = useState(true);
  const [saving, setSaving] = useState(false);
  const [minted, setMinted] = useState<{ key: string; label: string; keyId: string; reminder?: ReminderResult } | null>(null);
  const [copied, setCopied] = useState(false);
  const [confirmRevoke, setConfirmRevoke] = useState<AgentKey | null>(null);
  const [editing, setEditing] = useState<AgentKey | null>(null);
  const [editLabel, setEditLabel] = useState('');
  const [editScope, setEditScope] = useState<KeyPolicy>(EMPTY);
  const [blocking, setBlocking] = useState<AgentKey | null>(null);

  const headers = { 'x-user-id': userId, 'Content-Type': 'application/json' };

  const flash = (kind: 'ok' | 'error', text: string) => {
    setStatusMsg({ kind, text });
    setTimeout(() => setStatusMsg(null), kind === 'ok' ? 2500 : 6000);
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

  const grantable = grantableServices.filter((s) => !(tenantPolicy?.disabledServices ?? []).includes(s));
  const tenantReadOnly = tenantPolicy?.readOnlyServices ?? [];
  const keysAllowed = tenantPolicy?.allowKeys !== false;
  const ownersMayWiden = tenantPolicy?.ownersMayWiden !== false;
  const maxDays = Math.min(365, tenantPolicy?.maxKeyLifetimeDays ?? 365);
  const daysValue = Number(expiresInDays);
  const daysValid = Number.isInteger(daysValue) && daysValue >= 1 && daysValue <= maxDays;

  const resetForm = () => {
    setLabel(''); setScope(EMPTY); setExpiresInDays(String(Math.min(90, maxDays))); setWantReminder(true); setShowForm(false);
  };

  const parseError = async (res: Response) => {
    const data = (await res.json().catch(() => ({}))) as { error?: string };
    return data.error ?? `HTTP ${res.status}`;
  };

  const handleCreate = async (e: React.FormEvent) => {
    e.preventDefault();
    setSaving(true);
    try {
      const body: Record<string, unknown> = {
        label: label.trim(),
        services: scope.services,
        readOnlyServices: scope.readOnlyServices.filter((s) => scope.services.includes(s)),
        expiresInDays: daysValue,
        reminder: wantReminder,
      };
      if (scope.services.includes('sharepoint') && scope.allowedSites.length > 0) body.allowedSites = scope.allowedSites;
      const res = await fetch('/api/manage/agent-keys', { method: 'POST', credentials: 'same-origin', headers, body: JSON.stringify(body) });
      if (!res.ok) throw new Error(await parseError(res));
      const data = (await res.json()) as { key: string; keyId: string; label: string; reminder?: ReminderResult };
      setMinted({ key: data.key, label: data.label, keyId: data.keyId, reminder: data.reminder });
      setCopied(false);
      resetForm();
      await load();
    } catch (err: unknown) {
      flash('error', err instanceof Error ? err.message : String(err));
    } finally {
      setSaving(false);
    }
  };

  const startEdit = (k: AgentKey) => {
    setBlocking(null);
    setEditing(k);
    setEditLabel(k.label);
    setEditScope({ services: [...k.policy.services], readOnlyServices: [...k.policy.readOnlyServices], allowedSites: k.policy.allowedSites.map((s) => ({ ...s })) });
  };

  const editWidens = editing ? !isNarrowingOf(editScope, editing.policy) : false;
  const editBlocked = editWidens && !ownersMayWiden;

  const handleEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editing) return;
    setSaving(true);
    try {
      const res = await fetch(`/api/manage/agent-keys?keyId=${encodeURIComponent(editing.keyId)}`, {
        method: 'PATCH', credentials: 'same-origin', headers,
        body: JSON.stringify({
          label: editLabel.trim(),
          services: editScope.services,
          readOnlyServices: editScope.readOnlyServices.filter((s) => editScope.services.includes(s)),
          allowedSites: editScope.services.includes('sharepoint') ? editScope.allowedSites : [],
        }),
      });
      if (!res.ok) throw new Error(await parseError(res));
      const data = (await res.json()) as { widened?: boolean };
      flash('ok', data.widened ? `“${editLabel.trim()}” widened. The app holding it has the new access from its next call.` : `“${editLabel.trim()}” updated.`);
      setEditing(null);
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
      if (!res.ok) throw new Error(await parseError(res));
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
        may see. Under <strong>Folders</strong> you can also keep a key out of a folder, mail folder or calendar
        you can still reach yourself. The key can never reach anything you cannot, and your admin can tighten every key at once.
        Every key has an expiry you choose, up to {maxDays} days, and by default a reminder lands on your
        calendar a few days before it. Signing out everywhere ends your keys too, and each one stops working
        the moment you revoke it here. The key value is shown once.
      </p>
      <p className="section-hint">
        <strong>One key per app.</strong> Keys are cheap to mint and easy to edit, so give each app or agent its own.
        A key shared between two apps cannot be narrowed or revoked for just one of them, and its audit trail
        cannot tell them apart.
        {' '}
        {ownersMayWiden
          ? <>You can <strong>widen or narrow</strong> a key here after minting it. The app holding it gets the change on its next call, without touching the key itself, and every change is recorded with before and after.</>
          : <>Your organization allows a key to be <strong>narrowed only</strong> after minting. To give an app more access, mint a new key and retire the old one.</>}
      </p>

      {!keysAllowed && (
        <p className="error-msg">Your administrator has turned agent keys off for this organization. Existing keys stop working while this is set.</p>
      )}

      {minted && (
        <div className="confirm-dialog" role="dialog" aria-label="New agent key">
          <p><strong>Key “{minted.label}” created.</strong> Copy it now. It is not shown again.</p>
          <pre style={{ userSelect: 'all', whiteSpace: 'pre-wrap', wordBreak: 'break-all', background: '#f6f6f6', padding: 8, borderRadius: 4, fontSize: 12 }}>{minted.key}</pre>
          {minted.reminder && 'eventId' in minted.reminder && (
            <p className="section-hint">A reminder is on your calendar a few days before it expires.</p>
          )}
          {minted.reminder && 'error' in minted.reminder && (
            <p className="error-msg">{minted.reminder.error} Note the expiry date yourself.</p>
          )}
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
                <td>{k.label}{k.policyUpdatedAt ? <span className="service-desc"><br />edited {when(k.policyUpdatedAt)}</span> : null}</td>
                <td><code>{k.keyId}</code></td>
                <td>{serviceNames(k.policy.services)}</td>
                <td>{serviceNames(k.policy.readOnlyServices)}</td>
                <td>{k.policy.allowedSites.length ? k.policy.allowedSites.map((s) => s.name).join(', ') : 'as tenant'}</td>
                <td>{when(k.lastUsedAt)}</td>
                <td>{when(k.expiresAt)}</td>
                <td>{k.reminderSet ? 'on calendar' : '—'}</td>
                <td style={{ whiteSpace: 'nowrap' }}>
                  <button type="button" className="confirm-btn" style={{ marginRight: 6 }} onClick={() => startEdit(k)}>Edit</button>
                  <button type="button" className="confirm-btn" style={{ marginRight: 6 }} onClick={() => { setEditing(null); setBlocking(k); }}>Folders</button>
                  <button type="button" className="remove-btn" onClick={() => setConfirmRevoke(k)}>Revoke</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {blocking && !editing && (
        <AgentKeyDenyList
          key={blocking.keyId}
          userId={userId}
          keyId={blocking.keyId}
          label={blocking.label}
          services={blocking.policy.services}
          sites={blocking.policy.allowedSites.length ? blocking.policy.allowedSites : allowedSites}
          mayRemove={ownersMayWiden}
          onClose={() => setBlocking(null)}
        />
      )}

      {editing && (
        <form onSubmit={handleEdit} className="card" style={{ marginTop: 12 }}>
          <p><strong>Edit “{editing.label}”</strong> <code>{editing.keyId}</code></p>
          <p className="section-hint">
            {ownersMayWiden
              ? <>You can add or remove services, change read-only, and change sites. <strong>Widening takes effect on the key's next call</strong>, wherever the key is deployed, so if this key is somewhere you no longer trust, revoke it and mint a new one instead of widening it. Every change is recorded with before and after.</>
              : <>Your organization allows a key to be <strong>narrowed only</strong>: you can remove services, make a service read-only, or remove sites. Anything wider needs a new key. Every change is recorded with before and after.</>}
            {' '}The key value and its expiry do not change.
          </p>
          <label style={{ display: 'block', marginBottom: 10 }}>
            <span className="service-name">Label</span><br />
            <input type="text" value={editLabel} onChange={(e) => setEditLabel(e.target.value)} maxLength={64} required style={{ width: '100%', maxWidth: 360 }} />
          </label>
          <KeyScopeFields
            value={editScope}
            onChange={setEditScope}
            grantable={grantable}
            tenantReadOnly={tenantReadOnly}
            siteOptions={allowedSites}
            narrowOnlyFrom={ownersMayWiden ? undefined : editing.policy}
            disabled={saving}
          />
          {editWidens && ownersMayWiden && (
            <p className="status-msg status-msg--error" style={{ marginTop: 8 }}>
              This change <strong>widens</strong> the key. Whatever holds it gains the new access on its next call.
            </p>
          )}
          <div className="confirm-actions">
            <button type="submit" className="confirm-btn" disabled={saving || !editLabel.trim() || editScope.services.length === 0 || editBlocked}>{saving ? 'Saving…' : editWidens ? 'Save and widen' : 'Save'}</button>
            <button type="button" className="confirm-btn" onClick={() => setEditing(null)} disabled={saving}>Cancel</button>
          </div>
        </form>
      )}

      {confirmRevoke && (
        <div className="confirm-dialog" role="dialog" aria-label="Revoke agent key">
          <p>Revoke “{confirmRevoke.label}”? Anything using it stops working within about half a minute. Its calendar reminder is removed.</p>
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

      {keysAllowed && !showForm && !editing && (
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
            <span className="service-desc"> Name the app or agent that will hold it.</span>
          </label>

          <KeyScopeFields value={scope} onChange={setScope} grantable={grantable} tenantReadOnly={tenantReadOnly} siteOptions={allowedSites} disabled={saving} />

          <label style={{ display: 'block', margin: '10px 0' }}>
            <span className="service-name">Expires in days</span>{' '}
            <input type="number" min={1} max={maxDays} required value={expiresInDays} onChange={(e) => setExpiresInDays(e.target.value)} style={{ width: 80 }} />
            <span className="service-desc" style={{ marginLeft: 8 }}>1 to {maxDays}. {maxDays < 365 ? 'The maximum is set by your admin.' : 'Up to a year.'} The expiry cannot be changed later; mint a new key when it is near.</span>
          </label>

          <label className="service-toggle" style={{ display: 'flex', margin: '0 0 10px' }}>
            <input type="checkbox" checked={wantReminder} onChange={(e) => setWantReminder(e.target.checked)} />
            <span className="service-info">
              <span className="service-name">Calendar reminder</span>
              <span className="service-desc">Put a short event on my calendar before the key expires (a week before for keys of two weeks or more, two days before for shorter ones).</span>
            </span>
          </label>

          <div className="confirm-actions">
            <button type="submit" className="confirm-btn" disabled={saving || !label.trim() || scope.services.length === 0 || !daysValid}>{saving ? 'Creating…' : 'Create key'}</button>
            <button type="button" className="confirm-btn" onClick={resetForm} disabled={saving}>Cancel</button>
          </div>
        </form>
      )}

      {statusMsg && <p className={`status-msg status-msg--${statusMsg.kind}`}>{statusMsg.text}</p>}
    </>
  );
}
