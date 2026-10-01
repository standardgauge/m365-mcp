import React, { useCallback, useEffect, useMemo, useState } from 'react';

const ALL_SERVICES: { key: string; label: string }[] = [
  { key: 'mail', label: 'Mail' },
  { key: 'sharepoint', label: 'SharePoint' },
  { key: 'calendar', label: 'Calendar' },
  { key: 'onedrive', label: 'OneDrive' },
  { key: 'onenote', label: 'OneNote' },
  { key: 'contacts', label: 'Contacts' },
  { key: 'teams', label: 'Teams' },
];

interface TenantUser {
  userId: string;
  email: string;
  displayName: string;
  accountEnabled: boolean;
  userType: string;
  licensed: boolean;
  installed: boolean;
}

interface Props {
  /** Logged-in admin's userId (for auth headers) */
  userId: string;
  /** Tenant-level enabled services — used to gray out tenant-disabled toggles */
  enabledServices: string[];
}

export default function UserManagement({ userId, enabledServices }: Props) {
  const [users, setUsers] = useState<TenantUser[]>([]);
  const [showAll, setShowAll] = useState(false);
  const [usersLoading, setUsersLoading] = useState(false);
  const [usersError, setUsersError] = useState<string | null>(null);
  const [selectedUser, setSelectedUser] = useState<string>('');
  const [disabledServices, setDisabledServices] = useState<string[]>([]);
  const [loading, setLoading] = useState(false);
  const [savingKey, setSavingKey] = useState<string | null>(null);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);
  const [confirmToggle, setConfirmToggle] = useState<{ key: string; checked: boolean } | null>(null);
  // Enforced draft mode: tenant-wide policy is read-only here; the per-user one is editable.
  const [draftPolicy, setDraftPolicy] = useState<{ tenant: boolean; user: boolean } | null>(null);
  const [savingDraftPolicy, setSavingDraftPolicy] = useState(false);
  const [draftPolicyMsg, setDraftPolicyMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);

  const headers = useMemo(() => ({
    'x-user-id': userId,
    'Content-Type': 'application/json',
  }), [userId]);

  // Load tenant users (Graph-sourced, filtered server-side)
  useEffect(() => {
    let cancelled = false;
    setUsersLoading(true);
    setUsersError(null);
    const url = `/api/manage/tenant-users${showAll ? '?includeAll=true' : ''}`;
    fetch(url, { headers })
      .then(async (r) => {
        if (!r.ok) {
          const data = (await r.json().catch(() => ({}))) as { error?: string };
          throw new Error(data.error ?? `HTTP ${r.status}`);
        }
        return r.json() as Promise<{ users?: TenantUser[] }>;
      })
      .then((data) => {
        if (cancelled) return;
        setUsers(Array.isArray(data.users) ? data.users : []);
      })
      .catch((err: unknown) => {
        if (cancelled) return;
        setUsers([]);
        setUsersError(err instanceof Error ? err.message : String(err));
      })
      .finally(() => { if (!cancelled) setUsersLoading(false); });
    return () => { cancelled = true; };
  }, [headers, showAll]);

  // Load overrides when a user is selected
  const loadOverrides = useCallback(async (targetId: string) => {
    if (!targetId) { setDisabledServices([]); return; }
    setLoading(true);
    setStatusMsg(null);
    try {
      const res = await fetch(`/api/manage/user-services?userId=${encodeURIComponent(targetId)}`, { headers });
      const data = (await res.json()) as { disabledServices?: string[] };
      setDisabledServices(Array.isArray(data.disabledServices) ? data.disabledServices : []);
    } catch {
      setDisabledServices([]);
    } finally {
      setLoading(false);
    }
  }, [headers]);

  useEffect(() => { loadOverrides(selectedUser); }, [selectedUser, loadOverrides]);

  const loadDraftPolicy = useCallback(async (targetId: string) => {
    if (!targetId) { setDraftPolicy(null); return; }
    try {
      const res = await fetch(`/api/manage/email-output-policy?userId=${encodeURIComponent(targetId)}`, { headers });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = (await res.json()) as { tenant?: { enforceDraft?: boolean }; user?: { enforceDraft?: boolean } };
      setDraftPolicy({ tenant: data.tenant?.enforceDraft === true, user: data.user?.enforceDraft === true });
    } catch {
      setDraftPolicy(null);
    }
  }, [headers]);

  useEffect(() => { loadDraftPolicy(selectedUser); }, [selectedUser, loadDraftPolicy]);

  const handleDraftPolicyToggle = async (enforceDraft: boolean) => {
    setSavingDraftPolicy(true);
    setDraftPolicyMsg(null);
    try {
      const res = await fetch('/api/manage/email-output-policy', {
        method: 'POST',
        headers,
        body: JSON.stringify({ scope: 'user', userId: selectedUser, enforceDraft }),
      });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      setDraftPolicy((prev) => ({ tenant: prev?.tenant === true, user: enforceDraft }));
      setDraftPolicyMsg({ kind: 'ok', text: 'Saved.' });
      setTimeout(() => setDraftPolicyMsg(null), 2000);
    } catch (err: unknown) {
      setDraftPolicyMsg({ kind: 'error', text: err instanceof Error ? err.message : String(err) });
      setTimeout(() => setDraftPolicyMsg(null), 4000);
    } finally {
      setSavingDraftPolicy(false);
    }
  };

  const handleToggle = async (key: string, checked: boolean) => {
    if (!checked && !confirmToggle) {
      setConfirmToggle({ key, checked });
      return;
    }
    setConfirmToggle(null);

    const updated = checked
      ? disabledServices.filter((s) => s !== key)
      : [...disabledServices, key];

    setSavingKey(key);
    setStatusMsg(null);
    try {
      const res = await fetch('/api/manage/user-services', {
        method: 'POST',
        headers,
        body: JSON.stringify({ userId: selectedUser, disabledServices: updated }),
      });
      if (!res.ok) {
        const data = (await res.json()) as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      setDisabledServices(updated);
      setStatusMsg({ kind: 'ok', text: 'Saved.' });
      setTimeout(() => setStatusMsg(null), 2000);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatusMsg({ kind: 'error', text: msg });
      setTimeout(() => setStatusMsg(null), 4000);
    } finally {
      setSavingKey(null);
    }
  };

  const selectedUserInfo = users.find((u) => u.userId === selectedUser);

  return (
    <>
      {/* User picker */}
      <div className="user-picker">
        <label className="user-picker-label">Select user:</label>
        <select
          className="user-picker-select"
          value={selectedUser}
          onChange={(e) => setSelectedUser(e.target.value)}
          disabled={usersLoading}
        >
          <option value="">
            {usersLoading ? 'Loading users…' : '— Choose a user —'}
          </option>
          {users.map((u) => (
            <option key={u.userId} value={u.userId}>
              {u.displayName} ({u.email}){u.installed ? ' ✓' : ''}
              {showAll && !u.accountEnabled ? ' [disabled]' : ''}
              {showAll && !u.licensed ? ' [unlicensed]' : ''}
            </option>
          ))}
        </select>
      </div>

      <label className="user-picker-show-all" style={{ display: 'flex', alignItems: 'center', gap: 6, marginTop: 6, fontSize: 12, color: '#666' }}>
        <input
          type="checkbox"
          checked={showAll}
          onChange={(e) => setShowAll(e.target.checked)}
        />
        Show all users (include disabled, unlicensed, and guest accounts)
      </label>

      {usersError && (
        <p className="status-msg status-msg--error" style={{ marginTop: 8 }}>
          Failed to load users: {usersError}
        </p>
      )}

      {/* Service toggles for selected user */}
      {selectedUser && (
        <>
          {loading ? (
            <p style={{ color: '#666', fontSize: 13, marginTop: 12 }}>Loading overrides…</p>
          ) : (
            <div className="services-list" style={{ marginTop: 12 }}>
              {ALL_SERVICES.map((svc) => {
                const tenantEnabled = enabledServices.includes(svc.key);
                const userDisabled = disabledServices.includes(svc.key);
                const effectiveEnabled = tenantEnabled && !userDisabled;
                const isSaving = savingKey === svc.key;

                return (
                  <div key={svc.key} className={`service-row${!tenantEnabled ? ' service-row--tenant-disabled' : ''}`}>
                    <label className="service-toggle">
                      <input
                        type="checkbox"
                        checked={effectiveEnabled}
                        disabled={!tenantEnabled || isSaving}
                        onChange={(e) => handleToggle(svc.key, e.target.checked)}
                      />
                      <span className="service-info">
                        <span className="service-name">{svc.label}</span>
                        {!tenantEnabled && (
                          <span className="service-desc">Disabled at tenant level</span>
                        )}
                        {tenantEnabled && userDisabled && (
                          <span className="service-desc service-desc--override">Disabled for this user</span>
                        )}
                      </span>
                    </label>
                    {userDisabled && tenantEnabled && (
                      <span className="override-badge">Override</span>
                    )}
                  </div>
                );
              })}
            </div>
          )}

          {/* Confirmation dialog */}
          {confirmToggle && (
            <div className="confirm-dialog">
              <p>
                Disable <strong>{ALL_SERVICES.find((s) => s.key === confirmToggle.key)?.label}</strong> for{' '}
                <strong>{selectedUserInfo?.displayName ?? selectedUser}</strong>?
              </p>
              <div className="confirm-actions">
                <button
                  className="confirm-btn confirm-btn--danger"
                  onClick={() => handleToggle(confirmToggle.key, confirmToggle.checked)}
                >
                  Disable
                </button>
                <button
                  className="confirm-btn"
                  onClick={() => setConfirmToggle(null)}
                >
                  Cancel
                </button>
              </div>
            </div>
          )}

          {statusMsg && (
            <p className={`status-msg status-msg--${statusMsg.kind}`}>{statusMsg.text}</p>
          )}

          {/* Enforced draft mode for this user */}
          {enabledServices.includes('mail') && draftPolicy && (
            <div className="services-list" style={{ marginTop: 16 }}>
              <div className="service-row">
                <label className="service-toggle">
                  <input
                    type="checkbox"
                    checked={draftPolicy.tenant || draftPolicy.user}
                    disabled={draftPolicy.tenant || savingDraftPolicy}
                    onChange={(e) => handleDraftPolicyToggle(e.target.checked)}
                  />
                  <span className="service-info">
                    <span className="service-name">Enforce draft mode</span>
                    <span className="service-desc">
                      {draftPolicy.tenant
                        ? 'Enforced for the whole organization under Email Output Policy; clear it there to change it per user.'
                        : draftPolicy.user
                          ? 'This user and any agent acting for them cannot switch email output to send mode.'
                          : 'Pin this user to draft mode so neither they nor an agent acting for them can switch to send mode.'}
                    </span>
                  </span>
                </label>
                {(draftPolicy.tenant || draftPolicy.user) && (
                  <span className="override-badge">{draftPolicy.tenant ? 'Tenant' : 'Enforced'}</span>
                )}
              </div>
              {draftPolicyMsg && (
                <p className={`status-msg status-msg--${draftPolicyMsg.kind}`}>{draftPolicyMsg.text}</p>
              )}
            </div>
          )}
        </>
      )}

      {!selectedUser && !usersLoading && (
        <p className="empty-state" style={{ marginTop: 12 }}>
          Select a user to manage their service access.
        </p>
      )}
    </>
  );
}
