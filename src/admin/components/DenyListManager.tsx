import React, { useCallback, useEffect, useState } from 'react';

interface DenyEntry {
  path: string;
  description: string;
  addedBy: string;
  addedByName?: string;
  addedAt: string;
}

type DenyListType = 'sharepoint' | 'mail' | 'calendar' | 'onedrive' | 'onenote' | 'contacts' | 'teams';

const TYPE_LABELS: Record<DenyListType, string> = {
  sharepoint: 'SharePoint Folders',
  mail: 'Mail Folders',
  calendar: 'Calendars',
  onedrive: 'OneDrive Folders',
  onenote: 'OneNote Notebooks',
  contacts: 'Contact Folders',
  teams: 'Teams',
};

const ALL_TYPES: DenyListType[] = ['sharepoint', 'mail', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'];

interface Props {
  /** 'global' = admin-managed deny list; 'user' = per-user personal list */
  scope: 'global' | 'user';
  userId: string;
  /** When scope='user', this is the target user ID (may differ from admin's userId) */
  targetUserId?: string;
  /** When true, hides add/remove controls and shows a "Global Admin Setting" badge instead */
  readOnly?: boolean;
  /** Overrides the default card heading */
  label?: string;
  /** Pre-fills the SharePoint path input (e.g. from FolderBrowser selection) */
  externalSpPath?: string;
  /** Pre-fills the mail path input (e.g. from MailFolderBrowser selection) */
  externalMailPath?: string;
  /** Display name of the logged-in user, stored alongside the UUID addedBy */
  addedByName?: string;
  /** When true, hides the path input and Block button (entries added externally) */
  hideAddInput?: boolean;
  /** Limit which service types are rendered (default: all) */
  visibleTypes?: DenyListType[];
  /** When true, skip the .card wrapper and h2 heading (used when nested in CollapsibleSection) */
  bare?: boolean;
}

async function apiFetch(path: string, userId: string, options?: RequestInit) {
  const res = await fetch(`/api/${path}`, {
    ...options,
    headers: { 'x-user-id': userId, 'Content-Type': 'application/json', ...(options?.headers ?? {}) },
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error ?? `${res.status} ${res.statusText}`);
  }
  return res.json();
}

export default function DenyListManager({ scope, userId, targetUserId, readOnly = false, label, externalSpPath, externalMailPath, addedByName, hideAddInput = false, visibleTypes, bare = false }: Props) {
  const [entries, setEntries] = useState<Record<DenyListType, DenyEntry[]>>({
    sharepoint: [], mail: [], calendar: [], onedrive: [], onenote: [], contacts: [], teams: [],
  });
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [newPath, setNewPath] = useState<Record<DenyListType, string>>({
    sharepoint: '', mail: '', calendar: '', onedrive: '', onenote: '', contacts: '', teams: '',
  });
  const [newDesc, setNewDesc] = useState<Record<DenyListType, string>>({
    sharepoint: '', mail: '', calendar: '', onedrive: '', onenote: '', contacts: '', teams: '',
  });

  useEffect(() => { if (externalSpPath) setNewPath((p) => ({ ...p, sharepoint: externalSpPath })); }, [externalSpPath]);
  useEffect(() => { if (externalMailPath) setNewPath((p) => ({ ...p, mail: externalMailPath })); }, [externalMailPath]);

  const listEndpoint = scope === 'global' ? 'manage/deny-list/global' : 'manage/deny-list/user';
  const resolvedTarget = scope === 'user' ? (targetUserId ?? userId) : undefined;
  const activeTypes = visibleTypes ?? ALL_TYPES;

  const load = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const results = await Promise.all(
        activeTypes.map((type) =>
          apiFetch(
            `${listEndpoint}?type=${type}${resolvedTarget ? `&targetUserId=${resolvedTarget}` : ''}`,
            userId
          ).then((data) => ({ type, entries: (data.entries ?? []) as DenyEntry[] }))
        )
      );
      setEntries((prev) => {
        const next = { ...prev };
        for (const r of results) next[r.type] = r.entries;
        return next;
      });
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to load deny list');
    } finally {
      setLoading(false);
    }
  }, [listEndpoint, userId, resolvedTarget, activeTypes.join(',')]); // eslint-disable-line react-hooks/exhaustive-deps

  useEffect(() => { load(); }, [load]);

  const addEntry = async (type: DenyListType, path: string, description: string) => {
    if (!path.trim()) return;
    setError(null);
    try {
      await apiFetch(
        scope === 'global' ? 'manage/deny-list/global' : 'manage/deny-list/user',
        userId,
        {
          method: 'POST',
          body: JSON.stringify({
            type,
            path: path.trim(),
            description: description.trim(),
            ...(resolvedTarget ? { targetUserId: resolvedTarget } : {}),
            ...(addedByName ? { addedByName } : {}),
          }),
        }
      );
      setNewPath((p) => ({ ...p, [type]: '' }));
      setNewDesc((d) => ({ ...d, [type]: '' }));
      await load();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to add entry');
    }
  };

  const removeEntry = async (type: DenyListType, path: string) => {
    setError(null);
    try {
      await apiFetch(
        scope === 'global' ? 'manage/deny-list/global' : 'manage/deny-list/user',
        userId,
        {
          method: 'DELETE',
          body: JSON.stringify({
            type,
            path,
            ...(resolvedTarget ? { targetUserId: resolvedTarget } : {}),
          }),
        }
      );
      await load();
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to remove entry');
    }
  };

  const renderTable = (type: DenyListType) => {
    const typeEntries = entries[type];
    const path = newPath[type];
    const desc = newDesc[type];
    return (
      <div key={type} style={{ marginBottom: 20 }}>
        <h3>
          {TYPE_LABELS[type]}
          <span className="badge-count">{typeEntries.length}</span>
        </h3>

        {typeEntries.length === 0 ? (
          <p className="empty-state">No entries.</p>
        ) : (
          <table className="deny-table">
            <thead>
              <tr>
                <th>Path</th>
                {scope === 'global' && <th>Description</th>}
                <th>Added by</th>
                <th>Date</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {typeEntries.map((e) => (
                <tr key={e.path}>
                  <td style={{ fontFamily: 'monospace', fontSize: 12 }}>{e.path}</td>
                  {scope === 'global' && <td>{e.description}</td>}
                  <td>{e.addedByName || e.addedBy}</td>
                  <td>{new Date(e.addedAt).toLocaleDateString()}</td>
                  <td>
                    {readOnly ? (
                      <span className="admin-setting-badge">Global Admin Setting</span>
                    ) : (
                      <button className="remove-btn" onClick={() => removeEntry(type, e.path)}>
                        Remove
                      </button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}

        {!readOnly && !hideAddInput && (
          <div style={{ display: 'flex', gap: 8, marginTop: 10 }}>
            <input
              style={{ flex: 2, padding: '5px 8px', fontSize: 13, border: '1px solid #ccc', borderRadius: 4 }}
              placeholder="Path"
              value={path}
              onChange={(e) => setNewPath((p) => ({ ...p, [type]: e.target.value }))}
              onKeyDown={(e) => e.key === 'Enter' && addEntry(type, path, desc)}
            />
            {scope === 'global' && (
              <input
                style={{ flex: 1, padding: '5px 8px', fontSize: 13, border: '1px solid #ccc', borderRadius: 4 }}
                placeholder="Description (optional)"
                value={desc}
                onChange={(e) => setNewDesc((d) => ({ ...d, [type]: e.target.value }))}
                onKeyDown={(e) => e.key === 'Enter' && addEntry(type, path, desc)}
              />
            )}
            <button
              style={{ padding: '5px 14px', background: '#0078d4', color: '#fff', border: 'none', borderRadius: 4, cursor: 'pointer', fontSize: 13 }}
              onClick={() => addEntry(type, path, desc)}
            >
              Block
            </button>
          </div>
        )}
      </div>
    );
  };

  const title = label ?? (scope === 'global' ? 'Global Deny List' : 'Personal Deny List');

  const content = (
    <>
      {loading && <p style={{ color: '#666', fontSize: 13 }}>Loading…</p>}
      {error && <p className="error-msg">{error}</p>}
      {!loading && activeTypes.map((type) => renderTable(type))}
    </>
  );

  if (bare) return content;

  return (
    <div className="card">
      <h2>{title}</h2>
      {content}
    </div>
  );
}
