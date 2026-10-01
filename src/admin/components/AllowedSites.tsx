import React, { useEffect, useState } from 'react';
import FolderBrowser from './FolderBrowser';

interface AllowedSite {
  id: string;
  name: string;
}

interface Site {
  id: string;
  displayName: string;
  webUrl: string;
}

interface Props {
  userId: string;
  isGlobalAdmin: boolean;
}

export default function AllowedSites({ userId, isGlobalAdmin }: Props) {
  const [allowedSites, setAllowedSites] = useState<AllowedSite[]>([]);
  const [allSites, setAllSites] = useState<Site[]>([]);
  const [selectedSiteId, setSelectedSiteId] = useState<string>('');
  const [saving, setSaving] = useState(false);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);
  const [manualUrl, setManualUrl] = useState<string>('');
  const [resolving, setResolving] = useState(false);
  const [siteSearchError, setSiteSearchError] = useState<string | null>(null);

  useEffect(() => {
    if (!userId) return;
    fetch('/api/manage/allowed-sites', {
      headers: { 'x-user-id': userId },
    })
      .then((r) => r.json())
      .then((data: { allowedSites?: AllowedSite[] }) => {
        if (Array.isArray(data.allowedSites)) setAllowedSites(data.allowedSites);
      })
      .catch(() => {});
  }, [userId]);

  useEffect(() => {
    setSiteSearchError(null);
    fetch('/api/sharepoint/sites?admin=true', {
      headers: { 'x-user-id': userId },
    })
      .then(async (r) => {
        if (!r.ok) {
          const body = await r.json().catch(() => ({})) as { error?: string };
          throw new Error(body.error ?? `HTTP ${r.status}`);
        }
        return r.json();
      })
      .then((data: { sites?: Site[] }) => {
        if (Array.isArray(data.sites)) setAllSites(data.sites);
      })
      .catch((err: unknown) => {
        const msg = err instanceof Error ? err.message : String(err);
        setSiteSearchError(msg);
      });
  }, [userId]);

  const save = async (updated: AllowedSite[]) => {
    setSaving(true);
    setStatusMsg(null);
    try {
      const res = await fetch('/api/manage/allowed-sites', {
        method: 'POST',
        headers: {
          'x-user-id': userId,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ allowedSites: updated }),
      });
      if (!res.ok) {
        const data = await res.json() as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      setAllowedSites(updated);
      setStatusMsg({ kind: 'ok', text: 'Saved.' });
      setTimeout(() => setStatusMsg(null), 2000);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatusMsg({ kind: 'error', text: msg });
      setTimeout(() => setStatusMsg(null), 4000);
    } finally {
      setSaving(false);
    }
  };

  const handleAdd = () => {
    if (!selectedSiteId) return;
    const site = allSites.find((s) => s.id === selectedSiteId);
    if (!site) return;
    if (allowedSites.some((s) => s.id === selectedSiteId)) return;
    const updated = [...allowedSites, { id: site.id, name: site.displayName }];
    setSelectedSiteId('');
    save(updated);
  };

  const handleRemove = (id: string) => {
    save(allowedSites.filter((s) => s.id !== id));
  };

  /** Resolve a SharePoint site URL server-side (via the session) and add it. */
  const handleManualAdd = async () => {
    const url = manualUrl.trim();
    if (!url) return;
    setResolving(true);
    setStatusMsg(null);
    try {
      // Server resolves the URL to a Graph site id using the session token —
      // no client-side Graph token needed.
      const res = await fetch(`/api/sharepoint/resolve-site?url=${encodeURIComponent(url)}`, {
        headers: { 'x-user-id': userId },
      });
      if (!res.ok) {
        const body = await res.json().catch(() => ({})) as { error?: string };
        throw new Error(body.error ?? `Could not resolve site (HTTP ${res.status})`);
      }
      const site = await res.json() as { id: string; displayName: string };

      if (allowedSites.some((s) => s.id === site.id)) {
        setStatusMsg({ kind: 'error', text: 'Site is already in the allowlist.' });
        setTimeout(() => setStatusMsg(null), 3000);
        return;
      }

      const updated = [...allowedSites, { id: site.id, name: site.displayName }];
      setManualUrl('');
      await save(updated);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setStatusMsg({ kind: 'error', text: `Could not resolve site: ${msg}` });
      setTimeout(() => setStatusMsg(null), 5000);
    } finally {
      setResolving(false);
    }
  };

  const availableToAdd = allSites.filter((s) => !allowedSites.some((a) => a.id === s.id));

  return (
    <>
      {allowedSites.length === 0 ? (
        <p className="empty-state">
          No sites added yet — users will not be able to browse SharePoint.
        </p>
      ) : (
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13, marginBottom: 16 }}>
          <thead>
            <tr style={{ borderBottom: '1px solid #e0e0e0' }}>
              <th style={{ textAlign: 'left', padding: '6px 8px', fontWeight: 600 }}>Name</th>
              {isGlobalAdmin && <th style={{ width: 80 }} />}
            </tr>
          </thead>
          <tbody>
            {allowedSites.map((site) => (
              <tr key={site.id} style={{ borderBottom: '1px solid #f0f0f0' }}>
                <td style={{ padding: '6px 8px' }}>{site.name}</td>
                {isGlobalAdmin && (
                  <td style={{ padding: '4px 8px' }}>
                    <button
                      className="remove-btn"
                      disabled={saving}
                      onClick={() => handleRemove(site.id)}
                    >
                      Remove
                    </button>
                  </td>
                )}
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {isGlobalAdmin && (
        <>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <select
              style={{ flex: 1, padding: '6px 8px', fontSize: 13, borderRadius: 4, border: '1px solid #ccc' }}
              value={selectedSiteId}
              onChange={(e) => setSelectedSiteId(e.target.value)}
              disabled={saving}
            >
              <option value="">— select a site to add —</option>
              {availableToAdd.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.displayName}
                </option>
              ))}
            </select>
            <button
              className="primary-btn"
              disabled={!selectedSiteId || saving}
              onClick={handleAdd}
            >
              Add
            </button>
          </div>
          {siteSearchError && (
            <p style={{ fontSize: 12, color: '#b91c1c', margin: '6px 0 0' }}>
              Site search unavailable: {siteSearchError}
            </p>
          )}
          {availableToAdd.length === 0 && !siteSearchError && (
            <p style={{ fontSize: 12, color: '#6b7280', margin: '6px 0 0' }}>
              Don&apos;t see your site? SharePoint search only indexes sites with recent activity.
              Use <strong>Add by URL</strong> below for any site you have access to.
            </p>
          )}
          <div style={{ marginTop: 16, paddingTop: 12, borderTop: '1px solid #e5e7eb' }}>
            <p style={{ fontSize: 13, fontWeight: 500, margin: '0 0 6px', color: '#374151' }}>
              Add by URL (recommended)
            </p>
            <p style={{ fontSize: 12, color: '#6b7280', margin: '0 0 8px' }}>
              Works for any SharePoint site you have access to, including new sites not yet
              indexed by search.
            </p>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <input
                type="text"
                style={{ flex: 1, padding: '6px 8px', fontSize: 13, borderRadius: 4, border: '1px solid #ccc' }}
                placeholder="https://yourtenant.sharepoint.com/sites/your-site"
                value={manualUrl}
                onChange={(e) => setManualUrl(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter') handleManualAdd(); }}
                disabled={saving || resolving}
              />
              <button
                className="primary-btn"
                disabled={!manualUrl.trim() || saving || resolving}
                onClick={handleManualAdd}
              >
                {resolving ? 'Resolving…' : 'Add by URL'}
              </button>
            </div>
          </div>
        </>
      )}

      {statusMsg && (
        <p className={`status-msg status-msg--${statusMsg.kind}`}>{statusMsg.text}</p>
      )}
    </>
  );
}
