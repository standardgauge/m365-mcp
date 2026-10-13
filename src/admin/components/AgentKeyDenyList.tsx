import React, { useState } from 'react';
import DenyListManager from './DenyListManager';
import FolderBrowser from './FolderBrowser';
import MailFolderBrowser from './MailFolderBrowser';
import CalendarBrowser from './CalendarBrowser';
import OneDriveBrowser from './OneDriveBrowser';
import OneNoteBrowser from './OneNoteBrowser';
import ContactsBrowser from './ContactsBrowser';
import TeamsBrowser from './TeamsBrowser';
import type { SitePick } from './KeyScopeFields';

type DenyType = 'sharepoint' | 'mail' | 'calendar' | 'onedrive' | 'onenote' | 'contacts' | 'teams';
const DENY_TYPES: DenyType[] = ['sharepoint', 'mail', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'];

interface Props {
  userId: string;
  keyId: string;
  label: string;
  /** The services the key may use; only those get a list and a browser. */
  services: string[];
  /** The key's own site picks, else the tenant's allow-list, for the SharePoint browser. */
  sites?: SitePick[];
  /** False when the tenant allows narrowing only: entries can be added, not removed. */
  mayRemove: boolean;
  onClose: () => void;
}

/**
 * One agent key's own deny list: folders, mail folders, calendars and the rest
 * that this key may not reach even though its owner can. Uses the same folder
 * browsers as the owner's personal deny list; a Block here writes to the key,
 * not to the owner.
 */
export default function AgentKeyDenyList({ userId, keyId, label, services, sites, mayRemove, onClose }: Props) {
  const [refreshKey, setRefreshKey] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const types = DENY_TYPES.filter((t) => services.includes(t));
  const has = (t: DenyType) => types.includes(t);

  const block = async (type: DenyType, path: string) => {
    setError(null);
    const res = await fetch('/api/manage/deny-list/user', {
      method: 'POST',
      credentials: 'same-origin',
      headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
      body: JSON.stringify({ type, path, keyId }),
    });
    if (!res.ok) {
      const data = (await res.json().catch(() => ({}))) as { error?: string };
      setError(data.error ?? `HTTP ${res.status}`);
    }
    setRefreshKey((k) => k + 1);
  };
  const byName = (type: DenyType) => (_id: string, name: string) => block(type, name);

  return (
    <div className="card" style={{ marginTop: 12 }}>
      <p><strong>Blocked for “{label}”</strong> <code>{keyId}</code></p>
      <p className="section-hint">
        Keep this key out of a folder, mail folder or calendar you can still reach yourself. These entries apply
        only to calls made with this key, on top of your personal deny list and your organization's.
        They take effect on the key's next call.
        {!mayRemove && <> Your organization allows a key to be narrowed only, so an entry cannot be removed once added.</>}
      </p>
      {error && <p className="error-msg">{error}</p>}
      {types.length === 0 ? (
        <p className="empty-state">This key has no service with folders to block.</p>
      ) : (
        <>
          <DenyListManager
            key={refreshKey}
            scope="user"
            userId={userId}
            keyId={keyId}
            readOnly={!mayRemove}
            readOnlyBadge="Narrow only"
            hideAddInput
            visibleTypes={types}
            bare
          />
          <div className="browsers">
            {has('sharepoint') && <FolderBrowser userId={userId} onSelect={(path) => block('sharepoint', path)} allowedSites={sites} />}
            {has('mail') && <MailFolderBrowser userId={userId} onSelect={(path) => block('mail', path)} />}
            {has('calendar') && <CalendarBrowser userId={userId} onSelect={byName('calendar')} />}
            {has('onedrive') && <OneDriveBrowser userId={userId} onSelect={byName('onedrive')} />}
            {has('onenote') && <OneNoteBrowser userId={userId} onSelect={byName('onenote')} />}
            {has('contacts') && <ContactsBrowser userId={userId} onSelect={byName('contacts')} />}
            {has('teams') && <TeamsBrowser userId={userId} onSelect={byName('teams')} />}
          </div>
        </>
      )}
      <div className="confirm-actions">
        <button type="button" className="confirm-btn" onClick={onClose}>Done</button>
      </div>
    </div>
  );
}
