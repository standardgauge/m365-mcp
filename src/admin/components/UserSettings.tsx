import React, { useCallback, useEffect, useState } from 'react';
import DenyListManager from './DenyListManager';
import CollapsibleSection from './CollapsibleSection';
import FolderBrowser from './FolderBrowser';
import MailFolderBrowser from './MailFolderBrowser';
import CalendarBrowser from './CalendarBrowser';
import OneDriveBrowser from './OneDriveBrowser';
import OneNoteBrowser from './OneNoteBrowser';
import ContactsBrowser from './ContactsBrowser';
import TeamsBrowser from './TeamsBrowser';

const ALL_SERVICES: { key: string; label: string; description: string }[] = [
  { key: 'mail', label: 'Mail', description: 'Read and manage mailbox folders' },
  { key: 'sharepoint', label: 'SharePoint', description: 'Read and manage SharePoint sites and files' },
  { key: 'calendar', label: 'Calendar', description: 'Manage calendar events' },
  { key: 'onedrive', label: 'OneDrive', description: 'Access OneDrive files' },
  { key: 'onenote', label: 'OneNote', description: 'Read OneNote notebooks' },
  { key: 'contacts', label: 'Contacts', description: 'Manage contacts' },
  { key: 'teams', label: 'Teams', description: 'Read and manage Teams channels and messages' },
];

interface Props {
  userId: string;
  accountName?: string;
  enabledServices?: string[];
  allowedSites?: Array<{ id: string; name: string }>;
  instanceName?: string;
}

type EmailOutputMode = 'draft' | 'send';

export default function UserSettings({ userId, accountName, enabledServices = ['mail', 'sharepoint'], allowedSites, instanceName }: Props) {
  const [refreshKey, setRefreshKey] = useState(0);
  const [disabledServices, setDisabledServices] = useState<string[]>([]);
  const [savingKey, setSavingKey] = useState<string | null>(null);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);
  const [emailOutputMode, setEmailOutputMode] = useState<EmailOutputMode>('draft');
  const [savingEmailMode, setSavingEmailMode] = useState(false);
  const [emailStatusMsg, setEmailStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);
  // Enforced draft mode: when set, the radios are locked and only an admin can change it.
  const [emailModeEnforcedBy, setEmailModeEnforcedBy] = useState<'tenant' | 'user' | null>(null);

  // Load the user's own service overrides
  const loadOverrides = useCallback(async () => {
    try {
      const res = await fetch(`/api/manage/user-services?userId=${encodeURIComponent(userId)}`, {
        headers: { 'x-user-id': userId },
      });
      const data = (await res.json()) as { disabledServices?: string[] };
      if (Array.isArray(data.disabledServices)) setDisabledServices(data.disabledServices);
    } catch { /* ignore — defaults to empty */ }
  }, [userId]);

  // Load the user's email output mode (defaults to draft)
  const loadEmailMode = useCallback(async () => {
    try {
      const res = await fetch(`/api/mail/settings?userId=${encodeURIComponent(userId)}`, {
        headers: { 'x-user-id': userId },
      });
      const data = (await res.json()) as { emailOutputMode?: string; enforced?: boolean; enforcedBy?: 'tenant' | 'user' | null };
      if (data.emailOutputMode === 'send' || data.emailOutputMode === 'draft') {
        setEmailOutputMode(data.emailOutputMode);
      }
      setEmailModeEnforcedBy(data.enforced === true ? (data.enforcedBy ?? 'tenant') : null);
    } catch { /* ignore — defaults to draft */ }
  }, [userId]);

  useEffect(() => { loadOverrides(); }, [loadOverrides]);
  useEffect(() => { loadEmailMode(); }, [loadEmailMode]);

  const handleEmailModeChange = async (mode: EmailOutputMode) => {
    if (mode === emailOutputMode || emailModeEnforcedBy) return;
    setSavingEmailMode(true);
    setEmailStatusMsg(null);
    try {
      const res = await fetch('/api/mail/settings', {
        method: 'POST',
        headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
        body: JSON.stringify({ userId, emailOutputMode: mode }),
      });
      if (!res.ok) {
        const data = (await res.json()) as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      setEmailOutputMode(mode);
      setEmailStatusMsg({ kind: 'ok', text: 'Saved.' });
      setTimeout(() => setEmailStatusMsg(null), 2000);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setEmailStatusMsg({ kind: 'error', text: msg });
      setTimeout(() => setEmailStatusMsg(null), 4000);
    } finally {
      setSavingEmailMode(false);
    }
  };

  const handleServiceToggle = async (key: string, checked: boolean) => {
    const updated = checked
      ? disabledServices.filter((s) => s !== key)
      : [...disabledServices, key];

    setSavingKey(key);
    setStatusMsg(null);
    try {
      const res = await fetch('/api/manage/user-services', {
        method: 'POST',
        headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
        body: JSON.stringify({ userId, disabledServices: updated }),
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

  const spEnabled = enabledServices.includes('sharepoint');
  const mailEnabled = enabledServices.includes('mail');
  const calendarEnabled = enabledServices.includes('calendar');
  const onedriveEnabled = enabledServices.includes('onedrive');
  const onenoteEnabled = enabledServices.includes('onenote');
  const contactsEnabled = enabledServices.includes('contacts');
  const teamsEnabled = enabledServices.includes('teams');

  const handleSpBlock = async (path: string) => {
    await fetch('/api/manage/deny-list/user', {
      method: 'POST',
      headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'sharepoint', path, targetUserId: userId, ...(accountName ? { addedByName: accountName } : {}) }),
    });
    setRefreshKey((k) => k + 1);
  };

  const handleMailBlock = async (path: string) => {
    await fetch('/api/manage/deny-list/user', {
      method: 'POST',
      headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'mail', path, targetUserId: userId, ...(accountName ? { addedByName: accountName } : {}) }),
    });
    setRefreshKey((k) => k + 1);
  };

  const handleCalendarBlock = async (_id: string, name: string) => {
    await fetch('/api/manage/deny-list/user', {
      method: 'POST',
      headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'calendar', path: name, targetUserId: userId, ...(accountName ? { addedByName: accountName } : {}) }),
    });
    setRefreshKey((k) => k + 1);
  };

  const handleOnedriveBlock = async (_id: string, name: string) => {
    await fetch('/api/manage/deny-list/user', {
      method: 'POST',
      headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'onedrive', path: name, targetUserId: userId, ...(accountName ? { addedByName: accountName } : {}) }),
    });
    setRefreshKey((k) => k + 1);
  };

  const handleOnenoteBlock = async (_id: string, name: string) => {
    await fetch('/api/manage/deny-list/user', {
      method: 'POST',
      headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'onenote', path: name, targetUserId: userId, ...(accountName ? { addedByName: accountName } : {}) }),
    });
    setRefreshKey((k) => k + 1);
  };

  const handleContactsBlock = async (_id: string, name: string) => {
    await fetch('/api/manage/deny-list/user', {
      method: 'POST',
      headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'contacts', path: name, targetUserId: userId, ...(accountName ? { addedByName: accountName } : {}) }),
    });
    setRefreshKey((k) => k + 1);
  };

  const handleTeamsBlock = async (_id: string, name: string) => {
    await fetch('/api/manage/deny-list/user', {
      method: 'POST',
      headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
      body: JSON.stringify({ type: 'teams', path: name, targetUserId: userId, ...(accountName ? { addedByName: accountName } : {}) }),
    });
    setRefreshKey((k) => k + 1);
  };

  const visibleTypes = [
    ...(spEnabled ? ['sharepoint' as const] : []),
    ...(mailEnabled ? ['mail' as const] : []),
    ...(calendarEnabled ? ['calendar' as const] : []),
    ...(onedriveEnabled ? ['onedrive' as const] : []),
    ...(onenoteEnabled ? ['onenote' as const] : []),
    ...(contactsEnabled ? ['contacts' as const] : []),
    ...(teamsEnabled ? ['teams' as const] : []),
  ];

  const anyBrowserEnabled = spEnabled || mailEnabled || calendarEnabled || onedriveEnabled || onenoteEnabled || contactsEnabled || teamsEnabled;
  const optOutCount = disabledServices.length;

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
      {/* Welcome banner */}
      <div className="card welcome-banner">
        <h2>Welcome to {instanceName ?? 'the MCP Server'}</h2>
        <p>
          This page lets you control what Claude can access in your Microsoft 365 environment.
          Use the toggles below to opt out of specific services, or browse your folders to block
          access to individual items.
        </p>
      </div>

      {/* Self-service opt-out */}
      <CollapsibleSection
        storageKey="settings-my-services"
        title="My Services"
        count={optOutCount > 0 ? optOutCount : undefined}
      >
        <p className="section-hint">
          Uncheck a service to prevent Claude from accessing it for your account.
          Your admin controls which services are available organization-wide — services
          disabled by your admin cannot be re-enabled here.
        </p>
        <div className="services-list">
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
                    onChange={(e) => handleServiceToggle(svc.key, e.target.checked)}
                  />
                  <span className="service-info">
                    <span className="service-name">{svc.label}</span>
                    <span className="service-desc">
                      {!tenantEnabled
                        ? 'Not available — disabled by your admin'
                        : svc.description}
                    </span>
                  </span>
                </label>
                {userDisabled && tenantEnabled && (
                  <span className="override-badge">Opted out</span>
                )}
              </div>
            );
          })}
        </div>
        {statusMsg && (
          <p className={`status-msg status-msg--${statusMsg.kind}`}>{statusMsg.text}</p>
        )}
      </CollapsibleSection>

      {/* Email output mode — draft vs. send */}
      {mailEnabled && (
        <CollapsibleSection storageKey="settings-email-output-mode" title="My Email Settings">
          <p className="section-hint">
            Controls what happens when Claude composes an email for you. Draft mode keeps you in
            control — nothing leaves your mailbox until you review and send it yourself.
          </p>
          {emailModeEnforcedBy && (
            <p className="status-msg status-msg--ok" style={{ marginBottom: 8 }}>
              Draft mode is enforced by your administrator
              {emailModeEnforcedBy === 'tenant' ? ' for your organization' : ' for your account'}.
              It cannot be changed here, and Claude cannot change it either.
            </p>
          )}
          <div className="services-list">
            <label className="service-toggle">
              <input
                type="radio"
                name="email-output-mode"
                checked={emailOutputMode === 'draft'}
                disabled={savingEmailMode || emailModeEnforcedBy !== null}
                onChange={() => handleEmailModeChange('draft')}
              />
              <span className="service-info">
                <span className="service-name">Save as draft for review (recommended)</span>
                <span className="service-desc">
                  Emails Claude composes are saved to your Drafts folder. You review and send them.
                </span>
              </span>
            </label>
            <label className="service-toggle">
              <input
                type="radio"
                name="email-output-mode"
                checked={emailOutputMode === 'send'}
                disabled={savingEmailMode || emailModeEnforcedBy !== null}
                onChange={() => handleEmailModeChange('send')}
              />
              <span className="service-info">
                <span className="service-name">Send immediately</span>
                <span className="service-desc">
                  Emails Claude composes are delivered right away, without a review step.
                </span>
              </span>
            </label>
          </div>
          {emailStatusMsg && (
            <p className={`status-msg status-msg--${emailStatusMsg.kind}`}>{emailStatusMsg.text}</p>
          )}
        </CollapsibleSection>
      )}

      {/* Read-only view of the global deny list */}
      <CollapsibleSection storageKey="settings-org-deny-list" title="Organizational Deny List" defaultCollapsed>
        <p className="section-hint">
          The entries below are blocked for all users by your organization's administrators.
          This list is read-only. Contact your admin to request changes.
        </p>
        <DenyListManager
          scope="global"
          userId={userId}
          readOnly
          label="Organizational Deny List"
          visibleTypes={visibleTypes}
          bare
        />
      </CollapsibleSection>

      {/* Personal deny list */}
      <CollapsibleSection storageKey="settings-personal-deny-list" title="Personal Deny List">
        <DenyListManager
          key={refreshKey}
          scope="user"
          userId={userId}
          targetUserId={userId}
          label="Personal Deny List"
          hideAddInput
          visibleTypes={visibleTypes}
          bare
        />
      </CollapsibleSection>

      {/* Folder browsers */}
      {anyBrowserEnabled && (
        <CollapsibleSection storageKey="settings-browsers" title="Browse & Block">
          <p className="section-hint">
            Browse your folders and click Block to add items to your personal deny list.
          </p>
          <div className="browsers">
            {spEnabled && <FolderBrowser userId={userId} onSelect={handleSpBlock} allowedSites={allowedSites} />}
            {mailEnabled && <MailFolderBrowser userId={userId} onSelect={handleMailBlock} />}
            {calendarEnabled && <CalendarBrowser userId={userId} onSelect={handleCalendarBlock} />}
            {onedriveEnabled && <OneDriveBrowser userId={userId} onSelect={handleOnedriveBlock} />}
            {onenoteEnabled && <OneNoteBrowser userId={userId} onSelect={handleOnenoteBlock} />}
            {contactsEnabled && <ContactsBrowser userId={userId} onSelect={handleContactsBlock} />}
            {teamsEnabled && <TeamsBrowser userId={userId} onSelect={handleTeamsBlock} />}
          </div>
        </CollapsibleSection>
      )}
    </div>
  );
}
