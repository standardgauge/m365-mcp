import React, { useState } from 'react';

const ALL_SERVICES: { key: string; label: string; description: string; implemented: boolean }[] = [
  { key: 'mail', label: 'Mail', description: 'Read and manage mailbox folders', implemented: true },
  { key: 'sharepoint', label: 'SharePoint', description: 'Read and manage SharePoint sites and files', implemented: true },
  { key: 'calendar', label: 'Calendar', description: 'Manage calendar events', implemented: true },
  { key: 'onedrive', label: 'OneDrive', description: 'Access OneDrive files', implemented: true },
  { key: 'onenote', label: 'OneNote', description: 'Read OneNote notebooks', implemented: true },
  { key: 'contacts', label: 'Contacts', description: 'Manage contacts', implemented: true },
  { key: 'teams', label: 'Teams', description: 'Read and manage Teams channels and messages', implemented: true },
];

interface Props {
  enabledServices: string[];
  setEnabledServices: (services: string[]) => void;
  userId: string;
}

export default function ConnectedServices({ enabledServices, setEnabledServices, userId }: Props) {
  const [savingKey, setSavingKey] = useState<string | null>(null);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);

  const handleToggle = async (key: string, checked: boolean) => {
    const updated = checked
      ? [...enabledServices, key]
      : enabledServices.filter((s) => s !== key);

    setSavingKey(key);
    setStatusMsg(null);
    try {
      const res = await fetch('/api/manage/services', {
        method: 'POST',
        headers: {
          'x-user-id': userId,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ enabledServices: updated }),
      });
      if (!res.ok) {
        const data = await res.json() as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      setEnabledServices(updated);
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

  return (
    <>
      <div className="services-list">
        {ALL_SERVICES.map((svc) => {
          const enabled = enabledServices.includes(svc.key);
          const isSaving = savingKey === svc.key;
          return (
            <div key={svc.key} className={`service-row${!svc.implemented ? ' service-row--coming-soon' : ''}`}>
              <label className="service-toggle">
                <input
                  type="checkbox"
                  checked={enabled}
                  disabled={!svc.implemented || isSaving}
                  onChange={(e) => handleToggle(svc.key, e.target.checked)}
                />
                <span className="service-info">
                  <span className="service-name">{svc.label}</span>
                  <span className="service-desc">{svc.description}</span>
                </span>
              </label>
              {!svc.implemented && <span className="coming-soon-badge">Coming soon</span>}
            </div>
          );
        })}
      </div>
      {statusMsg && (
        <p className={`status-msg status-msg--${statusMsg.kind}`}>{statusMsg.text}</p>
      )}
    </>
  );
}
