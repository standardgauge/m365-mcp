import React, { useCallback, useEffect, useState } from 'react';

interface Policy {
  enforceDraft: boolean;
  updatedAt?: string;
  updatedBy?: string;
}

interface Props {
  /** Logged-in admin's userId (for auth headers) */
  userId: string;
}

/**
 * Tenant-wide enforced draft mode.
 *
 * The self-service email output mode can be flipped by the user, or by an
 * agent acting as the user through set_email_output_mode. This switch pins
 * every user in the tenant to draft mode so that neither path can lift it;
 * only an administrator can, here.
 */
export default function EmailOutputPolicy({ userId }: Props) {
  const [policy, setPolicy] = useState<Policy | null>(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);

  const headers = { 'x-user-id': userId, 'Content-Type': 'application/json' };

  const load = useCallback(async () => {
    setLoading(true);
    try {
      const res = await fetch('/api/manage/email-output-policy', { headers: { 'x-user-id': userId } });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      const data = (await res.json()) as { tenant?: Policy };
      setPolicy(data.tenant ?? { enforceDraft: false });
    } catch (err: unknown) {
      setPolicy(null);
      setStatusMsg({ kind: 'error', text: err instanceof Error ? err.message : String(err) });
    } finally {
      setLoading(false);
    }
  }, [userId]);

  useEffect(() => { load(); }, [load]);

  const handleToggle = async (enforceDraft: boolean) => {
    setSaving(true);
    setStatusMsg(null);
    try {
      const res = await fetch('/api/manage/email-output-policy', {
        method: 'POST',
        headers,
        body: JSON.stringify({ scope: 'tenant', enforceDraft }),
      });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      const data = (await res.json()) as Policy;
      setPolicy({ enforceDraft: data.enforceDraft, updatedAt: data.updatedAt, updatedBy: data.updatedBy });
      setStatusMsg({ kind: 'ok', text: 'Saved.' });
      setTimeout(() => setStatusMsg(null), 2000);
    } catch (err: unknown) {
      setStatusMsg({ kind: 'error', text: err instanceof Error ? err.message : String(err) });
      setTimeout(() => setStatusMsg(null), 4000);
    } finally {
      setSaving(false);
    }
  };

  const enforced = policy?.enforceDraft === true;

  return (
    <>
      <p className="section-hint">
        Draft mode keeps a person between an AI-composed email and its delivery. Users can
        switch their own mode to "send immediately", and so can an agent acting for them, which
        is what a prompt-injected agent would do. Enforcing draft mode here pins every user in
        the organization to draft mode: <code>send_mail</code> always saves to Drafts,{' '}
        <code>send_draft</code> refuses, and <code>set_email_output_mode</code> is refused and
        logged. Per-user enforcement is under User Management.
      </p>
      {loading ? (
        <p style={{ color: '#666', fontSize: 13 }}>Loading policy…</p>
      ) : (
        <div className="services-list">
          <div className="service-row">
            <label className="service-toggle">
              <input
                type="checkbox"
                checked={enforced}
                disabled={saving || policy === null}
                onChange={(e) => handleToggle(e.target.checked)}
              />
              <span className="service-info">
                <span className="service-name">Enforce draft mode for all users</span>
                <span className="service-desc">
                  {enforced
                    ? `Enforced${policy?.updatedAt ? ` since ${new Date(policy.updatedAt).toLocaleString()}` : ''}. Users and agents cannot switch to send mode.`
                    : 'Not enforced. Each user chooses their own mode; new users default to draft.'}
                </span>
              </span>
            </label>
            {enforced && <span className="override-badge">Enforced</span>}
          </div>
        </div>
      )}
      {statusMsg && (
        <p className={`status-msg status-msg--${statusMsg.kind}`}>{statusMsg.text}</p>
      )}
    </>
  );
}
