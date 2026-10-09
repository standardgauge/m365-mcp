import React, { useCallback, useEffect, useState } from 'react';

type Mode = 'allow' | 'internal' | 'block';
type Channel = 'calendarInvites' | 'eventResponses' | 'teamsMessages';

interface Policy {
  calendarInvites: Mode;
  eventResponses: Mode;
  teamsMessages: Mode;
  updatedAt?: string;
}

const CHANNELS: { key: Channel; service: string; label: string; desc: string }[] = [
  {
    key: 'calendarInvites',
    service: 'calendar',
    label: 'Calendar invitations',
    desc: 'create_event, update_event and move_event on a meeting the user organizes. Attendees get the invitation or update as soon as the event is written.',
  },
  {
    key: 'eventResponses',
    service: 'calendar',
    label: 'Comments on meeting responses',
    desc: 'respond_to_event with a comment, which goes to the organizer. A plain accept, tentative or decline is not affected.',
  },
  {
    key: 'teamsMessages',
    service: 'teams',
    label: 'Teams messages',
    desc: 'send_chat_message and send_channel_message. Under "internal only", every member of the chat or channel is checked.',
  },
];

const MODE_LABEL: Record<Mode, string> = {
  allow: 'Allow',
  internal: 'Internal only',
  block: 'Block',
};

const RANK: Record<Mode, number> = { allow: 0, internal: 1, block: 2 };

interface Props {
  /** Logged-in admin's userId (for auth headers) */
  userId: string;
  /** Tenant-level enabled services; a channel whose service is off is shown disabled */
  enabledServices: string[];
  /** Set to edit one user's row; omit for the tenant-wide policy */
  targetUserId?: string;
}

/**
 * Outbound policy: the channels enforced draft mode does not hold. Each one
 * can be allowed, limited to recipients inside the organization, or blocked
 * for agents. A per-user setting can only tighten the tenant-wide one.
 */
export default function OutboundPolicy({ userId, enabledServices, targetUserId }: Props) {
  const [policy, setPolicy] = useState<Policy | null>(null);
  const [tenantFloor, setTenantFloor] = useState<Policy | null>(null);
  const [saving, setSaving] = useState<Channel | null>(null);
  const [statusMsg, setStatusMsg] = useState<{ kind: 'ok' | 'error'; text: string } | null>(null);

  const load = useCallback(async () => {
    try {
      const url = targetUserId
        ? `/api/manage/outbound-policy?userId=${encodeURIComponent(targetUserId)}`
        : '/api/manage/outbound-policy';
      const res = await fetch(url, { headers: { 'x-user-id': userId } });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      const data = (await res.json()) as { tenant: Policy; user?: Policy };
      setTenantFloor(targetUserId ? data.tenant : null);
      setPolicy(targetUserId ? (data.user ?? null) : data.tenant);
    } catch (err: unknown) {
      setPolicy(null);
      setStatusMsg({ kind: 'error', text: err instanceof Error ? err.message : String(err) });
    }
  }, [userId, targetUserId]);

  useEffect(() => { load(); }, [load]);

  const handleChange = async (channel: Channel, mode: Mode) => {
    setSaving(channel);
    setStatusMsg(null);
    try {
      const res = await fetch('/api/manage/outbound-policy', {
        method: 'POST',
        headers: { 'x-user-id': userId, 'Content-Type': 'application/json' },
        body: JSON.stringify(targetUserId
          ? { scope: 'user', userId: targetUserId, [channel]: mode }
          : { scope: 'tenant', [channel]: mode }),
      });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      const data = (await res.json()) as { policy: Policy };
      setPolicy(data.policy);
      setStatusMsg({ kind: 'ok', text: 'Saved.' });
      setTimeout(() => setStatusMsg(null), 2000);
    } catch (err: unknown) {
      setStatusMsg({ kind: 'error', text: err instanceof Error ? err.message : String(err) });
      setTimeout(() => setStatusMsg(null), 4000);
    } finally {
      setSaving(null);
    }
  };

  return (
    <>
      {!targetUserId && (
        <p className="section-hint">
          Enforced draft mode holds email. These channels deliver as soon as the tool runs, so
          there is nothing to hold: an agent's call is either allowed, allowed only when every
          recipient is inside the organization (the tenant's verified domains), or refused with a
          message telling the user to act in Outlook or Teams. Refusals are written to the audit
          log. A per-user setting under User Management can tighten this but not loosen it.
        </p>
      )}
      {policy === null ? (
        <p style={{ color: '#666', fontSize: 13 }}>{statusMsg ? '' : 'Loading policy…'}</p>
      ) : (
        <div className="services-list" style={targetUserId ? { marginTop: 16 } : undefined}>
          {CHANNELS.map((c) => {
            const serviceOff = !enabledServices.includes(c.service);
            const floor = tenantFloor?.[c.key] ?? 'allow';
            const effective: Mode = RANK[floor] >= RANK[policy[c.key]] ? floor : policy[c.key];
            return (
              <div key={c.key} className={`service-row${serviceOff ? ' service-row--tenant-disabled' : ''}`}>
                <span className="service-info">
                  <span className="service-name">{c.label}</span>
                  <span className="service-desc">
                    {c.desc}
                    {targetUserId && floor !== 'allow' && ` Organization setting: ${MODE_LABEL[floor]}.`}
                  </span>
                </span>
                <select
                  value={policy[c.key]}
                  disabled={saving !== null}
                  onChange={(e) => handleChange(c.key, e.target.value as Mode)}
                  aria-label={c.label}
                >
                  {(['allow', 'internal', 'block'] as Mode[]).map((m) => (
                    <option key={m} value={m} disabled={targetUserId !== undefined && RANK[m] < RANK[floor]}>
                      {MODE_LABEL[m]}
                    </option>
                  ))}
                </select>
                {effective !== 'allow' && (
                  <span className="override-badge">
                    {targetUserId && RANK[floor] >= RANK[policy[c.key]] ? 'Tenant' : MODE_LABEL[effective]}
                  </span>
                )}
              </div>
            );
          })}
        </div>
      )}
      {statusMsg && (
        <p className={`status-msg status-msg--${statusMsg.kind}`}>{statusMsg.text}</p>
      )}
    </>
  );
}
