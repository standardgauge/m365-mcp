import { createGraphClient } from './graphClient.js';
import { getValidAccessTokenForSession } from './tokenCache.js';
import type { UserSession } from './tokenCache.js';
import { resolveMailboxTimeZone } from './mailboxTimeZone.js';
import { resolveFrontendUrl } from './frontendUrl.js';
import type { AgentKeyRecord } from './agentKeys.js';

/**
 * A calendar reminder before an agent key expires.
 *
 * The server already holds the owner's delegated token, so when a key is
 * minted it can put a short event on the owner's own calendar a few days
 * before the expiry it just recorded. Nobody else is invited, so nothing is
 * sent and the outbound policy is not involved; the event is the owner's note
 * to themselves, written at their explicit request on the key form. It is best
 * effort: a failure is reported back and audited, and the key is created
 * regardless.
 */

const DAY_MS = 24 * 60 * 60 * 1000;
/** Local wall-clock hour the reminder event starts. */
const REMINDER_HOUR = 9;
const REMINDER_LENGTH_MIN = 30;

/**
 * How far before expiry the reminder lands: a week for keys that live at
 * least two weeks, two days for keys of four days or more, and none at all
 * for shorter keys, where a reminder would arrive before the owner had
 * finished wiring the key in.
 */
export function reminderLeadDays(lifetimeMs: number): number {
  if (lifetimeMs >= 14 * DAY_MS) return 7;
  if (lifetimeMs >= 4 * DAY_MS) return 2;
  return 0;
}

/** The reminder's start as a wall-clock string (no zone), or null when no reminder applies. */
export function reminderStart(record: Pick<AgentKeyRecord, 'createdAt' | 'expiresAt'>): string | null {
  const lead = reminderLeadDays(record.expiresAt - record.createdAt);
  if (lead === 0) return null;
  // The calendar date is taken in UTC and sent with the mailbox's own zone,
  // so near midnight it can land a day either side. For a lead of days that
  // does not matter, and it avoids mapping Windows zone names to IANA here.
  const d = new Date(record.expiresAt - lead * DAY_MS);
  const ymd = d.toISOString().slice(0, 10);
  return `${ymd}T${String(REMINDER_HOUR).padStart(2, '0')}:00:00`;
}

function addMinutes(wallClock: string, minutes: number): string {
  const [date, time] = wallClock.split('T');
  const [h, m] = time.split(':').map(Number);
  const total = h * 60 + m + minutes;
  return `${date}T${String(Math.floor(total / 60) % 24).padStart(2, '0')}:${String(total % 60).padStart(2, '0')}:00`;
}

export function reminderEventBody(record: AgentKeyRecord, frontendUrl: string): { subject: string; content: string } {
  const expires = new Date(record.expiresAt).toUTCString();
  const services = record.policy.services.join(', ') || 'none';
  const ro = record.policy.readOnlyServices.length ? ` (read-only: ${record.policy.readOnlyServices.join(', ')})` : '';
  return {
    subject: `Agent key "${record.label}" expires soon`,
    content: [
      `The agent key "${record.label}" (id ${record.keyId}) expires on ${expires}.`,
      `It may use: ${services}${ro}.`,
      '',
      'To keep whatever uses it running: mint a replacement under My Settings, Agent Keys, switch the app to the new key, then revoke this one.',
      frontendUrl ? `Open: ${frontendUrl}` : '',
    ].filter((l) => l !== undefined).join('\n'),
  };
}

/**
 * Create the reminder on the owner's default calendar. Returns the Graph event
 * id, or null when the key is too short-lived for one. Throws on a Graph
 * failure; the caller decides that is not fatal.
 */
export async function createKeyExpiryReminder(session: UserSession, record: AgentKeyRecord): Promise<string | null> {
  const start = reminderStart(record);
  if (!start) return null;
  const token = await getValidAccessTokenForSession(session);
  const graph = createGraphClient(token);
  const tz = await resolveMailboxTimeZone(graph);
  const { subject, content } = reminderEventBody(record, resolveFrontendUrl());
  const event = {
    subject,
    body: { contentType: 'Text', content },
    start: { dateTime: start, timeZone: tz },
    end: { dateTime: addMinutes(start, REMINDER_LENGTH_MIN), timeZone: tz },
    isAllDay: false,
    showAs: 'free',
    isReminderOn: true,
    reminderMinutesBeforeStart: 0,
    categories: ['Agent key'],
  };
  const created = (await graph.api('/me/events').post(event)) as { id?: string };
  if (!created?.id) throw new Error('Graph returned no event id');
  return created.id;
}

/** Remove a reminder, as the owner. 404 (already gone) is not an error. */
export async function deleteKeyExpiryReminder(session: UserSession, eventId: string): Promise<void> {
  const token = await getValidAccessTokenForSession(session);
  const graph = createGraphClient(token);
  try {
    await graph.api(`/me/events/${encodeURIComponent(eventId)}`).delete();
  } catch (err: unknown) {
    const status = (err as { statusCode?: number } | null)?.statusCode;
    if (status !== 404) throw err;
  }
}
