/**
 * Resolve the default time zone for a calendar write.
 *
 * create_event / update_event must never guess a time zone. Before this, both
 * defaulted a missing `timeZone` to a hardcoded `America/New_York`, so every
 * event a Pacific user created without an explicit zone landed three hours
 * early — silently, because the response looked successful. A parent shows up
 * at 2:00pm for a 5:00pm school event; a client call is worse.
 *
 * The correct default is the mailbox's OWN configured zone
 * (mailboxSettings.timeZone), which Exchange returns as a Windows zone name
 * ("Pacific Standard Time"). Graph's event API accepts both Windows and IANA
 * zone names in start/end.timeZone, so the value flows straight through and the
 * echoed timeZone round-trips.
 *
 * Reading /me/mailboxSettings needs the MailboxSettings.Read delegated scope,
 * which is granted at the app-registration level (see graphClient.ts and
 * docs/entra-setup.md) — the same scope find_meeting_times already relies on.
 *
 * If the zone genuinely cannot be resolved, this throws rather than falling
 * back to a guess: a loud failure the caller can fix by passing an explicit
 * timeZone beats another wrong-time event written into a live calendar.
 */

import { Client } from '@microsoft/microsoft-graph-client';

/**
 * Fetch the mailbox's configured time zone. `mailboxPath` is the Graph base for
 * the target mailbox (`/me` for the signed-in user, `/users/{id}` otherwise);
 * it MUST already be a trusted/validated value, as it is interpolated into the
 * Graph path.
 *
 * @throws when mailboxSettings.timeZone is absent or not a string.
 */
export async function resolveMailboxTimeZone(graph: Client, mailboxPath = '/me'): Promise<string> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const settings: any = await graph.api(`${mailboxPath}/mailboxSettings`).select('timeZone').get();
  const tz = settings?.timeZone;
  if (!tz || typeof tz !== 'string') {
    throw new Error(
      'Could not determine a time zone for this calendar write: no explicit timeZone was supplied and the ' +
      'mailbox has no configured time zone (mailboxSettings.timeZone). Pass an explicit timeZone ' +
      "(e.g. \"America/Los_Angeles\") and retry.",
    );
  }
  return tz;
}
