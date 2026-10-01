/**
 * Unit tests for resolveMailboxTimeZone.
 *
 * create_event / update_event default a missing timeZone to the mailbox's own
 * configured zone instead of a hardcoded America/New_York. When the zone can't
 * be resolved the resolver throws rather than guessing.
 */

import { resolveMailboxTimeZone } from '../services/mailboxTimeZone.js';
import type { Client } from '@microsoft/microsoft-graph-client';

function graphReturning(settings: unknown, capture?: { path?: string }): Client {
  return {
    api(path: string) {
      if (capture) capture.path = path;
      return {
        select() {
          return this;
        },
        get() {
          return Promise.resolve(settings);
        },
      };
    },
  } as unknown as Client;
}

describe('resolveMailboxTimeZone', () => {
  it('returns the mailbox timeZone (Windows zone name) from /me/mailboxSettings', async () => {
    const capture: { path?: string } = {};
    const tz = await resolveMailboxTimeZone(graphReturning({ timeZone: 'Pacific Standard Time' }, capture));
    expect(tz).toBe('Pacific Standard Time');
    expect(capture.path).toBe('/me/mailboxSettings');
  });

  it('targets an explicit mailbox path when supplied', async () => {
    const capture: { path?: string } = {};
    await resolveMailboxTimeZone(graphReturning({ timeZone: 'Eastern Standard Time' }, capture), '/users/bob@x.com');
    expect(capture.path).toBe('/users/bob@x.com/mailboxSettings');
  });

  it('throws (never guesses) when the mailbox has no configured time zone', async () => {
    await expect(resolveMailboxTimeZone(graphReturning({}))).rejects.toThrow(/Could not determine a time zone/);
  });

  it('throws when timeZone is not a string', async () => {
    await expect(resolveMailboxTimeZone(graphReturning({ timeZone: 123 }))).rejects.toThrow(/Could not determine a time zone/);
  });
});
