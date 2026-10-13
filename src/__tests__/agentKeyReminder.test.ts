/**
 * The calendar reminder written when an agent key is minted
 * (services/agentKeyReminder.ts).
 *
 * Covers:
 *   - lead time by key lifetime: 7 days, 2 days, or none for very short keys
 *   - the event: owner's default calendar, no attendees, free, reminder on,
 *     start in the mailbox's own zone, subject and body name the key
 *   - a Graph response without an id is an error, so the caller never
 *     records a reminder that does not exist
 *   - deletion treats an already-gone event (404) as success
 */
import { jest } from '@jest/globals';

const mockPost = jest.fn<(body: unknown) => Promise<unknown>>();
const mockDelete = jest.fn<() => Promise<unknown>>();
const mockApi = jest.fn((path: string) => ({ post: mockPost, delete: mockDelete, path }));
const mockResolveTz = jest.fn<() => Promise<string>>();

jest.mock('../services/graphClient.js', () => ({ createGraphClient: () => ({ api: (p: string) => mockApi(p) }) }));
jest.mock('../services/tokenCache.js', () => ({ getValidAccessTokenForSession: async () => 'at' }));
jest.mock('../services/mailboxTimeZone.js', () => ({ resolveMailboxTimeZone: () => mockResolveTz() }));
jest.mock('../services/frontendUrl.js', () => ({ resolveFrontendUrl: () => 'https://mcp.example.test/' }));

import {
  createKeyExpiryReminder,
  deleteKeyExpiryReminder,
  reminderEventBody,
  reminderLeadDays,
  reminderStart,
} from '../services/agentKeyReminder.js';
import type { AgentKeyRecord } from '../services/agentKeys.js';
import type { UserSession } from '../services/tokenCache.js';

const DAY = 24 * 60 * 60 * 1000;
const session = {} as UserSession;
function record(days: number, over: Partial<AgentKeyRecord> = {}): AgentKeyRecord {
  const createdAt = Date.UTC(2026, 9, 13, 20, 0, 0);
  return {
    keyId: 'abc123def456', tenantId: 't', userId: 'u', userEmail: 'adele@fabrikam.com', label: 'nightly',
    sessionStorageKey: 'row', policy: { services: ['sharepoint', 'mail'], readOnlyServices: ['mail'], allowedSites: [] },
    createdAt, expiresAt: createdAt + days * DAY, lastUsedAt: 0, revokedAt: 0, ...over,
  };
}

beforeEach(() => {
  jest.clearAllMocks();
  mockPost.mockResolvedValue({ id: 'evt-1' });
  mockDelete.mockResolvedValue(undefined);
  mockResolveTz.mockResolvedValue('Pacific Standard Time');
});

describe('lead time', () => {
  it('is a week for keys of two weeks or more, two days above four days, none below', () => {
    expect(reminderLeadDays(365 * DAY)).toBe(7);
    expect(reminderLeadDays(14 * DAY)).toBe(7);
    expect(reminderLeadDays(13 * DAY)).toBe(2);
    expect(reminderLeadDays(4 * DAY)).toBe(2);
    expect(reminderLeadDays(3 * DAY)).toBe(0);
  });
  it('places the start at 09:00 on the lead date, or none', () => {
    expect(reminderStart(record(90))).toBe('2027-01-04T09:00:00'); // expiry 2027-01-11 minus 7
    expect(reminderStart(record(5))).toBe('2026-10-16T09:00:00');  // expiry 10-18 minus 2
    expect(reminderStart(record(2))).toBeNull();
  });
});

describe('createKeyExpiryReminder', () => {
  it('writes a free, attendee-less event with a reminder on the owner’s default calendar', async () => {
    const created = await createKeyExpiryReminder(session, record(90));
    expect(created).toEqual({ eventId: 'evt-1', timeZone: 'Pacific Standard Time', timeZoneFallback: false, fallbackReason: undefined });
    expect(mockApi).toHaveBeenCalledWith('/me/events');
    const event = mockPost.mock.calls[0][0] as Record<string, unknown>;
    expect(event).toMatchObject({
      subject: 'Agent key "nightly" expires soon',
      start: { dateTime: '2027-01-04T09:00:00', timeZone: 'Pacific Standard Time' },
      end: { dateTime: '2027-01-04T09:30:00', timeZone: 'Pacific Standard Time' },
      isAllDay: false, showAs: 'free', isReminderOn: true, reminderMinutesBeforeStart: 0,
    });
    expect(event.attendees).toBeUndefined();
    const body = (event.body as { content: string }).content;
    expect(body).toContain('abc123def456');
    expect(body).toContain('sharepoint, mail');
    expect(body).toContain('read-only: mail');
    expect(body).toContain('https://mcp.example.test/');
    expect(body).toContain('My Settings');
    expect(body).not.toContain('UTC');
  });

  it('falls back to 16:00 UTC, says so in the body, and reports it when the mailbox zone cannot be read', async () => {
    mockResolveTz.mockRejectedValue(new Error('Access is denied. Check credentials and try again.'));
    const created = await createKeyExpiryReminder(session, record(90));
    expect(created).toMatchObject({ eventId: 'evt-1', timeZone: 'UTC', timeZoneFallback: true, fallbackReason: expect.stringContaining('Access is denied') });
    const event = mockPost.mock.calls[0][0] as Record<string, unknown>;
    expect(event).toMatchObject({
      start: { dateTime: '2027-01-04T16:00:00', timeZone: 'UTC' },
      end: { dateTime: '2027-01-04T16:30:00', timeZone: 'UTC' },
    });
    expect((event.body as { content: string }).content).toContain('MailboxSettings.Read');
  });

  it('returns null without calling Graph for a key too short to remind about', async () => {
    expect(await createKeyExpiryReminder(session, record(2))).toBeNull();
    expect(mockPost).not.toHaveBeenCalled();
  });

  it('throws when Graph returns no event id, so nothing false is recorded', async () => {
    mockPost.mockResolvedValue({});
    await expect(createKeyExpiryReminder(session, record(30))).rejects.toThrow(/event id/);
  });

  it('names the key in the body even when the frontend URL is unknown', () => {
    const { subject, content } = reminderEventBody(record(30), '');
    expect(subject).toContain('nightly');
    expect(content).not.toContain('Open:');
  });
});

describe('deleteKeyExpiryReminder', () => {
  it('deletes by id and ignores an event that is already gone', async () => {
    await deleteKeyExpiryReminder(session, 'evt-1');
    expect(mockApi).toHaveBeenCalledWith('/me/events/evt-1');
    mockDelete.mockRejectedValue({ statusCode: 404 });
    await expect(deleteKeyExpiryReminder(session, 'evt-1')).resolves.toBeUndefined();
    mockDelete.mockRejectedValue({ statusCode: 500 });
    await expect(deleteKeyExpiryReminder(session, 'evt-1')).rejects.toBeDefined();
  });
});
