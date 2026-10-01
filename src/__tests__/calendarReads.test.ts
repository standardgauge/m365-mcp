/**
 * Calendar-read fixes.
 *
 * The M365 connector's three calendar-read verbs were unusable for a
 * specific-day availability check:
 *   - list_events queried /events (recurring series masters) instead of
 *     /calendarView (expanded instances), so it could never answer "what is on
 *     this specific day" and returned events dated years in the past.
 *   - get_schedule fed loosely-formed datetimes to Graph and returned
 *     "FreeBusyViewOptions.TimeWindow is invalid".
 *   - find_meeting_times only applied its time window when BOTH bounds were
 *     present, silently falling back to today otherwise.
 *
 * Two layers:
 *   1. The pure window helpers in services/calendarWindow.ts.
 *   2. The MCP tool handlers (/api/mcp JSON-RPC), asserting the Graph call
 *      shape (calendarView vs /events, Prefer header, timeConstraint).
 *
 * All Graph, auth, and storage dependencies are mocked.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

import { normalizeGraphDateTime, shiftDays, resolveWindow } from '../services/calendarWindow.js';

// ── Pure helper unit tests (no mocks needed) ─────────────────────────────────

describe('calendarWindow — normalizeGraphDateTime', () => {
  it('expands a bare start date to the beginning of the day', () => {
    expect(normalizeGraphDateTime('2026-09-08', 'start')).toBe('2026-09-08T00:00:00');
  });
  it('expands a bare end date to the end of the day', () => {
    expect(normalizeGraphDateTime('2026-09-08', 'end')).toBe('2026-09-08T23:59:59');
  });
  it('strips a trailing Z that contradicts an explicit timeZone', () => {
    expect(normalizeGraphDateTime('2026-09-08T10:00:00Z', 'start')).toBe('2026-09-08T10:00:00');
  });
  it('strips a numeric UTC offset', () => {
    expect(normalizeGraphDateTime('2026-09-08T10:00:00-07:00', 'start')).toBe('2026-09-08T10:00:00');
  });
  it('leaves a well-formed naive-local datetime untouched', () => {
    expect(normalizeGraphDateTime('2026-09-08T09:30:00', 'end')).toBe('2026-09-08T09:30:00');
  });
});

describe('calendarWindow — shiftDays', () => {
  it('shifts forward by whole days', () => {
    expect(shiftDays('2026-09-08T00:00:00', 7)).toBe('2026-09-15T00:00:00');
  });
  it('shifts backward by whole days', () => {
    expect(shiftDays('2026-09-08T23:59:59', -7)).toBe('2026-09-01T23:59:59');
  });
  it('returns the input unchanged when it cannot be parsed', () => {
    expect(shiftDays('not-a-date', 7)).toBe('not-a-date');
  });
});

describe('calendarWindow — resolveWindow', () => {
  it('returns null when neither bound is supplied', () => {
    expect(resolveWindow(undefined, undefined)).toBeNull();
  });
  it('normalizes both bounds when both are supplied', () => {
    expect(resolveWindow('2026-09-08', '2026-09-10')).toEqual({
      start: '2026-09-08T00:00:00',
      end: '2026-09-10T23:59:59',
    });
  });
  it('defaults the end 7 days out when only the start is given', () => {
    expect(resolveWindow('2026-09-08T00:00:00', undefined)).toEqual({
      start: '2026-09-08T00:00:00',
      end: '2026-09-15T00:00:00',
    });
  });
  it('defaults the start 7 days back when only the end is given', () => {
    expect(resolveWindow(undefined, '2026-09-08T00:00:00')).toEqual({
      start: '2026-09-01T00:00:00',
      end: '2026-09-08T00:00:00',
    });
  });
});

// ── Mock declarations for the endpoint-handler tests ─────────────────────────

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();
const mockGetValidAccessTokenForSession = jest.fn<() => Promise<string>>();
const mockGetTenantIdFromSession = jest.fn<() => string>();
const mockIsPathDenied = jest.fn<(t: string, u: string, type: string, path: string) => Promise<boolean>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<() => Promise<Array<{ id: string; name: string }>>>();
const mockGetUserServiceOverrides = jest.fn<() => Promise<string[]>>();
const mockResolveDefaultCalendarId = jest.fn<() => Promise<string | null>>();

const lastGraphCall: { path: string | null; postBody: unknown; headers: Record<string, string> } = {
  path: null, postBody: null, headers: {},
};
const mockGraphGet = jest.fn<() => Promise<unknown>>();
const mockGraphPost = jest.fn<(body?: unknown) => Promise<unknown>>();
function makeChain() {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand']) chain[m] = () => chain;
  chain.header = (k: string, v: string) => { lastGraphCall.headers[k] = v; return chain; };
  chain.get = (...a: unknown[]) => mockGraphGet(...(a as []));
  chain.post = (body?: unknown) => { lastGraphCall.postBody = body; return mockGraphPost(body); };
  chain.patch = () => Promise.resolve({});
  chain.delete = () => Promise.resolve(undefined);
  return chain;
}
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    lastGraphCall.path = path;
    return makeChain();
  },
}));

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({ authenticateRequest: () => mockAuthenticateRequest() }));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => mockGetValidAccessTokenForSession(),
  getTenantIdFromSession: () => mockGetTenantIdFromSession(),
}));
jest.mock('../services/graphClient.js', () => ({ createGraphClient: () => mockCreateGraphClient() }));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: (...args: unknown[]) =>
    mockIsPathDenied(args[0] as string, args[1] as string, args[2] as string, args[3] as string),
}));
jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => Promise.resolve(null),
  resolveDefaultCalendarId: () => mockResolveDefaultCalendarId(),
  resolveCalendarName: () => Promise.resolve('My Calendar'),
  resolveContactParentFolder: () => Promise.resolve(null),
  resolveDefaultContactFolder: () => Promise.resolve(null),
  resolveSectionNotebook: () => Promise.resolve(null),
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
}));
jest.mock('../services/sharepointFilter.js', () => ({ filterAndDisambiguateSites: (sites: unknown) => sites }));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => mockGetEnabledServices(),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => mockGetAllowedSites(),
}));
jest.mock('../services/userServiceOverrides.js', () => ({ getUserServiceOverrides: () => mockGetUserServiceOverrides() }));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'draft' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));
jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<() => Promise<boolean>>().mockResolvedValue(false),
}));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

// ── Import after mocks ──────────────────────────────────────────────────────

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint');
if (!registration) throw new Error('mcpEndpoint handler was not registered');
const handler = registration[1].handler;

const TENANT = 'test-tenant';
const USER = 'test-user';
const AUTH = { userId: USER, session: { userId: USER, tenantId: TENANT, accessToken: 'fake', sessionToken: 'sess' } };

function callTool(name: string, args: Record<string, unknown>): Promise<{ status: number; jsonBody?: unknown }> {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

function toolResult(res: { jsonBody?: unknown }): { parsed: unknown; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

beforeEach(() => {
  jest.clearAllMocks();
  lastGraphCall.path = null;
  lastGraphCall.postBody = null;
  lastGraphCall.headers = {};
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['calendar', 'mail']);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
  mockResolveDefaultCalendarId.mockResolvedValue('default-cal-id');
  mockGraphGet.mockResolvedValue({ value: [] });
  mockGraphPost.mockResolvedValue({});
});

// ── list_events ──────────────────────────────────────────────────────────────

describe('MCP tools/call — list_events', () => {
  it('uses calendarView with an expanded, normalized window when a date range is given', async () => {
    await callTool('list_events', { startDateTime: '2026-09-08', endDateTime: '2026-09-10' });
    expect(lastGraphCall.path).toContain('/me/calendarView');
    expect(lastGraphCall.path).toContain(`startDateTime=${encodeURIComponent('2026-09-08T00:00:00')}`);
    expect(lastGraphCall.path).toContain(`endDateTime=${encodeURIComponent('2026-09-10T23:59:59')}`);
    // Do NOT hit /events, which returns unexpanded series masters.
    expect(lastGraphCall.path).not.toContain('/me/events');
  });

  it('sends the Prefer: outlook.timezone hint so results come back in the requested zone', async () => {
    await callTool('list_events', { startDateTime: '2026-09-08T09:00:00', endDateTime: '2026-09-08T17:00:00' });
    expect(lastGraphCall.headers['Prefer']).toBe('outlook.timezone="America/Los_Angeles"');
  });

  it('honors a custom timeZone in the Prefer hint', async () => {
    await callTool('list_events', { startDateTime: '2026-09-08', endDateTime: '2026-09-09', timeZone: 'America/New_York' });
    expect(lastGraphCall.headers['Prefer']).toBe('outlook.timezone="America/New_York"');
  });

  it('targets the calendar-scoped calendarView when a calendarId is given', async () => {
    await callTool('list_events', { calendarId: 'AAMkCal123', startDateTime: '2026-09-08', endDateTime: '2026-09-09' });
    expect(lastGraphCall.path).toContain('/me/calendars/AAMkCal123/calendarView');
  });

  it('falls back to /me/events (no window) when no date range is given', async () => {
    await callTool('list_events', {});
    expect(lastGraphCall.path).toBe('/me/events');
  });
});

// ── get_schedule ──────────────────────────────────────────────────────────────

describe('MCP tools/call — get_schedule', () => {
  it('normalizes a bare-date window and sends the Prefer hint', async () => {
    mockGraphPost.mockResolvedValue({ value: [] });
    await callTool('get_schedule', { schedules: ['a@example.com'], startDateTime: '2026-09-08', endDateTime: '2026-09-08' });
    expect(lastGraphCall.path).toBe('/me/calendar/getSchedule');
    const post = lastGraphCall.postBody as { startTime: { dateTime: string }; endTime: { dateTime: string }; availabilityViewInterval: number };
    expect(post.startTime.dateTime).toBe('2026-09-08T00:00:00');
    expect(post.endTime.dateTime).toBe('2026-09-08T23:59:59');
    expect(post.availabilityViewInterval).toBe(30);
    expect(lastGraphCall.headers['Prefer']).toContain('outlook.timezone');
  });

  it('strips an absolute-time marker that would contradict the timeZone field', async () => {
    mockGraphPost.mockResolvedValue({ value: [] });
    await callTool('get_schedule', { schedules: ['a@example.com'], startDateTime: '2026-09-08T09:00:00Z', endDateTime: '2026-09-08T17:00:00Z' });
    const post = lastGraphCall.postBody as { startTime: { dateTime: string }; endTime: { dateTime: string } };
    expect(post.startTime.dateTime).toBe('2026-09-08T09:00:00');
    expect(post.endTime.dateTime).toBe('2026-09-08T17:00:00');
  });

  it('clamps an out-of-range availabilityViewInterval into [5, 1440]', async () => {
    mockGraphPost.mockResolvedValue({ value: [] });
    await callTool('get_schedule', { schedules: ['a@example.com'], startDateTime: '2026-09-08T09:00:00', endDateTime: '2026-09-08T17:00:00', availabilityViewInterval: 99999 });
    const post = lastGraphCall.postBody as { availabilityViewInterval: number };
    expect(post.availabilityViewInterval).toBe(1440);
  });

  it('errors (no Graph call) when a bound is missing', async () => {
    const res = await callTool('get_schedule', { schedules: ['a@example.com'], startDateTime: '2026-09-08T09:00:00' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});

// ── find_meeting_times ────────────────────────────────────────────────────────

describe('MCP tools/call — find_meeting_times', () => {
  it('applies a timeConstraint when only one bound is given (defaults the other)', async () => {
    mockGraphPost.mockResolvedValue({ meetingTimeSuggestions: [] });
    await callTool('find_meeting_times', { attendees: ['a@example.com'], startDateTime: '2026-09-08T09:00:00' });
    const post = lastGraphCall.postBody as { timeConstraint?: { timeSlots: Array<{ start: { dateTime: string }; end: { dateTime: string } }> } };
    expect(post.timeConstraint).toBeDefined();
    expect(post.timeConstraint!.timeSlots[0].start.dateTime).toBe('2026-09-08T09:00:00');
    expect(post.timeConstraint!.timeSlots[0].end.dateTime).toBe('2026-09-15T09:00:00');
  });

  it('normalizes bare dates in the timeConstraint window', async () => {
    mockGraphPost.mockResolvedValue({ meetingTimeSuggestions: [] });
    await callTool('find_meeting_times', { attendees: ['a@example.com'], startDateTime: '2026-09-08', endDateTime: '2026-09-10' });
    const post = lastGraphCall.postBody as { timeConstraint?: { timeSlots: Array<{ start: { dateTime: string }; end: { dateTime: string } }> } };
    expect(post.timeConstraint!.timeSlots[0].start.dateTime).toBe('2026-09-08T00:00:00');
    expect(post.timeConstraint!.timeSlots[0].end.dateTime).toBe('2026-09-10T23:59:59');
  });

  it('omits the timeConstraint when no bound is given', async () => {
    mockGraphPost.mockResolvedValue({ meetingTimeSuggestions: [] });
    await callTool('find_meeting_times', { attendees: ['a@example.com'] });
    const post = lastGraphCall.postBody as { timeConstraint?: unknown };
    expect(post.timeConstraint).toBeUndefined();
  });
});
