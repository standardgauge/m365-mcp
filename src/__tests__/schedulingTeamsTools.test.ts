/**
 * Scheduling + Teams MCP exposure tests.
 *
 * Covers the tools added to enable AI scheduling and Teams messaging:
 *   - Calendar scheduling: get_schedule, find_meeting_times, list_rooms
 *   - create_event online-meeting enhancement (isOnlineMeeting → onlineMeetingUrl)
 *   - Teams messaging + discovery: list_teams, list_channels, send_chat_message,
 *     send_channel_message
 *
 * Two layers, mirroring calendarTools.test.ts:
 *   1. toolManifest — the new tools are registered in the TOOLS catalog.
 *   2. Remote MCP endpoint (/api/mcp JSON-RPC) — the tools surface in tools/list
 *      (gated by enabledServices) and are invokable via tools/call with deny-list
 *      and opaque-ID enforcement intact.
 *
 * All Graph, auth, and storage dependencies are mocked.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── Mock declarations ────────────────────────────────────────────────────────

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

// telemetry.js has import-time side effects (patches console) — stub it out
jest.mock('../services/telemetry.js', () => ({}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => mockAuthenticateRequest(),
}));

jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => mockGetValidAccessTokenForSession(),
  getTenantIdFromSession: () => mockGetTenantIdFromSession(),
}));

jest.mock('../services/mailboxTimeZone.js', () => ({
  resolveMailboxTimeZone: jest.fn(async () => 'Pacific Standard Time'),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => mockCreateGraphClient(),
}));

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

jest.mock('../services/sharepointFilter.js', () => ({
  filterAndDisambiguateSites: (sites: unknown) => sites,
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => mockGetEnabledServices(),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => mockGetAllowedSites(),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: () => mockGetUserServiceOverrides(),
}));

jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'draft' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<() => Promise<boolean>>().mockResolvedValue(false),
}));

jest.mock('../services/auditLog.js', () => ({
  logAccess: jest.fn(),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

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

function rpc(method: string, params: Record<string, unknown>): Promise<{ status: number; jsonBody?: unknown }> {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method, params }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

const callTool = (name: string, args: Record<string, unknown>) => rpc('tools/call', { name, arguments: args });

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
  mockGetEnabledServices.mockResolvedValue(['calendar', 'mail', 'sharepoint', 'teams']);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
  mockResolveDefaultCalendarId.mockResolvedValue('default-cal-id');
  mockGraphGet.mockResolvedValue({ value: [] });
  mockGraphPost.mockResolvedValue({});
});

// ─────────────────────────────────────────────────────────────────────────────
// toolManifest — new tools registered
// ─────────────────────────────────────────────────────────────────────────────

describe('toolManifest — scheduling + teams tools', () => {
  const NEW_TOOLS = ['get_schedule', 'find_meeting_times', 'list_rooms', 'list_teams', 'list_channels', 'send_chat_message', 'send_channel_message'];

  it('TOOLS contains every new scheduling + teams tool', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const names = TOOLS.map((t) => t.name);
    for (const t of NEW_TOOLS) expect(names).toContain(t);
  });

  // JSON-RPC-only tools carry endpoint '/api/mcp' (POST); list_teams / list_channels
  // additionally have a real REST Azure Function and keep their /api/teams/ route.
  it('JSON-RPC-only tools point at /api/mcp; REST-backed tools keep their real route', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const byName = Object.fromEntries(TOOLS.map((t) => [t.name, t]));
    for (const t of ['get_schedule', 'find_meeting_times', 'list_rooms', 'send_chat_message', 'send_channel_message']) {
      expect(byName[t].endpoint).toBe('/api/mcp');
      expect(byName[t].method).toBe('POST');
    }
    expect(byName['list_teams'].endpoint).toBe('/api/teams/teams');
    expect(byName['list_teams'].method).toBe('GET');
    expect(byName['list_channels'].endpoint).toBe('/api/teams/teams/{teamId}/channels');
    expect(byName['list_channels'].method).toBe('GET');
  });

  // Guard the finding that started this: the manifest must not advertise a REST route
  // (/api/calendar/... or /api/teams/.../messages) that has no Azure Function behind it.
  it('no new tool advertises a non-existent REST endpoint', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const byName = Object.fromEntries(TOOLS.map((t) => [t.name, t]));
    const PHANTOM_ROUTES = [
      '/api/calendar/getSchedule',
      '/api/calendar/findMeetingTimes',
      '/api/calendar/rooms',
      '/api/teams/chats/{chatId}/messages',
      '/api/teams/teams/{teamId}/channels/{channelId}/messages',
    ];
    for (const t of NEW_TOOLS) {
      expect(PHANTOM_ROUTES).not.toContain(byName[t].endpoint);
    }
  });

  it('create_event exposes isOnlineMeeting + onlineMeetingProvider params', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const createEvent = TOOLS.find((t) => t.name === 'create_event');
    const paramNames = (createEvent?.parameters ?? []).map((p) => p.name);
    expect(paramNames).toContain('isOnlineMeeting');
    expect(paramNames).toContain('onlineMeetingProvider');
  });

  it('getManifest() output includes the new tools', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string }> };
    const names = manifest.tools.map((t) => t.name);
    for (const t of NEW_TOOLS) expect(names).toContain(t);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Schema sync — the manifest catalog (getManifest) and the live tools/list
// surface must not drift. This is the regression guard for the
// review finding: the manifest advertised `schedules` / `attendees` as strings
// while the JSON-RPC handlers require arrays, so a client trusting the catalog
// would send a payload the handler rejects.
// ─────────────────────────────────────────────────────────────────────────────

describe('schema sync — manifest catalog vs live tools/list', () => {
  const NEW_TOOLS = ['get_schedule', 'find_meeting_times', 'list_rooms', 'list_teams', 'list_channels', 'send_chat_message', 'send_channel_message'];

  // Structure only — property names + JSON types + array item types + enum + required.
  // Human-facing descriptions are allowed to differ between the two surfaces.
  interface SchemaShape { type: unknown; required: string[]; props: Record<string, unknown> }
  function shape(inputSchema: any): SchemaShape {
    const props: Record<string, unknown> = {};
    for (const [name, spec] of Object.entries(inputSchema.properties ?? {})) {
      const s = spec as Record<string, unknown>;
      props[name] = {
        type: s.type,
        ...(s.items ? { items: { type: (s.items as Record<string, unknown>).type } } : {}),
        ...(s.enum ? { enum: s.enum } : {}),
      };
    }
    return { type: inputSchema.type, required: [...(inputSchema.required ?? [])].sort(), props };
  }

  it('every new tool has the same input-schema shape in getManifest() and tools/list', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string; inputSchema: any }> };
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string; inputSchema: any }> } };

    const manifestByName = Object.fromEntries(manifest.tools.map((t) => [t.name, t]));
    const liveByName = Object.fromEntries(body.result.tools.map((t) => [t.name, t]));

    for (const name of NEW_TOOLS) {
      expect(manifestByName[name]).toBeDefined();
      expect(liveByName[name]).toBeDefined();
      expect(shape(manifestByName[name].inputSchema)).toEqual(shape(liveByName[name].inputSchema));
    }
  });

  it('schedules / attendees are arrays of strings in the manifest, matching the handlers', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string; inputSchema: any }> };
    const byName = Object.fromEntries(manifest.tools.map((t) => [t.name, t]));
    expect(byName['get_schedule'].inputSchema.properties.schedules).toMatchObject({ type: 'array', items: { type: 'string' } });
    expect(byName['find_meeting_times'].inputSchema.properties.attendees).toMatchObject({ type: 'array', items: { type: 'string' } });
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/list — visibility gated by enabledServices
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/list — scheduling + teams exposure', () => {
  it('lists scheduling tools when calendar is enabled and teams tools when teams is enabled', async () => {
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    for (const t of ['get_schedule', 'find_meeting_times', 'list_rooms', 'list_teams', 'list_channels', 'send_chat_message', 'send_channel_message']) {
      expect(names).toContain(t);
    }
  });

  it('hides scheduling tools when calendar is disabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail', 'teams']);
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    expect(names).not.toContain('get_schedule');
    expect(names).not.toContain('find_meeting_times');
    expect(names).not.toContain('list_rooms');
  });

  it('hides teams tools when teams is disabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail', 'calendar']);
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    expect(names).not.toContain('send_channel_message');
    expect(names).not.toContain('list_teams');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// get_schedule
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — get_schedule', () => {
  it('posts to /me/calendar/getSchedule with the requested people and returns busy blocks', async () => {
    mockGraphPost.mockResolvedValue({
      value: [{
        scheduleId: 'jeremiah@example.com',
        availabilityView: '000022220000',
        scheduleItems: [{ status: 'busy', start: { dateTime: '2026-07-06T15:00:00' }, end: { dateTime: '2026-07-06T16:00:00' }, subject: 'Standup' }],
        workingHours: { daysOfWeek: ['monday'], startTime: '09:00:00', endTime: '17:00:00' },
      }],
    });
    const res = await callTool('get_schedule', {
      schedules: ['jeremiah@example.com'], startDateTime: '2026-07-06T09:00:00', endDateTime: '2026-07-06T17:00:00',
    });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(lastGraphCall.path).toBe('/me/calendar/getSchedule');
    const post = lastGraphCall.postBody as { schedules: string[]; availabilityViewInterval: number };
    expect(post.schedules).toEqual(['jeremiah@example.com']);
    expect(post.availabilityViewInterval).toBe(30);
    expect((parsed as Array<{ scheduleId: string; busy: unknown[] }>)[0]).toMatchObject({ scheduleId: 'jeremiah@example.com' });
    expect((parsed as Array<{ busy: unknown[] }>)[0].busy).toHaveLength(1);
  });

  it('errors (no Graph call) when schedules is empty', async () => {
    const res = await callTool('get_schedule', { schedules: [], startDateTime: 'a', endDateTime: 'b' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// find_meeting_times
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — find_meeting_times', () => {
  it('posts to /me/findMeetingTimes with an ISO duration + timezone hint and returns suggestions', async () => {
    mockGraphPost.mockResolvedValue({
      meetingTimeSuggestions: [
        { confidence: 100, meetingTimeSlot: { start: {}, end: {} }, suggestionReason: 'works', attendeeAvailability: [{ attendee: { emailAddress: { address: 'jeremiah@example.com' } }, availability: 'free' }] },
      ],
    });
    const res = await callTool('find_meeting_times', { attendees: ['jeremiah@example.com'], meetingDurationMinutes: 30 });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(lastGraphCall.path).toBe('/me/findMeetingTimes');
    const post = lastGraphCall.postBody as { meetingDuration: string; attendees: unknown[] };
    expect(post.meetingDuration).toBe('PT30M');
    expect(post.attendees).toHaveLength(1);
    expect(lastGraphCall.headers['Prefer']).toContain('outlook.timezone');
    expect((parsed as { suggestions: unknown[] }).suggestions).toHaveLength(1);
  });

  it('errors when attendees is empty', async () => {
    const res = await callTool('find_meeting_times', { attendees: [] });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// list_rooms
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — list_rooms', () => {
  it('gets the room places collection and maps room fields', async () => {
    mockGraphGet.mockResolvedValue({
      value: [{ id: 'room-1', displayName: 'Montgomery A', emailAddress: 'monta@example.com', building: '44 Montgomery', capacity: 8 }],
    });
    const res = await callTool('list_rooms', {});
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(lastGraphCall.path).toBe('/places/microsoft.graph.room');
    expect((parsed as { items: Array<{ emailAddress: string }> }).items[0]).toMatchObject({ displayName: 'Montgomery A', emailAddress: 'monta@example.com', capacity: 8 });
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// create_event — online meeting enhancement
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — create_event isOnlineMeeting', () => {
  it('sets isOnlineMeeting + provider on the Graph body and surfaces the Teams join URL', async () => {
    mockGraphPost.mockResolvedValue({
      id: 'ev-1', subject: 'Intro', start: {}, end: {}, webLink: 'https://outlook/ev',
      isOnlineMeeting: true, onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/xyz' },
    });
    const res = await callTool('create_event', { subject: 'Intro', start: '2026-07-06T15:00:00', end: '2026-07-06T15:30:00', isOnlineMeeting: true });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    const post = lastGraphCall.postBody as { isOnlineMeeting: boolean; onlineMeetingProvider: string };
    expect(post.isOnlineMeeting).toBe(true);
    expect(post.onlineMeetingProvider).toBe('teamsForBusiness');
    expect(parsed).toMatchObject({ isOnlineMeeting: true, onlineMeetingUrl: 'https://teams.microsoft.com/l/meetup-join/xyz' });
  });

  it('does not set online-meeting fields when isOnlineMeeting is omitted', async () => {
    mockGraphPost.mockResolvedValue({ id: 'ev-2', subject: 'Plain', start: {}, end: {} });
    const res = await callTool('create_event', { subject: 'Plain', start: '2026-07-06T15:00:00', end: '2026-07-06T15:30:00' });
    const { parsed } = toolResult(res);
    const post = lastGraphCall.postBody as Record<string, unknown>;
    expect(post.isOnlineMeeting).toBeUndefined();
    expect(parsed).toMatchObject({ onlineMeetingUrl: null });
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Teams — list_teams / list_channels
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — list_teams / list_channels', () => {
  it('list_teams gets /me/joinedTeams and maps id/name', async () => {
    mockGraphGet.mockResolvedValue({ value: [{ id: 'team-1', displayName: 'Trading', description: 'desk' }] });
    const res = await callTool('list_teams', {});
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(lastGraphCall.path).toBe('/me/joinedTeams');
    expect((parsed as { items: Array<{ id: string; name: string }> }).items[0]).toMatchObject({ id: 'team-1', name: 'Trading' });
  });

  it('list_channels gets /teams/{teamId}/channels', async () => {
    mockGraphGet.mockResolvedValue({ value: [{ id: 'chan-1', displayName: 'General', membershipType: 'standard' }] });
    const res = await callTool('list_channels', { teamId: 'team-1' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(lastGraphCall.path).toBe('/teams/team-1/channels');
    expect((parsed as { items: Array<{ displayName: string }> }).items[0]).toMatchObject({ displayName: 'General' });
  });

  it('list_channels blocks a deny-listed team before the Graph call', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'team-1');
    const res = await callTool('list_channels', { teamId: 'team-1' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it('rejects a path-injection teamId before any Graph call (opaque-ID guard)', async () => {
    const res = await callTool('list_channels', { teamId: 'team/../../users/victim' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('Bad request');
    expect(mockGraphGet).not.toHaveBeenCalled();
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Teams — send_chat_message / send_channel_message
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — send_chat_message', () => {
  it('posts to /chats/{chatId}/messages and returns sent status', async () => {
    mockGraphPost.mockResolvedValue({ id: 'msg-1', webUrl: 'https://teams/msg' });
    const res = await callTool('send_chat_message', { chatId: 'chat-1', content: 'hello' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(lastGraphCall.path).toBe('/chats/chat-1/messages');
    const post = lastGraphCall.postBody as { body: { contentType: string; content: string } };
    expect(post.body).toMatchObject({ contentType: 'text', content: 'hello' });
    expect(parsed).toMatchObject({ status: 'sent', chatId: 'chat-1' });
  });

  it('blocks a deny-listed chat', async () => {
    mockIsPathDenied.mockResolvedValue(true);
    const res = await callTool('send_chat_message', { chatId: 'chat-1', content: 'hi' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});

describe('MCP tools/call — send_channel_message', () => {
  it('posts to /teams/{teamId}/channels/{channelId}/messages with html when requested', async () => {
    mockGraphPost.mockResolvedValue({ id: 'msg-2', webUrl: 'https://teams/chan-msg' });
    const res = await callTool('send_channel_message', { teamId: 'team-1', channelId: 'chan-1', content: '<b>green</b>', contentType: 'html' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(lastGraphCall.path).toBe('/teams/team-1/channels/chan-1/messages');
    const post = lastGraphCall.postBody as { body: { contentType: string } };
    expect(post.body.contentType).toBe('html');
    expect(parsed).toMatchObject({ status: 'sent', teamId: 'team-1', channelId: 'chan-1' });
  });

  it('blocks when the channel is deny-listed', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'chan-1');
    const res = await callTool('send_channel_message', { teamId: 'team-1', channelId: 'chan-1', content: 'x' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('is rejected when the teams service is not enabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail', 'calendar']);
    const res = await callTool('send_channel_message', { teamId: 'team-1', channelId: 'chan-1', content: 'x' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('not enabled');
  });

  it('rejects a path-injection channelId before any Graph call', async () => {
    const res = await callTool('send_channel_message', { teamId: 'team-1', channelId: 'a/../b', content: 'x' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});
