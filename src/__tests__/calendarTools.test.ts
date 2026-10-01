/**
 * Calendar MCP exposure tests.
 *
 * Two layers:
 *   1. toolManifest — all 6 calendar tools are registered in the TOOLS catalog
 *      and surface through getManifest() with correct params and HTTP methods.
 *   2. Remote MCP endpoint (/api/mcp JSON-RPC) — the calendar tools, including
 *      the new get_event, appear in tools/list and are invokable via tools/call
 *      with deny-list (by ID and by name) and opaque-ID enforcement intact.
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
const mockResolveCalendarName = jest.fn<() => Promise<string>>();

const lastGraphCall: { path: string | null } = { path: null };
const mockGraphGet = jest.fn<() => Promise<unknown>>();
function makeChain() {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header']) chain[m] = () => chain;
  chain.get = (...a: unknown[]) => mockGraphGet(...(a as []));
  chain.post = () => Promise.resolve({});
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
  resolveCalendarName: () => mockResolveCalendarName(),
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
const DEFAULT_CAL_ID = 'default-cal-id';
const EVENT_ID = 'AAMkEventGuid007';
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
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['calendar', 'mail', 'sharepoint']);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
  mockResolveDefaultCalendarId.mockResolvedValue(DEFAULT_CAL_ID);
  mockResolveCalendarName.mockResolvedValue('My Calendar');
  mockGraphGet.mockResolvedValue({
    id: EVENT_ID, subject: 'Quarterly Review', start: { dateTime: '2026-07-01T10:00:00' },
    end: { dateTime: '2026-07-01T11:00:00' }, location: { displayName: 'Room 4' },
    organizer: { emailAddress: { address: 'boss@example.com' } },
    attendees: [{ emailAddress: { address: 'a@example.com' } }], isAllDay: false,
    body: { contentType: 'text', content: 'Agenda' }, webLink: 'https://outlook/ev',
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// toolManifest — all 6 calendar tools registered
// ─────────────────────────────────────────────────────────────────────────────

describe('toolManifest — calendar tools', () => {
  const CAL_TOOLS = ['list_calendars', 'list_events', 'get_event', 'create_event', 'update_event', 'delete_event'];

  it('TOOLS contains all 6 calendar tools under /api/calendar/', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const names = TOOLS.filter((t) => t.endpoint.startsWith('/api/calendar/')).map((t) => t.name);
    for (const t of CAL_TOOLS) expect(names).toContain(t);
  });

  it('get_event requires eventId and is a GET', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const getEvent = TOOLS.find((t) => t.name === 'get_event');
    expect(getEvent?.method).toBe('GET');
    expect(getEvent?.endpoint).toBe('/api/calendar/events/{eventId}');
    expect(getEvent?.parameters.filter((p) => p.required).map((p) => p.name)).toEqual(['eventId']);
  });

  it('create_event requires subject, start, end; write tools use correct verbs', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const byName = Object.fromEntries(TOOLS.map((t) => [t.name, t]));
    expect(byName['create_event'].parameters.filter((p) => p.required).map((p) => p.name)).toEqual(
      expect.arrayContaining(['subject', 'start', 'end']),
    );
    expect(byName['create_event'].method).toBe('POST');
    expect(byName['update_event'].method).toBe('PATCH');
    expect(byName['delete_event'].method).toBe('DELETE');
    for (const t of ['update_event', 'delete_event']) {
      expect(byName[t].parameters.filter((p) => p.required).map((p) => p.name)).toContain('eventId');
    }
  });

  it('getManifest() output includes the calendar tools', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string }> };
    const names = manifest.tools.map((t) => t.name);
    for (const t of CAL_TOOLS) expect(names).toContain(t);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/list — calendar tools visible
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/list — calendar exposure', () => {
  it('lists all 6 calendar tools when the calendar service is enabled', async () => {
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string; inputSchema: unknown }> } };
    const names = body.result.tools.map((t) => t.name);
    for (const t of ['list_calendars', 'list_events', 'get_event', 'create_event', 'update_event', 'delete_event']) {
      expect(names).toContain(t);
    }
  });

  it('hides calendar tools when the calendar service is disabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail']);
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    expect(names).not.toContain('get_event');
    expect(names).not.toContain('list_events');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — get_event invokability + security
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — get_event', () => {
  it('returns the event payload and uses the calendar-scoped path when calendarId is given', async () => {
    const res = await callTool('get_event', { eventId: EVENT_ID, calendarId: 'AAMkExplicitCal' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ id: EVENT_ID, subject: 'Quarterly Review' });
    expect(lastGraphCall.path).toBe(`/me/calendars/AAMkExplicitCal/events/${EVENT_ID}`);
  });

  it('uses /me/events/{eventId} when calendarId is omitted', async () => {
    await callTool('get_event', { eventId: EVENT_ID });
    expect(lastGraphCall.path).toBe(`/me/events/${EVENT_ID}`);
  });

  it('blocks the read when the calendar ID is on the deny list (no Graph call)', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === DEFAULT_CAL_ID);
    const res = await callTool('get_event', { eventId: EVENT_ID });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it('blocks the read when the calendar NAME is on the deny list', async () => {
    mockResolveCalendarName.mockResolvedValue('HR Sensitive Calendar');
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'HR Sensitive Calendar');
    const res = await callTool('get_event', { eventId: EVENT_ID });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it('rejects a path-injection eventId before any Graph call (opaque-ID guard)', async () => {
    const res = await callTool('get_event', { eventId: 'legit/../../../users/victim/events' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('Bad request');
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it('rejects a path-injection calendarId before any Graph call', async () => {
    const res = await callTool('get_event', { eventId: EVENT_ID, calendarId: 'a/../../b' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(mockGraphGet).not.toHaveBeenCalled();
  });

  it('is rejected when the calendar service is not enabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail']);
    const res = await callTool('get_event', { eventId: EVENT_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('not enabled');
  });
});
