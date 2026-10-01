/**
 * Calendar access-control tests.
 *

 * Two controls, each enforced on BOTH surfaces (the HTTP routes AND the MCP
 * tools/call path) so neither can be used to bypass the other:
 *
 *   1. Read-only mode — when the calendar service is in the tenant's
 *      readOnlyServices set, create/update/delete are refused while
 *      list/get continue to work. Read-only is per-service, so a read-only
 *      mail setting must NOT block a calendar write. The write tools are
 *      also hidden from MCP tools/list, with the call-time refusal
 *      kept as the backstop.
 *   2. Deny-by-name — a deny-list entry recorded against a calendar's
 *      display name (not its opaque ID) blocks every calendar operation
 *      (list, get, create, update, delete), not just get_event.
 *
 * All Graph, auth, and storage dependencies are mocked.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Controllable mocks ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockGetValidAccessTokenForSession = jest.fn<() => Promise<string>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetReadOnlyServices = jest.fn<() => Promise<string[]>>();
const mockIsPathDenied = jest.fn<(t: string, u: string, type: string, path: string) => Promise<boolean>>();
const mockResolveDefaultCalendarId = jest.fn<() => Promise<string | null>>();
const mockResolveCalendarName = jest.fn<() => Promise<string>>();

const graphCalls: { get: number; post: number; patch: number; delete: number; lastPath: string | null } = {
  get: 0, post: 0, patch: 0, delete: 0, lastPath: null,
};
const graphResult: { value: unknown } = { value: { id: 'ev', subject: 's', start: {}, end: {}, webLink: 'http://x', value: [] } };
function makeChain() {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header']) chain[m] = () => chain;
  chain.get = () => { graphCalls.get++; return Promise.resolve(graphResult.value); };
  chain.post = () => { graphCalls.post++; return Promise.resolve(graphResult.value); };
  chain.patch = () => { graphCalls.patch++; return Promise.resolve(graphResult.value); };
  chain.delete = () => { graphCalls.delete++; return Promise.resolve(undefined); };
  return chain;
}
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => { graphCalls.lastPath = path; return makeChain(); },
}));

jest.mock('../services/telemetry.js', () => ({}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
}));

jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => mockGetValidAccessTokenForSession(),
  getTenantId: () => Promise.resolve('tenant-a'),
  getTenantIdFromSession: (session: { tenantId?: string }) => {
    if (!session?.tenantId) throw new Error('No tenantId in session');
    return session.tenantId;
  },
}));

jest.mock('../services/mailboxTimeZone.js', () => ({
  resolveMailboxTimeZone: jest.fn(async () => 'Pacific Standard Time'),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => mockCreateGraphClient(),
}));

jest.mock('../services/denyList.js', () => ({
  // Mirror the real filterDeniedPaths: exclude an item when its id OR name is denied.
  filterDeniedPaths: async (_t: unknown, _u: unknown, _s: unknown, items: Array<{ id?: string; name?: string }>) => {
    const out: Array<{ id?: string; name?: string }> = [];
    for (const it of items) {
      const idDenied = it.id ? await mockIsPathDenied('t', 'u', 'calendar', it.id) : false;
      const nameDenied = it.name ? await mockIsPathDenied('t', 'u', 'calendar', it.name) : false;
      if (!idDenied && !nameDenied) out.push(it);
    }
    return out;
  },
  isPathDenied: (...a: unknown[]) => mockIsPathDenied(a[0] as string, a[1] as string, a[2] as string, a[3] as string),
}));

jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => Promise.resolve(null),
  resolveDefaultCalendarId: () => mockResolveDefaultCalendarId(),
  resolveCalendarName: () => mockResolveCalendarName(),
  resolveContactParentFolder: () => Promise.resolve(null),
  resolveDefaultContactFolder: () => Promise.resolve(null),
  resolveSectionNotebook: () => Promise.resolve(null),
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
  clearResolverCaches: jest.fn(),
}));

jest.mock('../services/sharepointFilter.js', () => ({
  filterAndDisambiguateSites: (sites: unknown) => sites,
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => mockGetEnabledServices(),
  getReadOnlyServices: () => mockGetReadOnlyServices(),
  getAllowedSites: () => Promise.resolve([]),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: () => Promise.resolve([]),
  isServiceDisabledForUser: () => Promise.resolve(false),
}));

jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'draft' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: () => Promise.resolve(false),
}));

jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

// ── Import after mocks ────────────────────────────────────────────────────────

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';
import '../functions/calendar/listEvents.js';
import '../functions/calendar/getEvent.js';
import '../functions/calendar/createEvent.js';
import '../functions/calendar/updateEvent.js';
import '../functions/calendar/deleteEvent.js';

interface HttpRegistration {
  handler: (req: HttpRequest, ctx: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
function getHandler(name: string) {
  const reg = httpMock.mock.calls.find((c) => c[0] === name);
  if (!reg) throw new Error(`Handler not registered: ${name}`);
  return reg[1].handler;
}

// ── Fixtures ───────────────────────────────────────────────────────────────────

const TENANT = 'tenant-a';
const USER = 'user-example';
const DEFAULT_CAL_ID = 'default-cal-guid';
const EVENT_ID = 'AAMkEventGuid';
const AUTH = { userId: USER, session: { userId: USER, tenantId: TENANT, email: 'u@example.com', accessToken: 'x', sessionToken: 's' } } as unknown as AuthResult;
const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;

// Capture handler references at module load — jest.clearAllMocks() in
// beforeEach wipes httpMock.mock.calls, so getHandler must run before then.
const mcpHandler = getHandler('mcpEndpoint');
const HTTP = {
  listEvents: getHandler('listEvents'),
  getEvent: getHandler('getEvent'),
  createEvent: getHandler('createEvent'),
  updateEvent: getHandler('updateEvent'),
  deleteEvent: getHandler('deleteEvent'),
};
function callTool(name: string, args: Record<string, unknown>) {
  const req = {
    method: 'POST', headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
  } as unknown as HttpRequest;
  return mcpHandler(req, ctx);
}
async function listToolNames(): Promise<string[]> {
  const req = {
    method: 'POST', headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} }),
  } as unknown as HttpRequest;
  const res = await mcpHandler(req, ctx);
  const body = res.jsonBody as { result?: { tools?: Array<{ name: string }> } };
  return (body.result?.tools ?? []).map((t) => t.name);
}
function toolIsError(res: { jsonBody?: unknown }): { isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  return { isError: body.result?.isError ?? false, text: body.result?.content?.[0]?.text ?? '' };
}
function makeReq(opts: { params?: Record<string, string>; query?: Record<string, string>; body?: unknown }): HttpRequest {
  return {
    json: async () => opts.body ?? {},
    params: opts.params ?? {},
    query: { get: (k: string) => (opts.query && k in opts.query ? opts.query[k] : null) },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

const CREATE_BODY = { subject: 'Sync', start: '2026-07-01T10:00:00', end: '2026-07-01T11:00:00' };

beforeEach(() => {
  jest.clearAllMocks();
  graphCalls.get = graphCalls.post = graphCalls.patch = graphCalls.delete = 0;
  graphCalls.lastPath = null;
  graphResult.value = { id: EVENT_ID, subject: 's', start: {}, end: {}, webLink: 'http://x', value: [] };
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetEnabledServices.mockResolvedValue(['calendar', 'mail', 'sharepoint']);
  mockGetReadOnlyServices.mockResolvedValue([]);
  mockIsPathDenied.mockResolvedValue(false);
  mockResolveDefaultCalendarId.mockResolvedValue(DEFAULT_CAL_ID);
  mockResolveCalendarName.mockResolvedValue('My Calendar');
});

// ─────────────────────────────────────────────────────────────────────────────
// Read-only mode — MCP tools/call
// ─────────────────────────────────────────────────────────────────────────────

describe('read-only mode — MCP tools/call', () => {
  beforeEach(() => mockGetReadOnlyServices.mockResolvedValue(['calendar']));

  it.each(['create_event', 'update_event', 'delete_event'])(
    'refuses %s and never reaches Graph when calendar is read-only', async (tool) => {
      const args = tool === 'create_event' ? CREATE_BODY : { eventId: EVENT_ID, subject: 'x' };
      const { isError, text } = toolIsError(await callTool(tool, args));
      expect(isError).toBe(true);
      expect(text).toContain('read-only mode');
      expect(graphCalls.post + graphCalls.patch + graphCalls.delete).toBe(0);
    },
  );

  it('still allows list_events and get_event in read-only mode', async () => {
    expect(toolIsError(await callTool('list_events', {})).isError).toBe(false);
    expect(toolIsError(await callTool('get_event', { eventId: EVENT_ID })).isError).toBe(false);
    expect(graphCalls.get).toBeGreaterThan(0);
  });

  it('does not block a calendar write when only mail is read-only (per-service scope)', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['mail']);
    const { isError } = toolIsError(await callTool('create_event', CREATE_BODY));
    expect(isError).toBe(false);
    expect(graphCalls.post).toBe(1);
  });

  it('allows all writes when read-only set is empty', async () => {
    mockGetReadOnlyServices.mockResolvedValue([]);
    expect(toolIsError(await callTool('create_event', CREATE_BODY)).isError).toBe(false);
    expect(graphCalls.post).toBe(1);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Read-only mode — MCP tools/list
// ─────────────────────────────────────────────────────────────────────────────

const CALENDAR_WRITE_TOOLS = ['create_calendar', 'create_event', 'update_event', 'delete_event', 'move_event', 'respond_to_event'];
const CALENDAR_READ_TOOLS = ['list_calendars', 'list_events', 'get_event', 'get_schedule', 'find_meeting_times', 'list_rooms'];

describe('read-only mode — MCP tools/list', () => {
  it('hides calendar write tools when calendar is read-only', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['calendar']);
    const names = await listToolNames();
    for (const t of CALENDAR_WRITE_TOOLS) expect(names).not.toContain(t);
  });

  it('keeps calendar read tools listed when calendar is read-only', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['calendar']);
    const names = await listToolNames();
    for (const t of CALENDAR_READ_TOOLS) expect(names).toContain(t);
  });

  it('does not hide another service\'s write tools (per-service scope)', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['calendar']);
    const names = await listToolNames();
    expect(names).toContain('send_mail');
    expect(names).toContain('write_sharepoint_file');
  });

  it('lists all calendar tools when the read-only set is empty', async () => {
    mockGetReadOnlyServices.mockResolvedValue([]);
    const names = await listToolNames();
    for (const t of [...CALENDAR_WRITE_TOOLS, ...CALENDAR_READ_TOOLS]) expect(names).toContain(t);
  });

  it('a hidden write tool is still refused at tools/call (backstop)', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['calendar']);
    expect(await listToolNames()).not.toContain('create_event');
    const { isError, text } = toolIsError(await callTool('create_event', CREATE_BODY));
    expect(isError).toBe(true);
    expect(text).toContain('read-only mode');
    expect(graphCalls.post).toBe(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Read-only mode — HTTP routes
// ─────────────────────────────────────────────────────────────────────────────

describe('read-only mode — HTTP routes', () => {
  beforeEach(() => mockGetReadOnlyServices.mockResolvedValue(['calendar']));

  it('createEvent returns 403 and never reaches Graph', async () => {
    const res = await HTTP.createEvent(makeReq({ body: CREATE_BODY }), ctx);
    expect(res.status).toBe(403);
    expect(graphCalls.post).toBe(0);
  });

  it('updateEvent returns 403 and never reaches Graph', async () => {
    const res = await HTTP.updateEvent(makeReq({ params: { eventId: EVENT_ID }, body: { subject: 'x' } }), ctx);
    expect(res.status).toBe(403);
    expect(graphCalls.patch).toBe(0);
  });

  it('deleteEvent returns 403 and never reaches Graph', async () => {
    const res = await HTTP.deleteEvent(makeReq({ params: { eventId: EVENT_ID } }), ctx);
    expect(res.status).toBe(403);
    expect(graphCalls.delete).toBe(0);
  });

  it('listEvents and getEvent still return 200 (reads unaffected)', async () => {
    expect((await HTTP.listEvents(makeReq({}), ctx)).status).toBe(200);
    expect((await HTTP.getEvent(makeReq({ params: { eventId: EVENT_ID } }), ctx)).status).toBe(200);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Deny-by-name — MCP tools/call (all calendar ops, not just get_event)
// ─────────────────────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────────────────────
// list_calendars — a deny-listed calendar is never enumerated (MCP)
// ─────────────────────────────────────────────────────────────────────────────

describe('list_calendars — deny filtering (MCP)', () => {
  it('excludes a calendar whose NAME is denied and keeps the rest', async () => {
    graphResult.value = { value: [{ id: 'c1', name: 'HR Sensitive' }, { id: 'c2', name: 'My Calendar' }] };
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'HR Sensitive');
    const { isError, text } = toolIsError(await callTool('list_calendars', {}));
    expect(isError).toBe(false);
    const parsed = JSON.parse(text) as { items: Array<{ id: string }> };
    expect(parsed.items.map((c) => c.id)).toEqual(['c2']);
  });

  it('excludes a calendar whose ID is denied', async () => {
    graphResult.value = { value: [{ id: 'c1', name: 'A' }, { id: 'c2', name: 'B' }] };
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'c1');
    const { text } = toolIsError(await callTool('list_calendars', {}));
    const parsed = JSON.parse(text) as { items: Array<{ id: string }> };
    expect(parsed.items.map((c) => c.id)).toEqual(['c2']);
  });
});

describe('deny-by-name — MCP tools/call', () => {
  beforeEach(() => {
    mockResolveCalendarName.mockResolvedValue('HR Sensitive');
    // ID not denied; only the resolved name is on the deny list.
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'HR Sensitive');
  });

  it.each(['list_events', 'get_event', 'create_event', 'update_event', 'delete_event'])(
    'blocks %s when the calendar NAME is denied', async (tool) => {
      const args =
        tool === 'create_event' ? CREATE_BODY :
        tool === 'list_events' ? {} :
        { eventId: EVENT_ID, subject: 'x' };
      const { isError } = toolIsError(await callTool(tool, args));
      expect(isError).toBe(true);
      expect(graphCalls.get + graphCalls.post + graphCalls.patch + graphCalls.delete).toBe(0);
    },
  );
});

// ─────────────────────────────────────────────────────────────────────────────
// Deny-by-name — HTTP routes
// ─────────────────────────────────────────────────────────────────────────────

describe('deny-by-name — HTTP routes', () => {
  beforeEach(() => {
    mockResolveCalendarName.mockResolvedValue('HR Sensitive');
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'HR Sensitive');
  });

  it('listEvents returns 403 on a name-denied calendar', async () => {
    const res = await HTTP.listEvents(makeReq({}), ctx);
    expect(res.status).toBe(403);
    expect(graphCalls.get).toBe(0);
  });

  it('createEvent returns 403 on a name-denied calendar', async () => {
    const res = await HTTP.createEvent(makeReq({ body: CREATE_BODY }), ctx);
    expect(res.status).toBe(403);
    expect(graphCalls.post).toBe(0);
  });

  it('updateEvent returns 403 on a name-denied calendar', async () => {
    const res = await HTTP.updateEvent(makeReq({ params: { eventId: EVENT_ID }, body: { subject: 'x' } }), ctx);
    expect(res.status).toBe(403);
    expect(graphCalls.patch).toBe(0);
  });

  it('deleteEvent returns 403 on a name-denied calendar', async () => {
    const res = await HTTP.deleteEvent(makeReq({ params: { eventId: EVENT_ID } }), ctx);
    expect(res.status).toBe(403);
    expect(graphCalls.delete).toBe(0);
  });
});
