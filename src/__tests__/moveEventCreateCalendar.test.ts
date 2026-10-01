/**
 * move_event + create_calendar MCP tool tests.
 *
 * Two calendar gaps, both MCP-native (endpoint /api/mcp):
 *   - create_calendar: POST /me/calendars, deny-by-name + read-only enforced.
 *   - move_event: Graph has no native event-move, so this is copy-then-delete.
 *     Tests cover field preservation, the attendee/organizer/recurring-occurrence
 *     safety refusals, the create-before-delete ordering, attachment copy, and
 *     deny-list / read-only / opaque-ID enforcement on both calendars.
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
const mockGetReadOnlyServices = jest.fn<() => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<() => Promise<Array<{ id: string; name: string }>>>();
const mockGetUserServiceOverrides = jest.fn<() => Promise<string[]>>();
const mockResolveDefaultCalendarId = jest.fn<() => Promise<string | null>>();
const mockResolveCalendarName = jest.fn<() => Promise<string>>();

interface GraphCall {
  path: string;
  method: 'get' | 'post' | 'delete';
  body?: unknown;
}
const graphCalls: GraphCall[] = [];
const mockGraphGet = jest.fn<(path: string) => Promise<unknown>>();
const mockGraphPost = jest.fn<(path: string, body: unknown) => Promise<unknown>>();

function makeChain(path: string) {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header']) chain[m] = () => chain;
  chain.get = async () => {
    graphCalls.push({ path, method: 'get' });
    return mockGraphGet(path);
  };
  chain.post = async (body: unknown) => {
    graphCalls.push({ path, method: 'post', body });
    return mockGraphPost(path, body);
  };
  chain.patch = () => Promise.resolve({});
  chain.delete = async () => {
    graphCalls.push({ path, method: 'delete' });
    return undefined;
  };
  return chain;
}
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => makeChain(path),
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
  getReadOnlyServices: () => mockGetReadOnlyServices(),
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
const TARGET_CAL_ID = 'target-cal-id';
const SOURCE_CAL_ID = 'source-cal-id';
const EVENT_ID = 'AAMkEventGuid007';
const NEW_EVENT_ID = 'AAMkNewEventGuid042';
const NEW_CAL_ID = 'AAMkNewCalGuid099';
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

function toolResult(res: { jsonBody?: unknown }): { parsed: any; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

function graphNotFound(code = 'ErrorItemNotFound', statusCode = 404): Error {
  const err = new Error(code) as Error & { statusCode: number; code: string };
  err.statusCode = statusCode;
  err.code = code;
  return err;
}

// A single-instance event the signed-in user organizes, with no attendees.
function baseEvent(overrides: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    id: EVENT_ID,
    subject: 'Quarterly budget review',
    body: { contentType: 'HTML', content: '<p>Numbers</p>' },
    start: { dateTime: '2026-08-01T09:00:00', timeZone: 'America/Los_Angeles' },
    end: { dateTime: '2026-08-01T10:00:00', timeZone: 'America/Los_Angeles' },
    location: { displayName: 'Room 5', locationType: 'default' },
    categories: ['Finance'],
    sensitivity: 'private',
    showAs: 'busy',
    importance: 'high',
    isAllDay: false,
    isReminderOn: true,
    reminderMinutesBeforeStart: 30,
    recurrence: null,
    attendees: [],
    isOrganizer: true,
    type: 'singleInstance',
    seriesMasterId: null,
    hasAttachments: false,
    ...overrides,
  };
}

const posts = () => graphCalls.filter((c) => c.method === 'post');
const deletes = () => graphCalls.filter((c) => c.method === 'delete');

beforeEach(() => {
  jest.clearAllMocks();
  graphCalls.length = 0;
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['calendar', 'mail', 'sharepoint']);
  mockGetReadOnlyServices.mockResolvedValue([]);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
  mockResolveDefaultCalendarId.mockResolvedValue(DEFAULT_CAL_ID);
  mockResolveCalendarName.mockResolvedValue('My Calendar');
  // Default source event + target create.
  mockGraphGet.mockImplementation(async (path) => {
    if (path.endsWith('/attachments')) return { value: [] };
    if (path.includes('/events/')) return baseEvent();
    return {};
  });
  mockGraphPost.mockImplementation(async (path) => {
    if (path === '/me/calendars') return { id: NEW_CAL_ID, name: 'Private', color: 'auto', isDefaultCalendar: false };
    if (path.endsWith('/events')) return { id: NEW_EVENT_ID, webLink: 'https://outlook/new' };
    return {};
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// toolManifest — registration + schema sync
// ─────────────────────────────────────────────────────────────────────────────

describe('toolManifest — create_calendar + move_event', () => {
  it('registers both as MCP-native POST tools with the right required params', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const byName = Object.fromEntries(TOOLS.map((t) => [t.name, t]));

    expect(byName['create_calendar']).toBeDefined();
    expect(byName['create_calendar'].endpoint).toBe('/api/mcp');
    expect(byName['create_calendar'].method).toBe('POST');
    expect(byName['create_calendar'].parameters.filter((p) => p.required).map((p) => p.name)).toEqual(['name']);
    expect(byName['create_calendar'].parameters.find((p) => p.name === 'color')?.enum).toContain('lightBlue');

    expect(byName['move_event']).toBeDefined();
    expect(byName['move_event'].endpoint).toBe('/api/mcp');
    expect(byName['move_event'].method).toBe('POST');
    expect(byName['move_event'].parameters.filter((p) => p.required).map((p) => p.name).sort()).toEqual(
      ['eventId', 'targetCalendarId'],
    );
  });

  it('getManifest() and live tools/list expose the same input-schema shape', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string; inputSchema: Record<string, unknown> }> };
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string; inputSchema: Record<string, unknown> }> } };

    interface SchemaShape { type: unknown; required: unknown[]; props: Record<string, { type: unknown; enum?: unknown }> }
    function shape(inputSchema: any): SchemaShape {
      const props: SchemaShape['props'] = {};
      for (const [name, spec] of Object.entries(inputSchema.properties ?? {})) {
        const s = spec as any;
        props[name] = { type: s.type, ...(s.enum ? { enum: s.enum } : {}) };
      }
      return { type: inputSchema.type, required: [...(inputSchema.required ?? [])].sort(), props };
    }

    for (const name of ['create_calendar', 'move_event']) {
      const fromManifest = manifest.tools.find((t) => t.name === name);
      const fromList = body.result.tools.find((t) => t.name === name);
      expect(fromManifest).toBeDefined();
      expect(fromList).toBeDefined();
      expect(shape(fromManifest!.inputSchema)).toEqual(shape(fromList!.inputSchema));
    }
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/list — exposure follows the calendar service toggle
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/list — exposure', () => {
  it('lists both tools when the calendar service is enabled', async () => {
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    expect(names).toContain('create_calendar');
    expect(names).toContain('move_event');
  });

  it('hides both tools when the calendar service is disabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail']);
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    expect(names).not.toContain('create_calendar');
    expect(names).not.toContain('move_event');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// create_calendar
// ─────────────────────────────────────────────────────────────────────────────

describe('create_calendar', () => {
  it('POSTs to /me/calendars with the name', async () => {
    const res = await callTool('create_calendar', { name: 'Private' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'created', id: NEW_CAL_ID, name: 'Private' });
    expect(posts()).toHaveLength(1);
    expect(posts()[0].path).toBe('/me/calendars');
    expect(posts()[0].body).toEqual({ name: 'Private' });
  });

  it('includes color when provided', async () => {
    await callTool('create_calendar', { name: 'Private', color: 'lightBlue' });
    expect(posts()[0].body).toEqual({ name: 'Private', color: 'lightBlue' });
  });

  it('refuses an empty name before any Graph call', async () => {
    const res = await callTool('create_calendar', { name: '   ' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(graphCalls).toHaveLength(0);
  });

  it('refuses when the calendar name is on the deny list (no Graph call)', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'Executive');
    const res = await callTool('create_calendar', { name: 'Executive' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('deny list');
    expect(graphCalls).toHaveLength(0);
  });

  it('refuses a whitespace-padded name that matches a deny entry (no Graph call)', async () => {
    // Regression: isPathDenied canonicalizes slash/case but not surrounding
    // whitespace, so " Executive " must be trimmed before the deny check or it
    // slips past a deny entry for "Executive".
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'Executive');
    const res = await callTool('create_calendar', { name: '  Executive  ' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('deny list');
    expect(graphCalls).toHaveLength(0);
  });

  it('trims surrounding whitespace before the deny check and the Graph create body', async () => {
    const res = await callTool('create_calendar', { name: '  Private  ' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'created' });
    // Deny check saw the trimmed name, and the create body is trimmed too.
    expect(mockIsPathDenied).toHaveBeenCalledWith(expect.anything(), expect.anything(), 'calendar', 'Private');
    expect(posts()).toHaveLength(1);
    expect(posts()[0].body).toEqual({ name: 'Private' });
  });

  it('is refused when the calendar service is in read-only mode', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['calendar']);
    const res = await callTool('create_calendar', { name: 'Private' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('read-only mode');
    expect(graphCalls).toHaveLength(0);
  });

  it('is rejected when the calendar service is not enabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail']);
    const res = await callTool('create_calendar', { name: 'Private' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('not enabled');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// move_event — happy path + field preservation
// ─────────────────────────────────────────────────────────────────────────────

describe('move_event — copy-then-delete', () => {
  it('copies the event to the target calendar, then deletes the source, and returns the new id', async () => {
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({
      status: 'moved',
      eventId: NEW_EVENT_ID,
      sourceEventId: EVENT_ID,
      targetCalendarId: TARGET_CAL_ID,
      attendeesNotified: false,
    });

    // create in target must precede delete of source.
    const createIdx = graphCalls.findIndex((c) => c.method === 'post' && c.path === `/me/calendars/${TARGET_CAL_ID}/events`);
    const deleteIdx = graphCalls.findIndex((c) => c.method === 'delete' && c.path === `/me/events/${EVENT_ID}`);
    expect(createIdx).toBeGreaterThanOrEqual(0);
    expect(deleteIdx).toBeGreaterThan(createIdx);
  });

  it('preserves body, location, categories, sensitivity, showAs, importance, reminders, and recurrence', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path.endsWith('/attachments')) return { value: [] };
      if (path.includes('/events/')) return baseEvent({ recurrence: { pattern: { type: 'weekly' } } });
      return {};
    });
    await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID });
    const createBody = posts().find((p) => p.path === `/me/calendars/${TARGET_CAL_ID}/events`)?.body as Record<string, unknown>;
    expect(createBody).toMatchObject({
      subject: 'Quarterly budget review',
      body: { contentType: 'HTML', content: '<p>Numbers</p>' },
      location: { displayName: 'Room 5', locationType: 'default' },
      categories: ['Finance'],
      sensitivity: 'private',
      showAs: 'busy',
      importance: 'high',
      isReminderOn: true,
      reminderMinutesBeforeStart: 30,
      recurrence: { pattern: { type: 'weekly' } },
    });
  });

  it('reads the source from an explicit source calendar when calendarId is given', async () => {
    await callTool('move_event', { eventId: EVENT_ID, calendarId: SOURCE_CAL_ID, targetCalendarId: TARGET_CAL_ID });
    expect(graphCalls.some((c) => c.method === 'get' && c.path === `/me/calendars/${SOURCE_CAL_ID}/events/${EVENT_ID}`)).toBe(true);
    expect(deletes()[0].path).toBe(`/me/calendars/${SOURCE_CAL_ID}/events/${EVENT_ID}`);
  });

  it('refuses when source and target calendars are identical', async () => {
    const res = await callTool('move_event', { eventId: EVENT_ID, calendarId: TARGET_CAL_ID, targetCalendarId: TARGET_CAL_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('same');
    expect(graphCalls).toHaveLength(0);
  });

  it('refuses when calendarId is omitted and the target IS the default calendar (no create, no delete)', async () => {
    // Omitted calendarId => the source is the default calendar. Targeting that
    // same default calendar is a no-op move; the guard must resolve the default
    // calendar ID and refuse rather than silently copy-then-delete in place
    // (which would churn the event id/webLink and, with attendees + force, fire
    // fresh invites plus a cancellation for nothing).
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: DEFAULT_CAL_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('same');
    expect(posts()).toHaveLength(0);
    expect(deletes()).toHaveLength(0);
  });

  it('errors clearly when the source event does not exist (no delete)', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path.includes('/events/')) throw graphNotFound();
      return {};
    });
    const res = await callTool('move_event', { eventId: 'AAMkNope', targetCalendarId: TARGET_CAL_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('No calendar event found');
    expect(deletes()).toHaveLength(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// move_event — safety refusals
// ─────────────────────────────────────────────────────────────────────────────

describe('move_event — safety', () => {
  it('refuses an event with attendees unless force=true (no create, no delete)', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path.endsWith('/attachments')) return { value: [] };
      if (path.includes('/events/')) return baseEvent({ attendees: [{ emailAddress: { address: 'a@x.com' }, type: 'required' }] });
      return {};
    });
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('force=true');
    expect(posts()).toHaveLength(0);
    expect(deletes()).toHaveLength(0);
  });

  it('proceeds with attendees when force=true, copies them, and reports attendeesNotified', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path.endsWith('/attachments')) return { value: [] };
      if (path.includes('/events/')) return baseEvent({ attendees: [{ emailAddress: { address: 'a@x.com' }, type: 'optional' }] });
      return {};
    });
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID, force: true });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'moved', attendeesNotified: true });
    const createBody = posts().find((p) => p.path === `/me/calendars/${TARGET_CAL_ID}/events`)?.body as Record<string, unknown>;
    expect(createBody.attendees).toEqual([{ emailAddress: { address: 'a@x.com' }, type: 'optional' }]);
  });

  it('refuses when the caller is not the organizer, even with force', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path.endsWith('/attachments')) return { value: [] };
      if (path.includes('/events/')) return baseEvent({ isOrganizer: false, attendees: [{ emailAddress: { address: 'org@x.com' } }] });
      return {};
    });
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID, force: true });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('not the organizer');
    expect(posts()).toHaveLength(0);
    expect(deletes()).toHaveLength(0);
  });

  it('refuses a single occurrence of a recurring series (points at the series master)', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path.includes('/events/')) return baseEvent({ type: 'occurrence', seriesMasterId: 'MASTER123' });
      return {};
    });
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('series master');
    expect(posts()).toHaveLength(0);
    expect(deletes()).toHaveLength(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// move_event — attachments
// ─────────────────────────────────────────────────────────────────────────────

describe('move_event — attachments', () => {
  it('copies file attachments to the new event and warns about non-file attachments', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path.endsWith('/attachments')) {
        return {
          value: [
            { '@odata.type': '#microsoft.graph.fileAttachment', name: 'budget.xlsx', contentType: 'application/vnd.ms-excel', contentBytes: 'QkJC', isInline: false },
            { '@odata.type': '#microsoft.graph.itemAttachment', name: 'linked-mail', id: 'att2' },
          ],
        };
      }
      if (path.includes('/events/')) return baseEvent({ hasAttachments: true });
      return {};
    });
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    const attPost = posts().find((p) => p.path === `/me/calendars/${TARGET_CAL_ID}/events/${NEW_EVENT_ID}/attachments`);
    expect(attPost).toBeDefined();
    expect(attPost!.body).toMatchObject({ '@odata.type': '#microsoft.graph.fileAttachment', name: 'budget.xlsx', contentBytes: 'QkJC' });
    expect(parsed.warnings.join(' ')).toContain('linked-mail');
    // Source still deleted after a successful copy.
    expect(deletes()).toHaveLength(1);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// move_event — deny-list / read-only / opaque-ID enforcement
// ─────────────────────────────────────────────────────────────────────────────

describe('move_event — security', () => {
  it('refuses when the target calendar is on the deny list (no create, no delete)', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === TARGET_CAL_ID);
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('deny list');
    expect(posts()).toHaveLength(0);
    expect(deletes()).toHaveLength(0);
  });

  it('refuses when the source (default) calendar is on the deny list', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === DEFAULT_CAL_ID);
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(posts()).toHaveLength(0);
    expect(deletes()).toHaveLength(0);
  });

  it('is refused when the calendar service is in read-only mode', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['calendar']);
    const res = await callTool('move_event', { eventId: EVENT_ID, targetCalendarId: TARGET_CAL_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('read-only mode');
    expect(graphCalls).toHaveLength(0);
  });

  it('rejects a path-injection targetCalendarId before any Graph call (opaque-ID guard)', async () => {
    const res = await callTool('move_event', {
      eventId: EVENT_ID, targetCalendarId: 'legit/../../../users/victim/events/x',
    });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('Bad request');
    expect(graphCalls).toHaveLength(0);
  });
});
