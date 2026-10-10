/**
 * Calendar create → get → update → delete round-trip through the native MCP
 * endpoint.
 *
 * This is the test the ticket asks for: "create → get → update → delete,
 * asserting the echoed timeZone matches the mailbox. Every defect here would
 * have been caught by that one test." It covers the two severity-1 defects:
 *
 *   1. Time zone — create_event/update_event defaulted a missing timeZone to a
 *      hardcoded America/New_York, so a Pacific user's events landed 3h early.
 *      The default is now the mailbox's own mailboxSettings.timeZone.
 *   2. ID round-trip — the native MCP handlers interpolated the Graph event ID
 *      into the URL path verbatim. Graph IDs contain `+`/`=`, which the SDK does
 *      not encode, so a raw `+` became a space and Graph answered "The Id is
 *      invalid" for get/update/delete/respond — including on IDs create_event
 *      had just returned. IDs are now percent-encoded (encodeGraphId).
 *
 * resolveMailboxTimeZone is deliberately NOT mocked here so the real mailbox
 * lookup is exercised end-to-end against a mocked Graph.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// A realistic Graph event ID: base64 with `+` and `=` but no `/` (so it passes
// assertOpaqueIds, exactly like the IDs Nate hit that still failed at Graph).
const EVENT_ID = 'AAMkAGI2gz6IS4+aCkhUgDjL9rVEUR==';
const ENCODED_EVENT_ID = encodeURIComponent(EVENT_ID); // AAMkAGI2gz6IS4%2BaCkhUgDjL9rVEUR%3D%3D
const MAILBOX_TZ = 'Pacific Standard Time';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();
const mockGetValidAccessTokenForSession = jest.fn<() => Promise<string>>();
const mockGetTenantIdFromSession = jest.fn<() => string>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<() => Promise<unknown[]>>();
const mockGetUserServiceOverrides = jest.fn<() => Promise<string[]>>();
const mockResolveDefaultCalendarId = jest.fn<() => Promise<string | null>>();
const mockResolveCalendarName = jest.fn<() => Promise<string>>();

interface GraphCall { path: string; verb: string; body?: unknown }
const graphCalls: GraphCall[] = [];

// Route a terminal Graph call to a canned response by path + verb.
function graphRespond(path: string, verb: string): unknown {
  if (path === '/me/mailboxSettings') return { timeZone: MAILBOX_TZ };
  if (verb === 'post' && path === '/me/events') {
    // create_event: echo the stored event, including the timeZone Graph kept.
    return {
      id: EVENT_ID,
      subject: 'TZ probe',
      start: { dateTime: '2026-09-09T17:00:00.0000000', timeZone: MAILBOX_TZ },
      end: { dateTime: '2026-09-09T17:15:00.0000000', timeZone: MAILBOX_TZ },
      webLink: 'https://outlook.office365.com/owa/?itemid=ev',
    };
  }
  if (verb === 'patch') return { id: EVENT_ID, subject: 'TZ probe', start: {}, end: {} };
  if (verb === 'delete') return undefined;
  // get_event / respond_to_event event lookups.
  return {
    id: EVENT_ID,
    subject: 'TZ probe',
    start: { dateTime: '2026-09-09T17:00:00.0000000', timeZone: MAILBOX_TZ },
    end: { dateTime: '2026-09-09T17:15:00.0000000', timeZone: MAILBOX_TZ },
    location: { displayName: null },
    organizer: { emailAddress: { address: 'me@x.com' } },
    attendees: [],
    isAllDay: false,
    body: { contentType: 'text', content: '' },
    webLink: 'https://outlook.office365.com/owa/?itemid=ev',
  };
}

function makeChain(path: string) {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header']) chain[m] = () => chain;
  chain.get = () => {
    graphCalls.push({ path, verb: 'get' });
    return Promise.resolve(graphRespond(path, 'get'));
  };
  chain.post = (body: unknown) => {
    graphCalls.push({ path, verb: 'post', body });
    return Promise.resolve(graphRespond(path, 'post'));
  };
  chain.patch = (body: unknown) => {
    graphCalls.push({ path, verb: 'patch', body });
    return Promise.resolve(graphRespond(path, 'patch'));
  };
  chain.delete = () => {
    graphCalls.push({ path, verb: 'delete' });
    return Promise.resolve(graphRespond(path, 'delete'));
  };
  return chain;
}
const mockCreateGraphClient = jest.fn(() => ({ api: (path: string) => makeChain(path) }));

// telemetry.js patches console at import time — stub it.
// The outbound policy is covered in outboundPolicy.test.ts; here it allows everything.
jest.mock('../services/outboundPolicy.js', () => ({
  ...jest.requireActual<typeof import('../services/outboundPolicy.js')>('../services/outboundPolicy.js'),
  enforceOutboundPolicy: () => Promise.resolve(),
  enforceEventUpdatePolicy: () => Promise.resolve(),
}));

jest.mock('../services/telemetry.js', () => ({}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => mockAuthenticateRequest(),
}));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => mockGetValidAccessTokenForSession(),
  getTenantIdFromSession: () => mockGetTenantIdFromSession(),
}));
jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => mockCreateGraphClient(),
}));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: () => mockIsPathDenied(),
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

const AUTH = { userId: 'u', session: { userId: 'u', tenantId: 't', accessToken: 'fake', sessionToken: 'sess' } };

function callTool(name: string, args: Record<string, unknown>) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

function toolResult(res: { jsonBody?: unknown }): { parsed: Record<string, any> | null; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : (JSON.parse(text) as Record<string, any>), isError, text };
}

beforeEach(() => {
  jest.clearAllMocks();
  graphCalls.length = 0;
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue('t');
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['calendar', 'mail']);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
  mockResolveDefaultCalendarId.mockResolvedValue('default-cal-id');
  mockResolveCalendarName.mockResolvedValue('Calendar');
});

const eventVerbCalls = () => graphCalls.filter((c) => c.path.includes('/events') || c.path.startsWith('/me/events'));

describe('create_event — time zone', () => {
  it('defaults a missing timeZone to the mailbox zone, not a hardcoded one', async () => {
    const res = await callTool('create_event', {
      subject: 'TZ probe', start: '2026-09-09T17:00:00', end: '2026-09-09T17:15:00',
    });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    // The mailbox lookup actually happened.
    expect(graphCalls.some((c) => c.path === '/me/mailboxSettings')).toBe(true);
    // The event was POSTed with the mailbox zone on both bounds.
    const post = graphCalls.find((c) => c.verb === 'post' && c.path === '/me/events');
    expect(post).toBeDefined();
    const body = post!.body as { start: { timeZone: string }; end: { timeZone: string } };
    expect(body.start.timeZone).toBe(MAILBOX_TZ);
    expect(body.end.timeZone).toBe(MAILBOX_TZ);
    // The echoed timeZone matches the mailbox — never America/New_York.
    expect(parsed!.start.timeZone).toBe(MAILBOX_TZ);
    expect(parsed!.start.timeZone).not.toBe('America/New_York');
  });

  it('honors an explicit timeZone and skips the mailbox lookup', async () => {
    await callTool('create_event', {
      subject: 'x', start: '2026-09-09T17:00:00', end: '2026-09-09T17:15:00', timeZone: 'America/Los_Angeles',
    });
    expect(graphCalls.some((c) => c.path === '/me/mailboxSettings')).toBe(false);
    const post = graphCalls.find((c) => c.verb === 'post' && c.path === '/me/events');
    const body = post!.body as { start: { timeZone: string } };
    expect(body.start.timeZone).toBe('America/Los_Angeles');
  });
});

describe('showAs — free/busy status', () => {
  it('create_event passes showAs through to the POST body when supplied', async () => {
    await callTool('create_event', {
      subject: 'soft hold', start: '2026-09-12T00:00:00', end: '2026-09-13T00:00:00',
      isAllDay: true, showAs: 'tentative',
    });
    const post = graphCalls.find((c) => c.verb === 'post' && c.path === '/me/events');
    expect(post).toBeDefined();
    expect((post!.body as { showAs?: string }).showAs).toBe('tentative');
  });

  it('create_event omits showAs entirely when not supplied (Graph default preserved)', async () => {
    await callTool('create_event', {
      subject: 'x', start: '2026-09-09T17:00:00', end: '2026-09-09T17:15:00',
    });
    const post = graphCalls.find((c) => c.verb === 'post' && c.path === '/me/events');
    expect(post).toBeDefined();
    expect(Object.prototype.hasOwnProperty.call(post!.body, 'showAs')).toBe(false);
  });

  it('update_event patches showAs when supplied', async () => {
    await callTool('update_event', { eventId: EVENT_ID, showAs: 'free' });
    const patch = graphCalls.find((c) => c.verb === 'patch');
    expect(patch).toBeDefined();
    expect((patch!.body as { showAs?: string }).showAs).toBe('free');
  });

  it('update_event omits showAs from the patch when not supplied', async () => {
    await callTool('update_event', { eventId: EVENT_ID, subject: 'renamed' });
    const patch = graphCalls.find((c) => c.verb === 'patch');
    expect(patch).toBeDefined();
    expect(Object.prototype.hasOwnProperty.call(patch!.body, 'showAs')).toBe(false);
  });
});

describe('event ID round-trip — get / update / delete / respond percent-encode the ID', () => {
  it('get_event addresses the encoded ID (no raw + in the path)', async () => {
    await callTool('get_event', { eventId: EVENT_ID });
    const get = graphCalls.find((c) => c.verb === 'get' && c.path.startsWith('/me/events/'));
    expect(get).toBeDefined();
    expect(get!.path).toBe(`/me/events/${ENCODED_EVENT_ID}`);
    expect(get!.path).toContain('%2B');
    expect(get!.path).not.toContain('+');
  });

  it('update_event addresses the encoded ID and defaults its timeZone to the mailbox', async () => {
    await callTool('update_event', { eventId: EVENT_ID, start: '2026-09-09T18:00:00', end: '2026-09-09T18:30:00' });
    const patch = graphCalls.find((c) => c.verb === 'patch');
    expect(patch!.path).toBe(`/me/events/${ENCODED_EVENT_ID}`);
    expect(patch!.path).not.toContain('+');
    const body = patch!.body as { start: { timeZone: string } };
    expect(body.start.timeZone).toBe(MAILBOX_TZ);
  });

  it('delete_event addresses the encoded ID', async () => {
    await callTool('delete_event', { eventId: EVENT_ID });
    const del = graphCalls.find((c) => c.verb === 'delete');
    expect(del!.path).toBe(`/me/events/${ENCODED_EVENT_ID}`);
    expect(del!.path).not.toContain('+');
  });

  it('respond_to_event addresses the encoded ID on both the lookup and the action', async () => {
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
    expect(toolResult(res).isError).toBe(false);
    // The initial event lookup and the accept POST both target the encoded ID.
    expect(graphCalls.some((c) => c.verb === 'get' && c.path === `/me/events/${ENCODED_EVENT_ID}`)).toBe(true);
    expect(graphCalls.some((c) => c.verb === 'post' && c.path === `/me/events/${ENCODED_EVENT_ID}/accept`)).toBe(true);
    // No path anywhere in the flow leaked a raw + that Graph would mis-decode.
    expect(eventVerbCalls().every((c) => !c.path.includes('+'))).toBe(true);
  });
});
