/**
 * The outbound policy on the tool surfaces, end to end through the real
 * policy service (only table storage and Graph are mocked).
 *
 * Covers:
 *   - create_event: blocked with attendees (no Graph write, refusal audited as
 *     denied); allowed without attendees; internal refuses an external attendee
 *   - update_event: blocked for an organizer edit to a meeting with attendees
 *   - move_event: force=true does not get past a block
 *   - respond_to_event: a comment is refused under block; a bare RSVP is not
 *   - send_chat_message: internal refuses a chat with a member from another tenant
 *   - send_channel_message: block refuses without reading the member list;
 *     internal checks /allMembers, catching an indirect shared-channel member
 *   - an unreadable policy row or domain list refuses with 403 / denied audit
 *   - REST POST /api/calendar/events: 403 under block, nothing written
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── Policy rows (OutboundPolicy table) ───────────────────────────────────────

let tenantRow: Record<string, unknown> | undefined;
let storageDown = false;
jest.mock('@azure/data-tables', () => ({
  TableClient: {
    fromConnectionString: () => ({
      getEntity: async (_pk: string, rk: string) => {
        if (storageDown) throw { statusCode: 503, message: 'storage unavailable' };
        if (rk === '__tenant__' && tenantRow) return tenantRow;
        throw { statusCode: 404 };
      },
      upsertEntity: async () => undefined,
      createTable: async () => undefined,
    }),
  },
}));

// ── Graph ────────────────────────────────────────────────────────────────────

let getResponses: Record<string, unknown> = {};
const graphWrites: Array<{ method: string; path: string; body?: unknown }> = [];
const graphGets: string[] = [];
function makeChain(path: string) {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header']) chain[m] = () => chain;
  chain.get = async () => {
    graphGets.push(path);
    if (path in getResponses) return getResponses[path];
    return { value: [] };
  };
  chain.post = async (body?: unknown) => { graphWrites.push({ method: 'POST', path, body }); return { id: 'new-id' }; };
  chain.patch = async (body?: unknown) => { graphWrites.push({ method: 'PATCH', path, body }); return { id: 'ev1' }; };
  chain.delete = async () => { graphWrites.push({ method: 'DELETE', path }); };
  return chain;
}
const fakeGraph = { api: (path: string) => makeChain(path) };

const TENANT = 'test-tenant';
const USER = 'test-user';
const AUTH = { userId: USER, session: { userId: USER, tenantId: TENANT, accessToken: 'fake', sessionToken: 'sess' } };
const mockLogAccess = jest.fn<(entry: Record<string, unknown>) => void>();

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({ authenticateRequest: async () => AUTH }));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: async () => 'access-token',
  getTenantIdFromSession: () => TENANT,
  getTenantId: async () => TENANT,
}));
jest.mock('../services/mailboxTimeZone.js', () => ({ resolveMailboxTimeZone: async () => 'Pacific Standard Time' }));
jest.mock('../services/graphClient.js', () => ({ createGraphClient: () => fakeGraph }));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: async () => false,
}));
jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: async () => null,
  resolveDefaultCalendarId: async () => 'default-cal-id',
  resolveCalendarName: async () => 'Calendar',
  resolveContactParentFolder: async () => null,
  resolveDefaultContactFolder: async () => null,
  resolveSectionNotebook: async () => null,
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
}));
jest.mock('../services/sharepointFilter.js', () => ({ filterAndDisambiguateSites: (s: unknown) => s }));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: async () => ['calendar', 'mail', 'teams'],
  getReadOnlyServices: async () => [],
  getAllowedSites: async () => [],
}));
jest.mock('../services/userServiceOverrides.js', () => ({ getUserServiceOverrides: async () => [] }));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: async () => ({ emailOutputMode: 'draft' }),
  setUserEmailSettings: async () => undefined,
}));
jest.mock('../services/userMailConfig.js', () => ({ isMailIndexingDisabled: async () => false }));
jest.mock('../services/auditLog.js', () => ({ logAccess: (e: Record<string, unknown>) => mockLogAccess(e) }));
// The REST route's wrapper checks are covered elsewhere; here it hands straight to the handler.
jest.mock('../services/policyEnforcement.js', () => ({
  withPolicyEnforcement: (_svc: unknown, handler: (r: unknown, c: unknown, a: unknown) => unknown) =>
    (req: unknown, ctx: unknown) => handler(req, ctx, AUTH),
}));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';
import '../functions/calendar/createEvent.js';
import { clearTenantDomainCache, OUTBOUND_POLICY_MARKER } from '../services/outboundPolicy.js';

type Handler = (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: { handler: Handler }) => void>;
function registered(name: string): Handler {
  const call = httpMock.mock.calls.find((c) => c[0] === name);
  if (!call) throw new Error(`${name} was not registered`);
  return call[1].handler;
}
const mcp = registered('mcpEndpoint');
const restCreateEvent = registered('createEvent');

const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;

async function callTool(name: string, args: Record<string, unknown>) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: async () => ({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
  } as unknown as HttpRequest;
  const res = await mcp(req, ctx);
  const body = res.jsonBody as { result: { content: Array<{ text: string }>; isError?: boolean } };
  return { text: body.result.content[0].text, isError: body.result.isError === true };
}

const ORG = { value: [{ verifiedDomains: [{ name: 'contoso.com' }] }] };
const START = '2026-11-02T10:00:00';
const END = '2026-11-02T10:30:00';

beforeEach(() => {
  jest.clearAllMocks();
  clearTenantDomainCache();
  tenantRow = undefined;
  storageDown = false;
  getResponses = { '/organization': ORG };
  graphWrites.length = 0;
  graphGets.length = 0;
});

function expectDeniedAudit(operation: string) {
  expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation, result: 'denied' }));
  expect(mockLogAccess).not.toHaveBeenCalledWith(expect.objectContaining({ operation, result: 'allowed' }));
}

describe('create_event', () => {
  it('is refused under block when it has attendees, writes nothing, and is audited as denied', async () => {
    tenantRow = { calendarInvites: 'block' };
    const r = await callTool('create_event', { subject: 's', start: START, end: END, attendees: ['pat@contoso.com'] });
    expect(r.isError).toBe(true);
    expect(r.text).toContain(OUTBOUND_POLICY_MARKER);
    expect(graphWrites).toEqual([]);
    expectDeniedAudit('create_event');
  });

  it('is allowed under block when there are no attendees', async () => {
    tenantRow = { calendarInvites: 'block' };
    const r = await callTool('create_event', { subject: 's', start: START, end: END });
    expect(r.isError).toBe(false);
    expect(graphWrites).toEqual([expect.objectContaining({ method: 'POST', path: '/me/events' })]);
  });

  it('under internal, allows verified-domain attendees and refuses an outside one', async () => {
    tenantRow = { calendarInvites: 'internal' };
    const ok = await callTool('create_event', { subject: 's', start: START, end: END, attendees: ['pat@contoso.com'] });
    expect(ok.isError).toBe(false);
    graphWrites.length = 0;
    const bad = await callTool('create_event', { subject: 's', start: START, end: END, attendees: ['pat@contoso.com', 'eve@fabrikam.com'] });
    expect(bad.isError).toBe(true);
    expect(bad.text).toContain('eve@fabrikam.com');
    expect(graphWrites).toEqual([]);
  });

  it('is unaffected when no policy is set', async () => {
    const r = await callTool('create_event', { subject: 's', start: START, end: END, attendees: ['eve@fabrikam.com'] });
    expect(r.isError).toBe(false);
    expect(graphGets).not.toContain('/organization');
  });
});

describe('update_event', () => {
  it('is refused under block for an organizer edit to a meeting with attendees', async () => {
    tenantRow = { calendarInvites: 'block' };
    getResponses['/me/events/ev1'] = { isOrganizer: true, attendees: [{ emailAddress: { address: 'pat@contoso.com' } }] };
    const r = await callTool('update_event', { eventId: 'ev1', body: 'new agenda' });
    expect(r.isError).toBe(true);
    expect(graphWrites).toEqual([]);
  });
});

describe('move_event', () => {
  it('force=true does not get past a block', async () => {
    tenantRow = { calendarInvites: 'block' };
    getResponses['/me/events/ev1'] = {
      id: 'ev1', subject: 's', start: {}, end: {}, isOrganizer: true, type: 'singleInstance',
      attendees: [{ emailAddress: { address: 'pat@contoso.com' }, type: 'required' }],
    };
    const r = await callTool('move_event', { eventId: 'ev1', targetCalendarId: 'other-cal', force: true });
    expect(r.isError).toBe(true);
    expect(r.text).toContain(OUTBOUND_POLICY_MARKER);
    expect(graphWrites).toEqual([]);
  });
});

describe('respond_to_event', () => {
  beforeEach(() => {
    tenantRow = { eventResponses: 'block' };
    getResponses['/me/events/ev1'] = { id: 'ev1', organizer: { emailAddress: { address: 'org@fabrikam.com' } } };
  });

  it('refuses a comment under block', async () => {
    const r = await callTool('respond_to_event', { messageOrEventId: 'ev1', response: 'accept', comment: 'see attached notes' });
    expect(r.isError).toBe(true);
    expect(graphWrites).toEqual([]);
    expectDeniedAudit('respond_to_event');
  });

  it('allows a bare RSVP, and a comment that is not sent', async () => {
    expect((await callTool('respond_to_event', { messageOrEventId: 'ev1', response: 'accept' })).isError).toBe(false);
    expect((await callTool('respond_to_event', { messageOrEventId: 'ev1', response: 'decline', comment: 'x', sendResponse: false })).isError).toBe(false);
    expect(graphWrites).toHaveLength(2);
  });
});

describe('Teams sends', () => {
  it('send_chat_message under internal refuses a chat with a member from another tenant', async () => {
    tenantRow = { teamsMessages: 'internal' };
    getResponses['/chats/chat1/members'] = {
      value: [
        { email: 'me@contoso.com', tenantId: TENANT },
        { email: 'guest@contoso.com', tenantId: 'other-tenant' },
      ],
    };
    const r = await callTool('send_chat_message', { chatId: 'chat1', content: 'hi' });
    expect(r.isError).toBe(true);
    expect(graphWrites).toEqual([]);
  });

  it('send_chat_message under internal allows an all-internal chat', async () => {
    tenantRow = { teamsMessages: 'internal' };
    getResponses['/chats/chat1/members'] = { value: [{ email: 'me@contoso.com', tenantId: TENANT }] };
    const r = await callTool('send_chat_message', { chatId: 'chat1', content: 'hi' });
    expect(r.isError).toBe(false);
    expect(graphWrites).toEqual([expect.objectContaining({ path: '/chats/chat1/messages' })]);
  });

  it('send_channel_message under internal checks allMembers, so an indirect shared-channel member refuses', async () => {
    tenantRow = { teamsMessages: 'internal' };
    // Direct members are all internal; the external tenant reaches the channel
    // only through a team it is shared with, which `/members` would not list.
    getResponses['/teams/team1/channels/chan1/members'] = { value: [{ email: 'me@contoso.com', tenantId: TENANT }] };
    getResponses['/teams/team1/channels/chan1/allMembers'] = {
      value: [
        { email: 'me@contoso.com', tenantId: TENANT },
        { email: 'pat@fabrikam.com', tenantId: 'other-tenant' },
      ],
    };
    const r = await callTool('send_channel_message', { teamId: 'team1', channelId: 'chan1', content: 'hi' });
    expect(r.isError).toBe(true);
    expect(r.text).toContain('pat@fabrikam.com');
    expect(graphGets).toContain('/teams/team1/channels/chan1/allMembers');
    expect(graphWrites).toEqual([]);
    expectDeniedAudit('send_channel_message');
  });

  it('send_channel_message under block refuses without reading members', async () => {
    tenantRow = { teamsMessages: 'block' };
    const r = await callTool('send_channel_message', { teamId: 'team1', channelId: 'chan1', content: 'hi' });
    expect(r.isError).toBe(true);
    expect(graphGets.some((p) => p.endsWith('/members'))).toBe(false);
    expect(graphWrites).toEqual([]);
  });
});

describe('REST POST /api/calendar/events', () => {
  it('returns 403 under block and writes nothing', async () => {
    tenantRow = { calendarInvites: 'block' };
    const req = {
      method: 'POST',
      headers: new Map<string, string>(),
      query: new URLSearchParams(),
      json: async () => ({ subject: 's', start: START, end: END, attendees: ['pat@contoso.com'] }),
    } as unknown as HttpRequest;
    const res = await restCreateEvent(req, ctx);
    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain(OUTBOUND_POLICY_MARKER);
    expect(graphWrites).toEqual([]);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'calendar.post', result: 'denied' }));
  });

  it('returns 403 and audits a denial when the policy cannot be read', async () => {
    storageDown = true;
    const req = {
      method: 'POST',
      headers: new Map<string, string>(),
      query: new URLSearchParams(),
      json: async () => ({ subject: 's', start: START, end: END, attendees: ['pat@contoso.com'] }),
    } as unknown as HttpRequest;
    const res = await restCreateEvent(req, ctx);
    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain(OUTBOUND_POLICY_MARKER);
    expect(graphWrites).toEqual([]);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'calendar.post', result: 'denied' }));
  });

  it('returns 403 and audits a denial when the domain list cannot be read', async () => {
    tenantRow = { calendarInvites: 'internal' };
    getResponses['/organization'] = { value: [] };
    const req = {
      method: 'POST',
      headers: new Map<string, string>(),
      query: new URLSearchParams(),
      json: async () => ({ subject: 's', start: START, end: END, attendees: ['pat@contoso.com'] }),
    } as unknown as HttpRequest;
    const res = await restCreateEvent(req, ctx);
    expect(res.status).toBe(403);
    expect(graphWrites).toEqual([]);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'calendar.post', result: 'denied' }));
  });
});

describe('fail closed through the MCP dispatcher', () => {
  it('an unreadable policy refuses create_event and is audited as denied', async () => {
    storageDown = true;
    const r = await callTool('create_event', { subject: 's', start: START, end: END, attendees: ['pat@contoso.com'] });
    expect(r.isError).toBe(true);
    expect(r.text).toContain(OUTBOUND_POLICY_MARKER);
    expect(graphWrites).toEqual([]);
    expectDeniedAudit('create_event');
  });
});
