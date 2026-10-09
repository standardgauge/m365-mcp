/**
 * respond_to_event MCP tool tests.
 *
 * Two layers:
 *   1. toolManifest — the tool is registered in the TOOLS catalog with the
 *      right params/enum, and its getManifest() schema matches the live
 *      tools/list schema (MCP-native tool: endpoint /api/mcp).
 *   2. Remote MCP endpoint (/api/mcp JSON-RPC) — accept / tentative / decline
 *      hit the right Graph action paths, a meeting-invite MESSAGE ID resolves
 *      to its underlying event, and deny-list / read-only / opaque-ID
 *      enforcement is intact.
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
  method: 'get' | 'post';
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
  chain.delete = () => Promise.resolve(undefined);
  return chain;
}
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => makeChain(path),
}));

// telemetry.js has import-time side effects (patches console) — stub it out
// Routing tests: the delegated-mailbox owner lookup (its own Graph call) is
// covered by mailboxOwner.test.ts and delegatedDenyList.test.ts.
jest.mock('../services/mailboxOwner.js', () => ({
  resolveDenySubject: (_g: unknown, callerId: string) => Promise.resolve(callerId),
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
const EVENT_ID = 'AAMkEventGuid007';
const MESSAGE_ID = 'AAMkInviteMsgGuid009';
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

function graphNotFound(code = 'ErrorItemNotFound', statusCode = 404): Error {
  const err = new Error(code) as Error & { statusCode: number; code: string };
  err.statusCode = statusCode;
  err.code = code;
  return err;
}

const postedRespondCalls = () => graphCalls.filter((c) => c.method === 'post');

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
  // Default: the ID is a real event; message lookups aren't needed.
  mockGraphGet.mockImplementation(async (path) => {
    if (path.includes('/events/')) return { id: EVENT_ID };
    if (path.endsWith('/calendar')) return { id: 'other-default-cal', name: 'Calendar' };
    return {};
  });
  mockGraphPost.mockResolvedValue({});
});

// ─────────────────────────────────────────────────────────────────────────────
// toolManifest — registration + schema sync with tools/list
// ─────────────────────────────────────────────────────────────────────────────

describe('toolManifest — respond_to_event', () => {
  it('is registered as an MCP-native POST tool with the right params', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const tool = TOOLS.find((t) => t.name === 'respond_to_event');
    expect(tool).toBeDefined();
    expect(tool?.endpoint).toBe('/api/mcp');
    expect(tool?.method).toBe('POST');
    expect(tool?.parameters.filter((p) => p.required).map((p) => p.name)).toEqual(
      ['messageOrEventId', 'response'],
    );
    expect(tool?.parameters.find((p) => p.name === 'response')?.enum).toEqual(
      ['accept', 'tentative', 'decline'],
    );
    const optional = tool?.parameters.filter((p) => !p.required).map((p) => p.name);
    expect(optional).toEqual(expect.arrayContaining(['comment', 'sendResponse', 'mailboxId']));
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

    const fromManifest = manifest.tools.find((t) => t.name === 'respond_to_event');
    const fromList = body.result.tools.find((t) => t.name === 'respond_to_event');
    expect(fromManifest).toBeDefined();
    expect(fromList).toBeDefined();
    expect(shape(fromManifest!.inputSchema)).toEqual(shape(fromList!.inputSchema));
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/list — exposure follows the calendar service toggle
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/list — respond_to_event exposure', () => {
  it('is listed when the calendar service is enabled', async () => {
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    expect(body.result.tools.map((t) => t.name)).toContain('respond_to_event');
  });

  it('is hidden when the calendar service is disabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail']);
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    expect(body.result.tools.map((t) => t.name)).not.toContain('respond_to_event');
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — Graph action paths and request bodies
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — respond_to_event actions', () => {
  it('accept POSTs to /me/events/{id}/accept with sendResponse defaulted to true', async () => {
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'accepted', eventId: EVENT_ID, sendResponse: true });
    const posts = postedRespondCalls();
    expect(posts).toHaveLength(1);
    expect(posts[0].path).toBe(`/me/events/${EVENT_ID}/accept`);
    expect(posts[0].body).toEqual({ sendResponse: true });
  });

  it('tentative maps to the tentativelyAccept Graph action', async () => {
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'tentative' });
    const { parsed } = toolResult(res);
    expect(parsed).toMatchObject({ status: 'tentativelyAccepted' });
    expect(postedRespondCalls()[0].path).toBe(`/me/events/${EVENT_ID}/tentativelyAccept`);
  });

  it('decline maps to the decline Graph action', async () => {
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'decline' });
    const { parsed } = toolResult(res);
    expect(parsed).toMatchObject({ status: 'declined' });
    expect(postedRespondCalls()[0].path).toBe(`/me/events/${EVENT_ID}/decline`);
  });

  it('forwards comment and an explicit sendResponse=false', async () => {
    await callTool('respond_to_event', {
      messageOrEventId: EVENT_ID, response: 'decline', comment: 'Conflict, sorry', sendResponse: false,
    });
    expect(postedRespondCalls()[0].body).toEqual({ sendResponse: false, comment: 'Conflict, sorry' });
  });

  it('rejects an invalid response value before any Graph call', async () => {
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'maybe' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('Invalid response');
    expect(graphCalls).toHaveLength(0);
  });

  it('targets /users/{mailboxId}/... when a mailbox is given', async () => {
    const res = await callTool('respond_to_event', {
      messageOrEventId: EVENT_ID, response: 'accept', mailboxId: 'exec-assistant-user-id',
    });
    const { isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(graphCalls.map((c) => c.path)).toContain('/users/exec-assistant-user-id/calendar');
    expect(postedRespondCalls()[0].path).toBe(`/users/exec-assistant-user-id/events/${EVENT_ID}/accept`);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — meeting-invite message ID resolution
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — respond_to_event message-ID fallback', () => {
  it('resolves a meeting-invite message ID to its event and responds to that event', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path === `/me/events/${MESSAGE_ID}`) throw graphNotFound();
      if (path === `/me/messages/${MESSAGE_ID}`) return { id: MESSAGE_ID, event: { id: EVENT_ID } };
      return {};
    });
    const res = await callTool('respond_to_event', { messageOrEventId: MESSAGE_ID, response: 'accept' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ status: 'accepted', eventId: EVENT_ID });
    expect(postedRespondCalls()[0].path).toBe(`/me/events/${EVENT_ID}/accept`);
  });

  it('errors clearly when the message exists but is not a meeting invitation', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path === `/me/events/${MESSAGE_ID}`) throw graphNotFound();
      if (path === `/me/messages/${MESSAGE_ID}`) return { id: MESSAGE_ID };
      return {};
    });
    const res = await callTool('respond_to_event', { messageOrEventId: MESSAGE_ID, response: 'accept' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('not a meeting invitation');
    expect(postedRespondCalls()).toHaveLength(0);
  });

  it('errors clearly when the ID matches neither an event nor a message', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path.startsWith('/me/events/') || path.startsWith('/me/messages/')) throw graphNotFound();
      return {};
    });
    const res = await callTool('respond_to_event', { messageOrEventId: 'AAMkNoSuchId', response: 'accept' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('No calendar event or message found');
    expect(postedRespondCalls()).toHaveLength(0);
  });

  it('propagates non-404 Graph errors from the event lookup without a message fallback', async () => {
    mockGraphGet.mockImplementation(async (path) => {
      if (path.startsWith('/me/events/')) throw graphNotFound('ErrorAccessDenied', 403);
      return {};
    });
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('ErrorAccessDenied');
    expect(graphCalls.filter((c) => c.path.startsWith('/me/messages/'))).toHaveLength(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — security enforcement
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — respond_to_event security', () => {
  it('blocks the response when the default calendar ID is on the deny list (no Graph call)', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === DEFAULT_CAL_ID);
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(graphCalls).toHaveLength(0);
  });

  it('blocks the response when the calendar NAME is on the deny list', async () => {
    mockResolveCalendarName.mockResolvedValue('HR Sensitive Calendar');
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'HR Sensitive Calendar');
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
    const { isError } = toolResult(res);
    expect(isError).toBe(true);
    expect(graphCalls).toHaveLength(0);
  });

  it("enforces the deny list against an explicit mailbox's default calendar", async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'other-default-cal');
    const res = await callTool('respond_to_event', {
      messageOrEventId: EVENT_ID, response: 'accept', mailboxId: 'exec-assistant-user-id',
    });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('deny list');
    expect(postedRespondCalls()).toHaveLength(0);
  });

  it('is refused when the calendar service is in read-only mode (write tool)', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['calendar']);
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('read-only mode');
    expect(graphCalls).toHaveLength(0);
  });

  it('rejects a path-injection messageOrEventId before any Graph call (opaque-ID guard)', async () => {
    const res = await callTool('respond_to_event', {
      messageOrEventId: 'legit/../../../users/victim/events/x', response: 'accept',
    });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('Bad request');
    expect(graphCalls).toHaveLength(0);
  });

  it('is rejected when the calendar service is not enabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail']);
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('not enabled');
  });
});
