/**
 * What a failing tools/call shows the MCP client.
 *
 * Messages written for the agent (policy refusals, argument errors, Graph 4xx
 * the user can act on) reach it as they are. Anything else (a Graph 5xx, a
 * storage or network error, a bug) can carry Graph or provider internals, so
 * the client gets a generic message with a reference and the detail is logged
 * server-side under that reference. The denied-audit classifier is unchanged.
 *
 * Driven through respond_to_event, whose event lookup rethrows any non-404
 * Graph error as-is. All Graph, auth, and storage dependencies are mocked.
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
import { GraphError } from '@microsoft/microsoft-graph-client';
import { logAccess } from '../services/auditLog.js';
import { clientFacingMessage, ToolCallError } from '../services/toolErrors.js';
import { ValidationError } from '../services/opaqueId.js';
import { OutboundPolicyError } from '../services/outboundPolicy.js';

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

let ctx: InvocationContext & { error: jest.Mock };

function rpc(method: string, params: Record<string, unknown>): Promise<{ status: number; jsonBody?: unknown }> {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method, params }),
  } as unknown as HttpRequest;
  ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext & { error: jest.Mock };
  return handler(req, ctx);
}

const callTool = (name: string, args: Record<string, unknown>) => rpc('tools/call', { name, arguments: args });

function toolResult(res: { jsonBody?: unknown }): { parsed: unknown; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

const mockLogAccess = logAccess as unknown as jest.Mock;

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
  mockGraphGet.mockResolvedValue({ id: EVENT_ID });
  mockGraphPost.mockResolvedValue({});
});

/** Make the event lookup throw `err`, call the tool, and return what the client and the log saw. */
async function failWith(err: unknown) {
  mockGraphGet.mockImplementation(async (path) => {
    if (path.startsWith('/me/events/')) throw err;
    return {};
  });
  const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
  const logged = ctx.error.mock.calls.map((c) => c.map(String).join(' ')).join('\n');
  return { ...toolResult(res), logged };
}

const REFERENCE_RE = /Reference: ([0-9a-f-]{36})\./;

const deniedAudits = () =>
  mockLogAccess.mock.calls.filter((c) => (c[0] as { result?: string }).result === 'denied');

describe('tools/call — internal failures are normalised', () => {
  it('replaces a Graph 5xx with a generic message and logs the detail under the same reference', async () => {
    const internal = Object.assign(
      new GraphError(503, 'Backend store shard-17.mailbox.internal.example.com unavailable'),
      { code: 'ServiceUnavailable' },
    );
    const { isError, text, logged } = await failWith(internal);

    expect(isError).toBe(true);
    expect(text).not.toContain('shard-17');
    expect(text).not.toContain('ServiceUnavailable');
    expect(text).toMatch(/^Error: The tool call failed because of an internal server error\. Reference: /);
    const reference = text.match(REFERENCE_RE)?.[1];
    expect(reference).toBeDefined();

    expect(logged).toContain(`[ref ${reference}]`);
    expect(logged).toContain('shard-17.mailbox.internal.example.com');
    expect(logged).toContain('status=503');
    expect(logged).toContain('code=ServiceUnavailable');
    expect(deniedAudits()).toHaveLength(0);
  });

  it('replaces a network or provider error', async () => {
    const { text, logged } = await failWith(new Error('connect ECONNREFUSED 10.20.30.40:443'));
    expect(text).not.toContain('10.20.30.40');
    expect(text).toMatch(REFERENCE_RE);
    expect(logged).toContain('ECONNREFUSED 10.20.30.40:443');
  });

  it('does not take a non-Graph error with a 4xx status for a Graph one', async () => {
    // A storage error carries statusCode too; its text names storage internals.
    const storage = Object.assign(new Error('The specified entity already exists. RequestId:abc'), {
      name: 'RestError', statusCode: 409, code: 'EntityAlreadyExists',
    });
    const { text } = await failWith(storage);
    expect(text).not.toContain('EntityAlreadyExists');
    expect(text).toMatch(REFERENCE_RE);
  });

  it('logs every cause in the chain', async () => {
    const wrapped = new Error('lookup failed', { cause: new Error('socket hang up') });
    const { text, logged } = await failWith(wrapped);
    expect(text).not.toContain('socket hang up');
    expect(logged).toContain('lookup failed <- caused by Error: socket hang up');
  });

  it('replaces a plain Error whose text happens to contain a policy phrase', async () => {
    const { text, logged } = await failWith(new Error('backend feature is not enabled on shard-17.internal.example.com'));
    expect(text).not.toContain('shard-17');
    expect(text).not.toContain('is not enabled');
    expect(text).toMatch(REFERENCE_RE);
    expect(logged).toContain('shard-17.internal.example.com');
  });

  it('gives each failure its own reference', async () => {
    const a = (await failWith(new Error('one'))).text.match(REFERENCE_RE)?.[1];
    const b = (await failWith(new Error('two'))).text.match(REFERENCE_RE)?.[1];
    expect(a).toBeDefined();
    expect(a).not.toBe(b);
  });
});

describe('tools/call — client-facing failures reach the agent', () => {
  it('passes a Graph 4xx the user can act on, with its status and code', async () => {
    const forbidden = Object.assign(new GraphError(403, 'Access is denied. Check credentials and try again.'), {
      code: 'ErrorAccessDenied',
    });
    const { text } = await failWith(forbidden);
    expect(text).toBe('Error: Microsoft Graph returned 403 ErrorAccessDenied: Access is denied. Check credentials and try again.');
  });

  it('passes Graph throttling through so the agent can back off', async () => {
    const throttled = Object.assign(new GraphError(429, 'Application is over its MailboxConcurrency limit.'), {
      code: 'ApplicationThrottled',
    });
    const { text } = await failWith(throttled);
    expect(text).toContain('Microsoft Graph returned 429 ApplicationThrottled');
  });

  it('passes an argument error a handler wrote for the caller', async () => {
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'maybe' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toBe('Error: Invalid response "maybe" — use accept, tentative, or decline');
  });

  it('passes a deny-list refusal and still audits it as denied', async () => {
    mockIsPathDenied.mockResolvedValue(true);
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toMatch(/restricted by (the )?deny list/);
    expect(text).not.toMatch(REFERENCE_RE);
    expect(deniedAudits()).toHaveLength(1);
    expect(deniedAudits()[0][0]).toMatchObject({ operation: 'respond_to_event', source: 'mcp' });
  });

  it('answers a re-authentication failure with the login link and no internal detail', async () => {
    mockGetValidAccessTokenForSession.mockRejectedValue(
      new Error('Token refresh failed for user 00000000-aaaa-bbbb-cccc-000000000000. Re-authentication required.'),
    );
    const res = await callTool('respond_to_event', { messageOrEventId: EVENT_ID, response: 'accept' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toMatch(/^Session expired\. Please re-authenticate at: /);
    expect(text).not.toContain('00000000-aaaa');
  });
});

describe('clientFacingMessage', () => {
  it('allows the error classes written for the caller', () => {
    expect(clientFacingMessage(new ToolCallError('pass `since`'))).toBe('pass `since`');
    expect(clientFacingMessage(new ValidationError('itemId must be a non-empty string'))).toBe('itemId must be a non-empty string');
    expect(clientFacingMessage(new OutboundPolicyError('Blocked by policy'))).toBe('Blocked by policy');
  });

  it('refuses a plain Error, a non-Error, and Graph statuses the user cannot act on', () => {
    expect(clientFacingMessage(new Error('anything'))).toBeNull();
    expect(clientFacingMessage('a string')).toBeNull();
    expect(clientFacingMessage(new GraphError(401, 'InvalidAuthenticationToken'))).toBeNull();
    expect(clientFacingMessage(new GraphError(500, 'Internal'))).toBeNull();
    expect(clientFacingMessage(new GraphError(-1, 'no status'))).toBeNull();
  });

  it('judges by type, not by wording: a policy phrase in a plain Error or Graph 5xx is still internal', () => {
    expect(clientFacingMessage(new Error('backend feature is not enabled on shard-17.internal'))).toBeNull();
    expect(clientFacingMessage(new Error('Access restricted by deny list'))).toBeNull();
    expect(clientFacingMessage(new GraphError(500, 'Feature is not enabled on shard 7'))).toBeNull();
    expect(clientFacingMessage(new ToolCallError('Service "teams" is not enabled for this tenant'))).toContain('is not enabled');
  });
});
