/**
 * reply_to_message / reply_all_to_message / forward_message MCP tool tests.
 *
 * Three layers:
 *   1. toolManifest — the tools are registered in the TOOLS catalog and their
 *      getManifest() schemas match the live tools/list schemas.
 *   2. Remote MCP endpoint (/api/mcp JSON-RPC) — each verb hits the right Graph
 *      createReply / createReplyAll / createForward path, the `comment` is
 *      rendered per bodyType, and the returned draft surfaces the threading
 *      fields (conversationId, resolved recipients).
 *   3. Enforcement — source-folder and Drafts deny lists, read-only mode, and
 *      the opaque-ID guard on messageId.
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
const mockResolveMailFolderName = jest.fn<() => Promise<string | null>>();

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
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'header', 'search']) chain[m] = () => chain;
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
  filterDeniedSearchHits: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: (...args: unknown[]) =>
    mockIsPathDenied(args[0] as string, args[1] as string, args[2] as string, args[3] as string),
}));

jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => mockResolveMailFolderName(),
  resolveDefaultCalendarId: () => Promise.resolve('cal'),
  resolveCalendarName: () => Promise.resolve('Calendar'),
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
const MESSAGE_ID = 'AAMkInboundMsgGuid001';
const DRAFT_ID = 'AAMkReplyDraftGuid002';
const CONVERSATION_ID = 'AAQkConversationGuid003';
const AUTH = { userId: USER, session: { userId: USER, tenantId: TENANT, accessToken: 'fake', sessionToken: 'sess' } };

const REPLY_TOOLS = ['reply_to_message', 'reply_all_to_message', 'forward_message'] as const;

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

// Valid minimal args per tool: forward_message requires a `to` recipient list,
// while the reply tools deliberately do not accept `to` ( rejects it).
const withRecipients = (name: string): Record<string, unknown> =>
  name === 'forward_message'
    ? { messageId: MESSAGE_ID, comment: 'hi', to: ['x@example.com'] }
    : { messageId: MESSAGE_ID, comment: 'hi' };

function toolResult(res: { jsonBody?: unknown }): { parsed: any; isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

const posts = () => graphCalls.filter((c) => c.method === 'post');

beforeEach(() => {
  jest.clearAllMocks();
  graphCalls.length = 0;
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockGetValidAccessTokenForSession.mockResolvedValue('access-token');
  mockGetTenantIdFromSession.mockReturnValue(TENANT);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetEnabledServices.mockResolvedValue(['mail', 'calendar', 'sharepoint']);
  mockGetReadOnlyServices.mockResolvedValue([]);
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetUserServiceOverrides.mockResolvedValue([]);
  mockResolveMailFolderName.mockResolvedValue('Inbox');
  mockGraphGet.mockResolvedValue({ parentFolderId: 'inbox-folder-id' });
  mockGraphPost.mockResolvedValue({
    id: DRAFT_ID,
    subject: 'RE: Onboarding paperwork',
    webLink: 'https://outlook.office.com/mail/drafts/id/reply',
    conversationId: CONVERSATION_ID,
    toRecipients: [{ emailAddress: { address: 'norleen@bakertilly.example' } }],
    ccRecipients: [{ emailAddress: { address: 'billing@bakertilly.example' } }],
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// toolManifest — registration + schema sync with tools/list
// ─────────────────────────────────────────────────────────────────────────────

describe('toolManifest — reply/forward tools', () => {
  it.each(REPLY_TOOLS)('%s is registered as an MCP-native POST tool', async (name) => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const tool = TOOLS.find((t) => t.name === name);
    expect(tool).toBeDefined();
    expect(tool?.endpoint).toBe('/api/mcp');
    expect(tool?.method).toBe('POST');
    expect(tool?.parameters.find((p) => p.name === 'bodyType')?.enum).toEqual(['text', 'html']);
    // Recipients are derived from the source message — reply tools must not take `to`.
    const paramNames = tool?.parameters.map((p) => p.name);
    expect(paramNames).toEqual(expect.arrayContaining(['messageId', 'comment', 'mailboxId']));
    if (name === 'forward_message') expect(paramNames).toContain('to');
    else expect(paramNames).not.toContain('to');
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

    for (const name of REPLY_TOOLS) {
      const fromManifest = manifest.tools.find((t) => t.name === name);
      const fromList = body.result.tools.find((t) => t.name === name);
      expect(fromManifest).toBeDefined();
      expect(fromList).toBeDefined();
      expect(shape(fromManifest!.inputSchema)).toEqual(shape(fromList!.inputSchema));
    }
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/list — exposure follows the mail service toggle
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/list — reply/forward exposure', () => {
  it('lists all three when the mail service is enabled', async () => {
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    expect(body.result.tools.map((t) => t.name)).toEqual(expect.arrayContaining([...REPLY_TOOLS]));
  });

  it('hides them when the mail service is disabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['calendar']);
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string }> } };
    const names = body.result.tools.map((t) => t.name);
    for (const name of REPLY_TOOLS) expect(names).not.toContain(name);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — Graph action paths and request bodies
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — reply/forward Graph actions', () => {
  it('reply_to_message POSTs to createReply and returns the draft with threading fields', async () => {
    const res = await callTool('reply_to_message', { messageId: MESSAGE_ID, comment: 'Sounds good, thanks.' });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(posts()).toHaveLength(1);
    expect(posts()[0].path).toBe(`/me/messages/${MESSAGE_ID}/createReply`);
    expect(posts()[0].body).toEqual({ comment: 'Sounds good, thanks.' });
    expect(parsed).toMatchObject({
      id: DRAFT_ID,
      status: 'draft',
      conversationId: CONVERSATION_ID,
      inReplyTo: MESSAGE_ID,
      to: ['norleen@bakertilly.example'],
      cc: ['billing@bakertilly.example'],
    });
  });

  it('reply_all_to_message POSTs to createReplyAll', async () => {
    const res = await callTool('reply_all_to_message', { messageId: MESSAGE_ID, comment: 'Adding the team.' });
    expect(toolResult(res).isError).toBe(false);
    expect(posts()[0].path).toBe(`/me/messages/${MESSAGE_ID}/createReplyAll`);
  });

  it('forward_message POSTs to createForward with toRecipients', async () => {
    const res = await callTool('forward_message', {
      messageId: MESSAGE_ID, comment: 'FYI', to: ['amy@nrgclean.example', 'ops@nrgclean.example'],
    });
    expect(toolResult(res).isError).toBe(false);
    expect(posts()[0].path).toBe(`/me/messages/${MESSAGE_ID}/createForward`);
    expect(posts()[0].body).toEqual({
      comment: 'FYI',
      toRecipients: [
        { emailAddress: { address: 'amy@nrgclean.example' } },
        { emailAddress: { address: 'ops@nrgclean.example' } },
      ],
    });
  });

  it('forward_message rejects an empty recipient list before any Graph call', async () => {
    const res = await callTool('forward_message', { messageId: MESSAGE_ID, comment: 'FYI', to: [] });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('at least one recipient');
    expect(graphCalls).toHaveLength(0);
  });

  it('targets /users/{mailboxId}/... when a mailbox is given', async () => {
    const res = await callTool('reply_to_message', {
      messageId: MESSAGE_ID, comment: 'On it.', mailboxId: 'exec-assistant-user-id',
    });
    expect(toolResult(res).isError).toBe(false);
    expect(posts()[0].path).toBe(`/users/exec-assistant-user-id/messages/${MESSAGE_ID}/createReply`);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — comment rendering
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — comment rendering', () => {
  it('escapes markup and converts newlines to <br> for the default text bodyType', async () => {
    await callTool('reply_to_message', {
      messageId: MESSAGE_ID,
      comment: 'Line one\nLine two\r\nQ & A <not a tag>',
    });
    expect(posts()[0].body).toEqual({
      comment: 'Line one<br>Line two<br>Q &amp; A &lt;not a tag&gt;',
    });
  });

  it('passes the comment through verbatim when bodyType is html', async () => {
    await callTool('reply_to_message', {
      messageId: MESSAGE_ID, comment: '<p>Hi <b>Norleen</b></p>', bodyType: 'html',
    });
    expect(posts()[0].body).toEqual({ comment: '<p>Hi <b>Norleen</b></p>' });
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// MCP tools/call — security enforcement
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — reply/forward security', () => {
  it.each(REPLY_TOOLS)('%s is blocked when the source message folder is deny-listed', async (name) => {
    mockResolveMailFolderName.mockResolvedValue('HR Confidential');
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'HR Confidential');
    const res = await callTool(name, withRecipients(name));
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('deny list');
    expect(posts()).toHaveLength(0);
  });

  it.each(REPLY_TOOLS)('%s is blocked when the Drafts folder is deny-listed', async (name) => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'Drafts');
    const res = await callTool(name, withRecipients(name));
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('Drafts folder is blocked');
    expect(posts()).toHaveLength(0);
  });

  it.each(REPLY_TOOLS)('%s is refused when the mail service is in read-only mode', async (name) => {
    mockGetReadOnlyServices.mockResolvedValue(['mail']);
    const res = await callTool(name, withRecipients(name));
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('read-only mode');
    expect(graphCalls).toHaveLength(0);
  });

  it.each(REPLY_TOOLS)('%s rejects a path-injection messageId (opaque-ID guard)', async (name) => {
    const res = await callTool(name, {
      messageId: 'legit/../../../users/victim/messages/x', comment: 'hi',
    });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('Bad request');
    expect(graphCalls).toHaveLength(0);
  });

  it.each(REPLY_TOOLS)('%s is rejected when the mail service is not enabled', async (name) => {
    mockGetEnabledServices.mockResolvedValue(['calendar']);
    const res = await callTool(name, withRecipients(name));
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('not enabled');
  });
});
