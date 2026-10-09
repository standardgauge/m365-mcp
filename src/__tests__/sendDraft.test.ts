/**
 * send_draft MCP tool tests.
 *
 * send_draft is the dispatch step was missing: reply_to_message and friends
 * create a draft and never send, and until now nothing on the MCP surface could
 * take a draft ID and deliver it.
 *
 * The load-bearing behaviour here is the email-output-mode gate. In 'draft' mode
 * (the default for every new user) the tool must REFUSE. A draft an
 * agent created moments earlier has not been reviewed by a human, so honouring the
 * send would turn the default mode into a no-op — an agent could chain
 * create_draft -> send_draft and deliver mail with nobody in the loop.
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
const mockGetUserEmailSettings = jest.fn<() => Promise<{ emailOutputMode: string }>>();

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
  getUserEmailSettings: () => mockGetUserEmailSettings(),
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
const DRAFT_ID = 'AAMkReplyDraftGuid002';
const CONVERSATION_ID = 'AAQkConversationGuid003';
const DRAFT_LINK = 'https://outlook.office.com/mail/drafts/id/reply';
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

const sendPosts = () => graphCalls.filter((c) => c.method === 'post' && c.path.endsWith('/send'));

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
  mockResolveMailFolderName.mockResolvedValue('Drafts');
  mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });
  mockGraphGet.mockResolvedValue({
    isDraft: true,
    parentFolderId: 'drafts-folder-id',
    subject: 'RE: threading test',
    webLink: DRAFT_LINK,
    conversationId: CONVERSATION_ID,
    toRecipients: [{ emailAddress: { address: 'nprodrom@gmail.example' } }],
    ccRecipients: [],
  });
  mockGraphPost.mockResolvedValue({});
});

// ─────────────────────────────────────────────────────────────────────────────
// toolManifest + tools/list
// ─────────────────────────────────────────────────────────────────────────────

describe('toolManifest — send_draft', () => {
  it('is registered as an MCP-native POST tool taking messageId, mailboxId and from', async () => {
    const { TOOLS } = await import('../mcp/toolManifest.js');
    const tool = TOOLS.find((t) => t.name === 'send_draft');
    expect(tool).toBeDefined();
    expect(tool?.endpoint).toBe('/api/mcp');
    expect(tool?.method).toBe('POST');
    expect(tool?.parameters.filter((p) => p.required).map((p) => p.name)).toEqual(['messageId']);
    expect(tool?.parameters.map((p) => p.name)).toEqual(['messageId', 'mailboxId', 'from']);
  });

  it('getManifest() and live tools/list expose the same input-schema shape', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string; inputSchema: Record<string, unknown> }> };
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string; inputSchema: Record<string, unknown> }> } };

    function shape(inputSchema: any) {
      const props: Record<string, unknown> = {};
      for (const [name, spec] of Object.entries(inputSchema.properties ?? {})) {
        const s = spec as any;
        props[name] = { type: s.type, ...(s.enum ? { enum: s.enum } : {}) };
      }
      return { type: inputSchema.type, required: [...(inputSchema.required ?? [])].sort(), props };
    }

    const fromManifest = manifest.tools.find((t) => t.name === 'send_draft');
    const fromList = body.result.tools.find((t) => t.name === 'send_draft');
    expect(fromManifest).toBeDefined();
    expect(fromList).toBeDefined();
    expect(shape(fromManifest!.inputSchema)).toEqual(shape(fromList!.inputSchema));
  });

  it('the reply tools now point callers at send_draft rather than a nonexistent send path', async () => {
    const res = await rpc('tools/list', {});
    const body = res.jsonBody as { result: { tools: Array<{ name: string; description: string }> } };
    for (const name of ['reply_to_message', 'reply_all_to_message']) {
      const tool = body.result.tools.find((t) => t.name === name);
      expect(tool!.description).toContain('send_draft');
      expect(tool!.description).not.toContain('send it with the draft ID');
    }
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Output-mode gate — the reason this tool needs care
// ─────────────────────────────────────────────────────────────────────────────

describe("MCP tools/call — send_draft output-mode gate", () => {
  it("sends when the user's output mode is 'send'", async () => {
    const res = await callTool('send_draft', { messageId: DRAFT_ID });
    const { parsed, isError } = toolResult(res);
    expect(isError).toBe(false);
    expect(sendPosts()).toHaveLength(1);
    expect(sendPosts()[0].path).toBe(`/me/messages/${DRAFT_ID}/send`);
    expect(sendPosts()[0].body).toEqual({});
    expect(parsed).toMatchObject({
      status: 'sent',
      messageId: DRAFT_ID,
      subject: 'RE: threading test',
      to: ['nprodrom@gmail.example'],
      conversationId: CONVERSATION_ID,
    });
  });

  it("refuses in 'draft' mode and points at the draft's webLink instead of sending", async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft' });
    const res = await callTool('send_draft', { messageId: DRAFT_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('sent by a person');
    expect(text).toContain(DRAFT_LINK);
    expect(sendPosts()).toHaveLength(0);
  });

  it("cannot be chained after create_draft to bypass 'draft' mode", async () => {
    // The bypass this gate exists to prevent: compose then immediately dispatch,
    // with no human between the two calls.
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft' });
    const created = await callTool('create_draft', { subject: 'Hi', to: ['x@example.com'], body: 'text' });
    expect(toolResult(created).isError).toBe(false);
    const sent = await callTool('send_draft', { messageId: DRAFT_ID });
    expect(toolResult(sent).isError).toBe(true);
    expect(sendPosts()).toHaveLength(0);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Draft-state and targeting
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — send_draft draft state and targeting', () => {
  it('refuses an already-sent message', async () => {
    mockGraphGet.mockResolvedValue({ isDraft: false, parentFolderId: 'sent-folder-id', subject: 'Sent thing' });
    const res = await callTool('send_draft', { messageId: DRAFT_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('not a draft');
    expect(sendPosts()).toHaveLength(0);
  });

  it('targets /users/{mailboxId}/... when a mailbox is given', async () => {
    const res = await callTool('send_draft', { messageId: DRAFT_ID, mailboxId: 'exec-assistant-user-id' });
    expect(toolResult(res).isError).toBe(false);
    expect(sendPosts()[0].path).toBe(`/users/exec-assistant-user-id/messages/${DRAFT_ID}/send`);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// Security enforcement
// ─────────────────────────────────────────────────────────────────────────────

describe('MCP tools/call — send_draft security', () => {
  it("is blocked when the draft's own folder is deny-listed", async () => {
    mockResolveMailFolderName.mockResolvedValue('HR Confidential');
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'HR Confidential');
    const res = await callTool('send_draft', { messageId: DRAFT_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('deny list');
    expect(sendPosts()).toHaveLength(0);
  });

  it('is blocked when Sent Items is deny-listed, since the sent copy lands there', async () => {
    mockIsPathDenied.mockImplementation(async (_t, _u, _type, path) => path === 'Sent Items');
    const res = await callTool('send_draft', { messageId: DRAFT_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('Sent Items folder is blocked');
    expect(sendPosts()).toHaveLength(0);
  });

  it('is refused when the mail service is in read-only mode', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['mail']);
    const res = await callTool('send_draft', { messageId: DRAFT_ID });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('read-only mode');
    expect(graphCalls).toHaveLength(0);
  });

  it('rejects a path-injection messageId before any Graph call (opaque-ID guard)', async () => {
    const res = await callTool('send_draft', { messageId: 'legit/../../../users/victim/messages/x' });
    const { isError, text } = toolResult(res);
    expect(isError).toBe(true);
    expect(text).toContain('Bad request');
    expect(graphCalls).toHaveLength(0);
  });

  it('is hidden from tools/list and rejected when the mail service is disabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['calendar']);
    const list = await rpc('tools/list', {});
    const body = list.jsonBody as { result: { tools: Array<{ name: string }> } };
    expect(body.result.tools.map((t) => t.name)).not.toContain('send_draft');
    const res = await callTool('send_draft', { messageId: DRAFT_ID });
    expect(toolResult(res).text).toContain('not enabled');
  });
});
