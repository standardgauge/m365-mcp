/**
 * search_mail / list_messages through the native MCP dispatch.
 *
 * The service routing itself is covered in mailSearch.test.ts; these tests pin
 * what the tools/call surface adds on top of it:
 *
 *   1. The schema declares the new criteria (participant / from / to / since)
 *      on both discovery surfaces, `q` is optional, and the arg check
 *      therefore accepts them.
 *   2. Deny-list enforcement: a denied folder blocks both tools up front; a
 *      mailbox-wide result set is post-filtered per message.
 *   3. The response carries the envelope plus the metadata
 *      (strategy / ordering / searchedFields / scan horizon / notes).
 *   4. list_messages is $orderby-driven, never $search, and lands in the
 *      tools/list for the mail service.
 */
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();
const mockIsPathDenied = jest.fn<(t: string, u: string, type: string, path: string) => Promise<boolean>>();
const mockResolveMailFolderName = jest.fn<(g: unknown, id: string) => Promise<string | null>>();

interface Call { path: string; search?: string; filter?: string; orderby?: string; top?: number }
const calls: Call[] = [];
let respond: (call: Call, index: number) => Record<string, unknown> | Error = () => ({ value: [] });

function makeChain(path: string) {
  const call: Call = { path };
  calls.push(call);
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'expand', 'header']) chain[m] = () => chain;
  chain.search = (v: string) => { call.search = v; return chain; };
  chain.filter = (v: string) => { call.filter = v; return chain; };
  chain.orderby = (v: string) => { call.orderby = v; return chain; };
  chain.top = (v: number) => { call.top = v; return chain; };
  chain.get = () => {
    const r = respond(call, calls.length - 1);
    return r instanceof Error ? Promise.reject(r) : Promise.resolve(r);
  };
  chain.post = () => Promise.resolve({ id: 'x' });
  chain.patch = () => Promise.resolve({ id: 'x' });
  chain.delete = () => Promise.resolve(undefined);
  return chain;
}

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({ authenticateRequest: () => mockAuthenticateRequest() }));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => Promise.resolve('access-token'),
  getTenantIdFromSession: () => 'tenant-1',
}));
jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => ({ api: (path: string) => makeChain(path) }),
}));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: (...args: unknown[]) =>
    mockIsPathDenied(args[0] as string, args[1] as string, args[2] as string, args[3] as string),
}));
jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: (g: unknown, id: string) => mockResolveMailFolderName(g, id),
  resolveDefaultCalendarId: () => Promise.resolve(null),
  resolveCalendarName: () => Promise.resolve(null),
  resolveContactParentFolder: () => Promise.resolve(null),
  resolveDefaultContactFolder: () => Promise.resolve(null),
  resolveSectionNotebook: () => Promise.resolve(null),
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
}));
jest.mock('../services/sharepointFilter.js', () => ({ filterAndDisambiguateSites: (sites: unknown) => sites }));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => Promise.resolve(['mail']),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => Promise.resolve([]),
}));
jest.mock('../services/userServiceOverrides.js', () => ({ getUserServiceOverrides: () => Promise.resolve([]) }));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'draft' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));
jest.mock('../services/userMailConfig.js', () => ({ isMailIndexingDisabled: () => Promise.resolve(false) }));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint');
if (!registration) throw new Error('mcpEndpoint handler was not registered');
const handler = registration[1].handler;

const AUTH = { userId: 'u1', session: { userId: 'u1', tenantId: 'tenant-1', accessToken: 'fake', sessionToken: 'sess' } };

function rpc(method: string, params: Record<string, unknown>) {
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

const SENT = 'AAMkSentFolder001';
function msg(id: string, over: Record<string, unknown> = {}) {
  return {
    id,
    subject: `Subject ${id}`,
    from: { emailAddress: { name: 'Nate', address: 'nate@example.com' } },
    toRecipients: [{ emailAddress: { name: 'Jordan', address: 'jdoe@fabrikam.com' } }],
    receivedDateTime: '2026-09-23T17:05:12Z',
    bodyPreview: 'p',
    parentFolderId: SENT,
    ...over,
  };
}

beforeEach(() => {
  jest.clearAllMocks();
  calls.length = 0;
  respond = () => ({ value: [] });
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockIsPathDenied.mockResolvedValue(false);
  mockResolveMailFolderName.mockResolvedValue(null);
});

// ── Schema ────────────────────────────────────────────────────────────────────

describe('schema', () => {
  const SEARCH_PROPS = ['q', 'participant', 'from', 'to', 'since', 'folderId', 'mailboxId', 'maxResults'];
  const LIST_PROPS = ['folderId', 'since', 'mailboxId', 'maxResults'];

  it('tools/list exposes both tools with the parameters and q optional', async () => {
    const res = await rpc('tools/list', {});
    const tools = (res.jsonBody as any).result.tools as Array<{ name: string; inputSchema: any; description: string }>;
    const search = tools.find((t) => t.name === 'search_mail');
    const list = tools.find((t) => t.name === 'list_messages');
    expect(Object.keys(search!.inputSchema.properties).sort()).toEqual([...SEARCH_PROPS].sort());
    expect(search!.inputSchema.required).toEqual([]);
    expect(search!.description).toMatch(/relevance-ranked/);
    expect(search!.description).toMatch(/list_messages/);
    expect(Object.keys(list!.inputSchema.properties).sort()).toEqual([...LIST_PROPS].sort());
  });

  it('the manifest catalog agrees with tools/list on both tools', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = (getManifest() as any).tools as Array<{ name: string; inputSchema: any }>;
    const res = await rpc('tools/list', {});
    const live = (res.jsonBody as any).result.tools as Array<{ name: string; inputSchema: any }>;
    const shape = (s: any) => ({
      required: [...(s.required ?? [])].sort(),
      props: Object.fromEntries(Object.entries(s.properties).map(([k, v]) => [k, (v as { type: string }).type])),
    });
    for (const name of ['search_mail', 'list_messages']) {
      const m = manifest.find((t) => t.name === name)!;
      const l = live.find((t) => t.name === name)!;
      expect(m).toBeDefined();
      expect(l).toBeDefined();
      expect(shape(m.inputSchema)).toEqual(shape(l.inputSchema));
    }
  });

  it('the unsupported-arg check accepts the new criteria and still rejects an undeclared one', async () => {
    respond = () => ({ value: [msg('m1')] });
    const ok = toolResult(await callTool('search_mail', { to: 'fabrikam.com', since: '2026-09-15', folderId: SENT }));
    expect(ok.isError).toBe(false);
    const bad = toolResult(await callTool('search_mail', { q: 'x', recipients: 'y' }));
    expect(bad.isError).toBe(true);
    expect(bad.text).toMatch(/Unsupported parameter\(s\) for search_mail: recipients/);
  });

  it('search_mail with no criteria is an error that points at list_messages, before any Graph call', async () => {
    const r = toolResult(await callTool('search_mail', { folderId: SENT }));
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/list_messages/);
    expect(calls.filter((c) => c.path.includes('/messages'))).toHaveLength(0);
  });
});

// ── search_mail ───────────────────────────────────────────────────────────────

describe('search_mail', () => {
  it('folder + to: scans newest-first, matches the recipient, and reports the scan metadata', async () => {
    respond = () => ({ value: [msg('miss', { toRecipients: [{ emailAddress: { address: 'x@y.com' } }] }), msg('hit')] });
    const { parsed, isError } = toolResult(await callTool('search_mail', { folderId: SENT, to: 'jdoe@fabrikam.com' }));
    expect(isError).toBe(false);
    const graph = calls.find((c) => c.path === `/me/mailFolders/${SENT}/messages`)!;
    expect(graph.search).toBeUndefined();
    expect(graph.orderby).toBe('receivedDateTime desc');
    expect(parsed.results.map((r: { id: string }) => r.id)).toEqual(['hit']);
    expect(parsed.results[0].to).toEqual([{ name: 'Jordan', address: 'jdoe@fabrikam.com' }]);
    expect(parsed).toMatchObject({ count: 1, limit: 25, truncated: false, strategy: 'scan', ordering: 'newest-first', scanned: 2, scanComplete: true });
    expect(parsed.searchedFields).toContain('toRecipients');
  });

  it('folder + bare text: the $filter, with the not-searched note', async () => {
    respond = () => ({ value: [] });
    const { parsed } = toolResult(await callTool('search_mail', { folderId: SENT, q: 'fabrikam' }));
    const graph = calls.find((c) => c.path === `/me/mailFolders/${SENT}/messages`)!;
    expect(graph.search).toBeUndefined();
    expect(graph.filter).toContain("contains(subject,'fabrikam')");
    expect(parsed).toMatchObject({ count: 0, strategy: 'filter', searchedFields: ['subject', 'from.address', 'from.name'] });
    expect(parsed.notes[0]).toMatch(/Recipient addresses/);
  });

  it('mailbox-wide: KQL with property restrictions, relevance note, truncation off nextLink', async () => {
    respond = () => ({ value: [msg('m1')], '@odata.nextLink': 'https://graph/next' });
    const { parsed } = toolResult(await callTool('search_mail', { q: 'roadmap', participant: 'fabrikam.com', maxResults: 5 }));
    const graph = calls.find((c) => c.path === '/me/messages')!;
    expect(graph.search).toBe('"roadmap" AND participants:"fabrikam.com"');
    expect(graph.top).toBe(5);
    expect(parsed).toMatchObject({ count: 1, limit: 5, truncated: true, strategy: 'kql-search', ordering: 'relevance' });
    expect(parsed.notes[0]).toMatch(/relevance-ranked/);
  });

  it('a scan that stops early reports truncated:true and scanComplete:false', async () => {
    respond = (_c, i) => ({ value: [msg(`m${i}`)], '@odata.nextLink': `https://graph/next/${i}` });
    const { parsed } = toolResult(await callTool('search_mail', { folderId: SENT, to: 'fabrikam', maxResults: 1 }));
    expect(parsed.results).toHaveLength(1);
    expect(parsed).toMatchObject({ truncated: true, scanComplete: false });
  });

  it('a denied folder is blocked before any Graph query', async () => {
    mockResolveMailFolderName.mockResolvedValue('Legal Hold');
    mockIsPathDenied.mockResolvedValue(true);
    const r = toolResult(await callTool('search_mail', { folderId: SENT, to: 'x' }));
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/deny list/);
    expect(calls.filter((c) => c.path.includes('/messages'))).toHaveLength(0);
  });

  it('a mailbox-wide result set is post-filtered per message against the deny list', async () => {
    respond = () => ({ value: [msg('keep', { parentFolderId: 'f-inbox' }), msg('drop', { parentFolderId: 'f-hr' })] });
    mockResolveMailFolderName.mockImplementation((_g, id) => Promise.resolve(id === 'f-hr' ? 'HR' : 'Inbox'));
    mockIsPathDenied.mockImplementation((_t, _u, _s, path) => Promise.resolve(path === 'HR'));
    const { parsed } = toolResult(await callTool('search_mail', { participant: 'fabrikam' }));
    expect(parsed.results.map((r: { id: string }) => r.id)).toEqual(['keep']);
  });

  it('a Graph failure is an error, not an empty result', async () => {
    respond = () => new Error('MailboxNotEnabledForRESTAPI');
    const r = toolResult(await callTool('search_mail', { folderId: SENT, to: 'x' }));
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/Mail search failed/);
  });
});

// ── list_messages ─────────────────────────────────────────────────────────────

describe('list_messages', () => {
  it('is visible in tools/list under the mail service', async () => {
    const res = await rpc('tools/list', {});
    const names = ((res.jsonBody as any).result.tools as Array<{ name: string }>).map((t) => t.name);
    expect(names).toContain('list_messages');
  });

  it('enumerates newest-first with $orderby and never $search, with the envelope and metadata', async () => {
    respond = () => ({ value: [msg('a'), msg('b')], '@odata.nextLink': 'https://graph/more' });
    const { parsed, isError } = toolResult(await callTool('list_messages', { folderId: SENT, since: '2026-09-15', maxResults: 2 }));
    expect(isError).toBe(false);
    const graph = calls.find((c) => c.path === `/me/mailFolders/${SENT}/messages`)!;
    expect(graph.search).toBeUndefined();
    expect(graph.orderby).toBe('receivedDateTime desc');
    expect(graph.filter).toBe('receivedDateTime ge 2026-09-15T00:00:00.000Z');
    expect(graph.top).toBe(2);
    expect(parsed.items.map((r: { id: string }) => r.id)).toEqual(['a', 'b']);
    expect(parsed).toMatchObject({ count: 2, limit: 2, truncated: true, strategy: 'list', ordering: 'newest-first' });
  });

  it('a malformed since is an error before any Graph call', async () => {
    const r = toolResult(await callTool('list_messages', { since: 'last tuesday' }));
    expect(r.isError).toBe(true);
    expect(r.text).toMatch(/ISO-8601/);
    expect(calls.filter((c) => c.path.includes('/messages'))).toHaveLength(0);
  });

  it('a denied folder is blocked, and a mailbox-wide list is post-filtered', async () => {
    mockResolveMailFolderName.mockResolvedValue('HR');
    mockIsPathDenied.mockResolvedValue(true);
    const blocked = toolResult(await callTool('list_messages', { folderId: SENT }));
    expect(blocked.isError).toBe(true);

    calls.length = 0;
    respond = () => ({ value: [msg('keep', { parentFolderId: 'f-inbox' }), msg('drop', { parentFolderId: 'f-hr' })] });
    mockResolveMailFolderName.mockImplementation((_g, id) => Promise.resolve(id === 'f-hr' ? 'HR' : 'Inbox'));
    mockIsPathDenied.mockImplementation((_t, _u, _s, path) => Promise.resolve(path === 'HR'));
    const { parsed } = toolResult(await callTool('list_messages', {}));
    expect(parsed.items.map((r: { id: string }) => r.id)).toEqual(['keep']);
  });
});
