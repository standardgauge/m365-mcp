/**
 * OneNote page MCP tools: list, read, search and update.
 *
 * Driven through the real /api/mcp JSON-RPC dispatch so service gating,
 * read-only mode, opaque-ID validation and the deny list are exercised the way
 * a client reaches them. Graph is a route table: each test sets what a path
 * returns and reads back the calls made, with the query options each one set.
 */
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── Graph route table ────────────────────────────────────────────────────────

interface GraphCall { verb: string; path: string; body?: unknown; options: Record<string, unknown> }
const graphCalls: GraphCall[] = [];
let routes: Record<string, unknown> = {};

function routeResult(verb: string, path: string): Promise<unknown> {
  const key = `${verb} ${path}`;
  if (!(key in routes)) return Promise.reject(Object.assign(new Error(`unrouted ${key}`), { statusCode: 404 }));
  const value = routes[key];
  return value instanceof Error ? Promise.reject(value) : Promise.resolve(value);
}

function makeChain(path: string) {
  const options: Record<string, unknown> = {};
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'top', 'orderby', 'filter', 'expand', 'skip', 'header']) {
    chain[m] = (v: unknown) => {
      options[m] = v;
      return chain;
    };
  }
  for (const verb of ['get', 'post', 'patch', 'delete']) {
    chain[verb] = (body?: unknown) => {
      graphCalls.push({ verb, path, body, options: { ...options } });
      return routeResult(verb, path);
    };
  }
  chain.getStream = () => {
    graphCalls.push({ verb: 'getStream', path, options: { ...options } });
    return routeResult('getStream', path).then((html) => (async function* () {
      yield Buffer.from(String(html));
    })());
  };
  return chain;
}

// ── Deny list: an entry matches a notebook or section by ID or by name ──────

let denied: string[] = [];
const matches = (v: unknown) => typeof v === 'string' && denied.includes(v);

const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetReadOnlyServices = jest.fn<() => Promise<string[]>>();

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => Promise.resolve({ userId: 'u1', session: { userId: 'u1', tenantId: 't1', accessToken: 'x', sessionToken: 's' } }),
}));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => Promise.resolve('access-token'),
  getTenantIdFromSession: () => 't1',
}));
jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => ({ api: (path: string) => makeChain(path) }),
}));
jest.mock('../services/denyList.js', () => ({
  isPathDenied: (_t: unknown, _u: unknown, type: string, path: string) => Promise.resolve(type === 'onenote' && matches(path)),
  filterDeniedSearchHits: (_t: unknown, _u: unknown, items: unknown) => Promise.resolve(items),
  filterDeniedPaths: (_t: unknown, _u: unknown, type: string, items: Array<{ id?: string; name?: string }>) =>
    Promise.resolve(type === 'onenote' ? items.filter((i) => !matches(i.id) && !matches(i.name)) : items),
}));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => mockGetEnabledServices(),
  getReadOnlyServices: () => mockGetReadOnlyServices(),
  getAllowedSites: () => Promise.resolve([]),
}));
jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: () => Promise.resolve([]),
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

import { app } from '@azure/functions';
import { logAccess } from '../services/auditLog.js';
import '../functions/mcp/mcpEndpoint.js';
import { onenoteHtmlToText, odataString } from '../services/onenote.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint');
if (!registration) throw new Error('mcpEndpoint handler was not registered');
const handler = registration[1].handler;

function rpc(method: string, params: Record<string, unknown>) {
  const payload = { jsonrpc: '2.0', id: 1, method, params };
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve(payload),
    text: () => Promise.resolve(JSON.stringify(payload)),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

async function call(name: string, args: Record<string, unknown>): Promise<{ parsed: any; isError: boolean; text: string }> {
  const res = await rpc('tools/call', { name, arguments: args });
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '';
  const isError = body.result?.isError ?? false;
  return { parsed: isError ? null : JSON.parse(text), isError, text };
}

async function listedTools(): Promise<string[]> {
  const res = await rpc('tools/list', {});
  return (res.jsonBody as { result: { tools: Array<{ name: string }> } }).result.tools.map((t) => t.name);
}

const contentCalls = () => graphCalls.filter((c) => c.path.includes('/content'));
const writes = () => graphCalls.filter((c) => c.verb !== 'get' && c.verb !== 'getStream');

const NOTEBOOK = { id: 'nb-1', displayName: 'Clients' };
const SECTION = { id: 'sec-1', displayName: 'Meeting notes', parentNotebook: NOTEBOOK };
const PAGE = {
  id: 'pg-1',
  title: 'Kickoff',
  createdDateTime: '2026-11-01T10:00:00Z',
  lastModifiedDateTime: '2026-11-02T10:00:00Z',
  parentSection: { id: 'sec-1', displayName: 'Meeting notes' },
  parentNotebook: NOTEBOOK,
};
const OTHER_PAGE = {
  id: 'pg-2',
  title: 'Kickoff (board)',
  lastModifiedDateTime: '2026-11-03T10:00:00Z',
  parentSection: { id: 'sec-9', displayName: 'Board' },
  parentNotebook: { id: 'nb-9', displayName: 'Board papers' },
};
const PAGE_HTML =
  '<!DOCTYPE html><html><head><title>Kickoff</title><meta name="created" content="2026-11-01" /></head>' +
  '<body data-absolute-enabled="true"><div id="div:{33f8a242-7c33-4bb2-90c5-8425a68cc5bf}{40}">' +
  '<p>Agenda &amp; goals</p><ul><li>Scope</li><li>Budget &lt;$5k&gt;</li></ul></div></body></html>';

const NEXT = 'https://graph.microsoft.com/v1.0/next';

const PAGE_TOOLS = ['list_onenote_pages', 'read_onenote_page', 'search_onenote_pages', 'update_onenote_page'];
const UPDATE = { pageId: 'pg-1', target: 'body', action: 'append', content: '<p>Next steps</p>' };

beforeEach(() => {
  graphCalls.length = 0;
  (logAccess as jest.Mock).mockClear();
  denied = [];
  mockGetEnabledServices.mockResolvedValue(['mail', 'onenote']);
  mockGetReadOnlyServices.mockResolvedValue([]);
  routes = {
    'get /me/onenote/sections/sec-1': SECTION,
    'get /me/onenote/sections/sec-1/pages': { value: [PAGE] },
    'get /me/onenote/pages/pg-1': PAGE,
    'getStream /me/onenote/pages/pg-1/content?includeIDs=true': PAGE_HTML,
    'getStream /me/onenote/pages/pg-1/content': PAGE_HTML,
    'patch /me/onenote/pages/pg-1/content': undefined,
    'get /me/onenote/pages': { value: [PAGE, OTHER_PAGE] },
  };
});

describe('tools/list', () => {
  it('lists the page tools when onenote is enabled', async () => {
    const names = await listedTools();
    for (const t of PAGE_TOOLS) expect(names).toContain(t);
  });

  it('hides them when onenote is not enabled', async () => {
    mockGetEnabledServices.mockResolvedValue(['mail']);
    const names = await listedTools();
    for (const t of PAGE_TOOLS) expect(names).not.toContain(t);
  });

  it('hides update_onenote_page, and keeps the reads, when onenote is read-only', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['onenote']);
    const names = await listedTools();
    expect(names).not.toContain('update_onenote_page');
    for (const t of ['list_onenote_pages', 'read_onenote_page', 'search_onenote_pages']) expect(names).toContain(t);
  });

  it('advertises the same input-schema shape in getManifest() and tools/list', async () => {
    const { getManifest } = await import('../mcp/toolManifest.js');
    const manifest = getManifest() as { tools: Array<{ name: string; inputSchema: any }> };
    const res = await rpc('tools/list', {});
    const live = (res.jsonBody as { result: { tools: Array<{ name: string; inputSchema: any }> } }).result.tools;
    const shape = (schema: any) => ({
      required: [...(schema.required ?? [])].sort(),
      props: Object.fromEntries(
        Object.entries(schema.properties ?? {}).map(([k, v]: [string, any]) => [k, { type: v.type, enum: v.enum }]),
      ),
    });
    for (const name of PAGE_TOOLS) {
      const m = manifest.tools.find((t) => t.name === name);
      const l = live.find((t) => t.name === name);
      expect(m).toBeDefined();
      expect(l).toBeDefined();
      expect(shape(m!.inputSchema)).toEqual(shape(l!.inputSchema));
    }
  });
});

describe('list_onenote_pages', () => {
  it('lists a section\'s pages with title, id and last-modified, newest first', async () => {
    const { parsed, isError } = await call('list_onenote_pages', { sectionId: 'sec-1' });
    expect(isError).toBe(false);
    expect(parsed.items).toEqual([expect.objectContaining({
      id: 'pg-1', title: 'Kickoff', lastModifiedDateTime: '2026-11-02T10:00:00Z', sectionId: 'sec-1', notebookId: 'nb-1',
    })]);
    expect(parsed).toMatchObject({ count: 1, limit: 100, truncated: false });
    const list = graphCalls.find((c) => c.path === '/me/onenote/sections/sec-1/pages')!;
    expect(list.options).toMatchObject({ orderby: 'lastModifiedDateTime desc', top: 100 });
    expect(list.options.skip).toBeUndefined();
  });

  it('passes skip through and reports truncation from a nextLink', async () => {
    routes['get /me/onenote/sections/sec-1/pages'] = { value: [PAGE], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next' };
    const { parsed } = await call('list_onenote_pages', { sectionId: 'sec-1', skip: 100 });
    expect(parsed.truncated).toBe(true);
    expect(parsed.nextSkip).toBe(101);
    expect(graphCalls.find((c) => c.path === '/me/onenote/sections/sec-1/pages')!.options.skip).toBe(100);
  });

  it('omits nextSkip when there is nothing more to read', async () => {
    const { parsed } = await call('list_onenote_pages', { sectionId: 'sec-1' });
    expect(parsed.truncated).toBe(false);
    expect(parsed).not.toHaveProperty('nextSkip');
  });

  it('advances nextSkip by the rows Graph returned when some were denied', async () => {
    denied = ['nb-9'];
    routes['get /me/onenote/sections/sec-1/pages'] = { value: [PAGE, OTHER_PAGE, OTHER_PAGE], '@odata.nextLink': NEXT };
    const { parsed } = await call('list_onenote_pages', { sectionId: 'sec-1', skip: 10 });
    expect(parsed).toMatchObject({ count: 1, truncated: true, nextSkip: 13 });
  });

  it('moves past a page where every row was denied instead of repeating it', async () => {
    denied = ['nb-9'];
    routes['get /me/onenote/sections/sec-1/pages'] = { value: [OTHER_PAGE, OTHER_PAGE], '@odata.nextLink': NEXT };
    const { parsed } = await call('list_onenote_pages', { sectionId: 'sec-1' });
    expect(parsed).toMatchObject({ items: [], count: 0, truncated: true, nextSkip: 2 });
  });

  it('refuses a deny-listed section by ID before calling Graph', async () => {
    denied = ['sec-1'];
    const { isError, text } = await call('list_onenote_pages', { sectionId: 'sec-1' });
    expect(isError).toBe(true);
    expect(text).toMatch(/deny list/);
    expect(graphCalls).toEqual([]);
  });

  it('refuses a section whose notebook is deny-listed by name', async () => {
    denied = ['Clients'];
    const { isError, text } = await call('list_onenote_pages', { sectionId: 'sec-1' });
    expect(isError).toBe(true);
    expect(text).toMatch(/deny list/);
    expect(graphCalls.map((c) => c.path)).not.toContain('/me/onenote/sections/sec-1/pages');
  });

  it('rejects a negative skip', async () => {
    const { isError, text } = await call('list_onenote_pages', { sectionId: 'sec-1', skip: -1 });
    expect(isError).toBe(true);
    expect(text).toMatch(/skip/);
  });
});

describe('read_onenote_page', () => {
  it('returns the HTML with element ids by default', async () => {
    const { parsed, isError } = await call('read_onenote_page', { pageId: 'pg-1' });
    expect(isError).toBe(false);
    expect(parsed).toMatchObject({ id: 'pg-1', title: 'Kickoff', sectionId: 'sec-1', notebookId: 'nb-1', format: 'html' });
    expect(parsed.content).toBe(PAGE_HTML);
    expect(contentCalls().map((c) => c.path)).toEqual(['/me/onenote/pages/pg-1/content?includeIDs=true']);
  });

  it('returns plain text with format text', async () => {
    const { parsed } = await call('read_onenote_page', { pageId: 'pg-1', format: 'text' });
    expect(parsed.format).toBe('text');
    expect(parsed.content).toBe('Kickoff\n\nAgenda & goals\n- Scope\n- Budget <$5k>');
  });

  it('refuses a page in a deny-listed section and reads no content', async () => {
    denied = ['sec-1'];
    const { isError, text } = await call('read_onenote_page', { pageId: 'pg-1' });
    expect(isError).toBe(true);
    expect(text).toMatch(/deny list/);
    expect(contentCalls()).toEqual([]);
    expect(logAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'read_onenote_page', result: 'denied', resource: 'pg-1' }));
  });

  it('refuses a page in a section deny-listed by name', async () => {
    denied = ['Meeting notes'];
    const { isError } = await call('read_onenote_page', { pageId: 'pg-1', format: 'text' });
    expect(isError).toBe(true);
    expect(contentCalls()).toEqual([]);
  });

  it('refuses a page whose notebook is deny-listed', async () => {
    denied = ['nb-1'];
    const { isError } = await call('read_onenote_page', { pageId: 'pg-1' });
    expect(isError).toBe(true);
    expect(contentCalls()).toEqual([]);
  });

  it('refuses a page whose parents Graph did not return (fail closed)', async () => {
    routes['get /me/onenote/pages/pg-1'] = { id: 'pg-1', title: 'Kickoff' };
    const { isError } = await call('read_onenote_page', { pageId: 'pg-1' });
    expect(isError).toBe(true);
    expect(contentCalls()).toEqual([]);
  });

  it('rejects a pageId that would alter the Graph path', async () => {
    const { isError } = await call('read_onenote_page', { pageId: '../../users/victim' });
    expect(isError).toBe(true);
    expect(graphCalls).toEqual([]);
  });
});

describe('search_onenote_pages', () => {
  it('filters titles case-insensitively and drops pages in deny-listed notebooks', async () => {
    denied = ['Board papers'];
    const { parsed, isError } = await call('search_onenote_pages', { query: "Kick'Off" });
    expect(isError).toBe(false);
    expect(parsed.items.map((p: { id: string }) => p.id)).toEqual(['pg-1']);
    const search = graphCalls.find((c) => c.path === '/me/onenote/pages')!;
    expect(search.options.filter).toBe("contains(tolower(title),'kick''off')");
  });

  it('advances nextSkip by the rows Graph returned when some were denied', async () => {
    denied = ['Board papers'];
    routes['get /me/onenote/pages'] = { value: [OTHER_PAGE, PAGE, OTHER_PAGE], '@odata.nextLink': NEXT };
    const { parsed } = await call('search_onenote_pages', { query: 'kickoff', skip: 100 });
    expect(parsed).toMatchObject({ count: 1, truncated: true, nextSkip: 103 });
    expect(graphCalls.find((c) => c.path === '/me/onenote/pages')!.options.skip).toBe(100);
  });

  it('moves past a page where every match was denied instead of repeating it', async () => {
    denied = ['Board papers'];
    routes['get /me/onenote/pages'] = { value: [OTHER_PAGE], '@odata.nextLink': NEXT };
    const { parsed } = await call('search_onenote_pages', { query: 'kickoff' });
    expect(parsed).toMatchObject({ items: [], count: 0, truncated: true, nextSkip: 1 });
  });

  it('searches one section when sectionId is given, after checking it', async () => {
    await call('search_onenote_pages', { query: 'kickoff', sectionId: 'sec-1' });
    expect(graphCalls.map((c) => c.path)).toEqual(['/me/onenote/sections/sec-1', '/me/onenote/sections/sec-1/pages']);
  });

  it('refuses a deny-listed sectionId', async () => {
    denied = ['sec-1'];
    const { isError } = await call('search_onenote_pages', { query: 'kickoff', sectionId: 'sec-1' });
    expect(isError).toBe(true);
    expect(graphCalls).toEqual([]);
  });

  it('requires a query', async () => {
    const { isError } = await call('search_onenote_pages', { query: '   ' });
    expect(isError).toBe(true);
    expect(graphCalls).toEqual([]);
  });
});

describe('update_onenote_page', () => {
  it('PATCHes one command to the page content', async () => {
    const { parsed, isError } = await call('update_onenote_page', { ...UPDATE, target: '#notes', action: 'insert', position: 'before' });
    expect(isError).toBe(false);
    expect(parsed).toEqual({ status: 'updated', pageId: 'pg-1' });
    expect(writes()).toEqual([expect.objectContaining({
      verb: 'patch',
      path: '/me/onenote/pages/pg-1/content',
      body: [{ target: '#notes', action: 'insert', content: '<p>Next steps</p>', position: 'before' }],
    })]);
  });

  it('accepts a generated element id from read_onenote_page', async () => {
    const { isError } = await call('update_onenote_page', { ...UPDATE, target: 'div:{33f8a242-7c33-4bb2-90c5-8425a68cc5bf}{40}', action: 'replace' });
    expect(isError).toBe(false);
  });

  it('refuses a page in a deny-listed section and writes nothing', async () => {
    denied = ['sec-1'];
    const { isError, text } = await call('update_onenote_page', UPDATE);
    expect(isError).toBe(true);
    expect(text).toMatch(/deny list/);
    expect(writes()).toEqual([]);
    expect(logAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'update_onenote_page', result: 'denied', resource: 'pg-1' }));
  });

  it('refuses a page whose notebook is deny-listed by name and writes nothing', async () => {
    denied = ['Clients'];
    const { isError } = await call('update_onenote_page', UPDATE);
    expect(isError).toBe(true);
    expect(writes()).toEqual([]);
  });

  it('is refused when onenote is read-only', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['onenote']);
    const { isError, text } = await call('update_onenote_page', UPDATE);
    expect(isError).toBe(true);
    expect(text).toMatch(/read-only/);
    expect(graphCalls).toEqual([]);
  });

  it.each([
    [{ target: 'body; drop' }, /target/],
    [{ target: '' }, /target/],
    [{ action: 'delete' }, /action/],
    [{ position: 'middle' }, /position/],
    [{ content: 42 }, /content/],
  ])('rejects %j before any Graph call', async (bad, msg) => {
    const { isError, text } = await call('update_onenote_page', { ...UPDATE, ...bad });
    expect(isError).toBe(true);
    expect(text).toMatch(msg);
    expect(graphCalls).toEqual([]);
  });
});

describe('helpers', () => {
  it('odataString doubles single quotes', () => {
    expect(odataString("o'brien")).toBe("'o''brien'");
  });

  it('onenoteHtmlToText drops scripts, styles and malformed tags', () => {
    expect(onenoteHtmlToText('<html><head><style>p{}</style></head><body><p>a<script>x()</script>b</p><p>c &#x263A; &#9731;</p><p>d<br/>e</p><b')).toBe(
      'ab\nc ☺ ☃\nd\ne',
    );
  });
});
