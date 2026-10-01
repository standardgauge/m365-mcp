/**
 * End-to-end tests for SharePoint file discovery + list paging.
 *
 * These drive the MCP `list_folders` and `list_sharepoint_list_items` tools
 * through the full tools/call path with the REAL deny-list service (backed by
 * in-memory Table Storage); only Microsoft Graph is mocked.
 *
 * Regressions covered:
 *   - list_folders returned only folders, so a file's itemId (required by
 *     read_file) was unreachable. `includeFiles` now surfaces files, each with
 *     its `id`, and files are deny-filtered exactly like folders.
 *   - list_sharepoint_list_items had no paging, so an item deep in a large list
 *     could not be reached. `offset` now pages into the collection.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── In-memory Azure Table Storage (backs the REAL denyList service) ─────────────

const globalEntries: Array<{ partitionKey: string; rowKey: string; path: string }> = [];
const userEntries: Array<{ partitionKey: string; rowKey: string; path: string }> = [];

jest.mock('@azure/data-tables', () => {
  const makeAsyncIterable = (rows: unknown[]) => ({
    [Symbol.asyncIterator]: async function* () {
      for (const row of rows) yield row;
    },
  });
  return {
    TableClient: {
      fromConnectionString: jest.fn((_conn: string, tableName: string) => ({
        createTable: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
        upsertEntity: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
        deleteEntity: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
        listEntities: jest.fn(({ queryOptions }: { queryOptions: { filter: string } }) => {
          const source = tableName === 'GlobalDenyList' ? globalEntries : userEntries;
          const pkMatch = queryOptions.filter.match(/PartitionKey eq '([^']+)'/);
          const pk = pkMatch ? pkMatch[1] : '';
          return makeAsyncIterable(source.filter((e) => e.partitionKey === pk));
        }),
      })),
    },
    odata: (strings: TemplateStringsArray, ...values: unknown[]) =>
      strings.reduce((acc, str, i) => `${acc}${str}${i < values.length ? `'${values[i]}'` : ''}`, ''),
  };
});

process.env.AZURE_STORAGE_CONNECTION_STRING = 'UseDevelopmentStorage=true';

const TENANT = 'list-tenant';
const USER = 'user-list';
const SITE_ID = 'contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const LIST_ID = 'list-guid-1';

// ── Graph mock: a chainable request builder over configured datasets ───────────

interface RawItem { id: string; name: string; folder?: { childCount: number }; file?: { mimeType: string }; size?: number; lastModifiedDateTime?: string; parentReference: { path: string }; fields?: Record<string, unknown> }

// Drive children under the parent folder, and list items, for the tests to set.
let children: RawItem[] = [];
let listItems: RawItem[] = [];

function makeBuilder(path: string) {
  const state = { path, filter: '' };
  const builder: Record<string, unknown> = {};
  const chain = () => builder;
  builder.select = chain;
  builder.top = chain;
  builder.expand = chain;
  builder.header = chain;
  builder.filter = (f: string) => { state.filter = f; return builder; };
  builder.get = () => {
    // Drive children collection.
    if (state.path.endsWith('/children')) {
      const base = state.filter === 'folder ne null' ? children.filter((c) => c.folder) : children;
      return Promise.resolve({ value: base });
    }
    // List items collection.
    if (/\/lists\/[^/]+\/items$/.test(state.path)) {
      return Promise.resolve({ value: listItems });
    }
    return Promise.resolve({ value: [] });
  };
  return builder;
}

const mockCreateGraphClient = jest.fn(() => ({ api: (path: string) => makeBuilder(path) }));

// ── Boundary mocks (everything except denyList) ──────────────────────────────

const SESSION = { userId: USER, tenantId: TENANT, email: 'list@test' };

jest.mock('../services/telemetry.js', () => ({}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => Promise.resolve({ userId: USER, session: SESSION } as unknown as AuthResult),
  checkGlobalAdmin: () => Promise.resolve(false),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => Promise.resolve(['sharepoint']),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => Promise.resolve([]), // allow-all so the siteId gate passes
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  isServiceDisabledForUser: () => Promise.resolve(false),
  getUserServiceOverrides: () => Promise.resolve([]),
}));

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: () => Promise.resolve(false),
}));

jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'draft' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));

jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: () => Promise.resolve(TENANT),
  getTenantIdFromSession: (session: { tenantId?: string }) => session?.tenantId ?? TENANT,
  getValidAccessTokenForSession: () => Promise.resolve('fake-token'),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: (...args: unknown[]) => (mockCreateGraphClient as (...a: unknown[]) => unknown)(...args),
}));

jest.mock('../services/sharepointFilter.js', () => ({
  filterAndDisambiguateSites: (sites: unknown) => sites,
}));

jest.mock('../services/opaqueId.js', () => ({
  assertOpaqueId: jest.fn(),
  assertOpaqueIds: jest.fn(),
  ValidationError: class ValidationError extends Error {},
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const mcpHandler = httpMock.mock.calls.find((c) => c[0] === 'mcpEndpoint')![1].handler;

const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;

async function callTool(name: string, args: Record<string, unknown>) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
  } as unknown as HttpRequest;
  const res = await mcpHandler(req, ctx);
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? 'null';
  return { data: JSON.parse(text), isError: body.result?.isError ?? false };
}

function folder(name: string, childCount = 0): RawItem {
  return { id: `folder-${name}`, name, folder: { childCount }, parentReference: { path: '/drive/root:/IT/Atlas Technica' } };
}
function file(name: string, size = 2048): RawItem {
  return { id: `file-${name}`, name, file: { mimeType: 'application/pdf' }, size, lastModifiedDateTime: '2026-02-01T00:00:00Z', parentReference: { path: '/drive/root:/IT/Atlas Technica' } };
}

function pushGlobalDeny(path: string) {
  globalEntries.push({ partitionKey: `${TENANT}:sharepoint`, rowKey: Buffer.from(path).toString('base64'), path });
}

function reset() {
  globalEntries.length = 0;
  userEntries.length = 0;
  children = [];
  listItems = [];
  jest.clearAllMocks();
  delete process.env.DEFAULT_SHAREPOINT_DENY_PATHS;
}

beforeEach(reset);

// ─────────────────────────────────────────────────────────────────────────────
// list_folders — files + deny-list + paging
// ─────────────────────────────────────────────────────────────────────────────

describe('list_folders', () => {
  it('returns only folders by default (includeFiles omitted)', async () => {
    children = [folder('Reports'), file('contract.pdf'), file('addendum.pdf')];

    const { data, isError } = await callTool('list_folders', { siteId: SITE_ID, parentId: 'p1' });

    expect(isError).toBe(false);
    expect(data.items.map((i: { name: string }) => i.name)).toEqual(['Reports']);
    expect(data.items[0].type).toBe('folder');
    expect(data.count).toBe(1);
  });

  it('surfaces files with their itemId when includeFiles=true (the read_file path)', async () => {
    children = [folder('Reports'), file('contract.pdf'), file('addendum.pdf')];

    const { data } = await callTool('list_folders', { siteId: SITE_ID, parentId: 'p1', includeFiles: true });

    const byName = Object.fromEntries(data.items.map((i: { name: string }) => [i.name, i]));
    expect(Object.keys(byName).sort()).toEqual(['Reports', 'addendum.pdf', 'contract.pdf']);
    // A file carries the id + metadata read_file needs.
    expect(byName['contract.pdf']).toMatchObject({ type: 'file', id: 'file-contract.pdf', size: 2048, lastModifiedDateTime: '2026-02-01T00:00:00Z' });
    expect(byName['Reports']).toMatchObject({ type: 'folder' });
  });

  it('applies the deny-list to files under a denied folder path', async () => {
    // Path-form deny on the containing folder — its files must not leak (this is
    // the same filterDeniedPaths matching folders already get; files now share it).
    children = [file('public.pdf'), file('secret.pdf')];
    pushGlobalDeny('/IT/Atlas Technica');

    const { data } = await callTool('list_folders', { siteId: SITE_ID, parentId: 'p1', includeFiles: true });

    // Every child sits under the denied "/IT/Atlas Technica" folder → all removed.
    expect(data.items).toEqual([]);
    expect(data.count).toBe(0);
  });

  it('drops a file denied by its own path while keeping its siblings', async () => {
    children = [file('public.pdf'), file('secret.pdf')];
    pushGlobalDeny('/IT/Atlas Technica/secret.pdf');

    const { data } = await callTool('list_folders', { siteId: SITE_ID, parentId: 'p1', includeFiles: true });

    expect(data.items.map((i: { name: string }) => i.name)).toEqual(['public.pdf']);
  });

  it('pages children with offset and reports hasMore/nextOffset', async () => {
    children = Array.from({ length: 5 }, (_, i) => file(`f${i}.pdf`));

    const first = await callTool('list_folders', { siteId: SITE_ID, parentId: 'p1', includeFiles: true, maxResults: 2, offset: 0 });
    expect(first.data.items.map((i: { name: string }) => i.name)).toEqual(['f0.pdf', 'f1.pdf']);
    expect(first.data.hasMore).toBe(true);
    expect(first.data.nextOffset).toBe(2);

    const next = await callTool('list_folders', { siteId: SITE_ID, parentId: 'p1', includeFiles: true, maxResults: 2, offset: 4 });
    expect(next.data.items.map((i: { name: string }) => i.name)).toEqual(['f4.pdf']);
    expect(next.data.hasMore).toBe(false);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// list_sharepoint_list_items — paging
// ─────────────────────────────────────────────────────────────────────────────

describe('list_sharepoint_list_items', () => {
  it('pages into the list with offset and reports continuation metadata', async () => {
    listItems = Array.from({ length: 10 }, (_, i) => ({ id: `${i}`, name: `row-${i}`, parentReference: { path: '' }, fields: { Title: `row-${i}` } }));

    const page = await callTool('list_sharepoint_list_items', { siteId: SITE_ID, listId: LIST_ID, offset: 5, maxResults: 3 });

    expect(page.data.items.map((i: { id: string }) => i.id)).toEqual(['5', '6', '7']);
    expect(page.data.items[0].fields).toEqual({ Title: 'row-5' });
    expect(page.data.count).toBe(3);
    expect(page.data.hasMore).toBe(true);
    expect(page.data.nextOffset).toBe(8);
  });

  it('reports hasMore=false at the tail of the list', async () => {
    listItems = Array.from({ length: 4 }, (_, i) => ({ id: `${i}`, name: `row-${i}`, parentReference: { path: '' }, fields: {} }));

    const page = await callTool('list_sharepoint_list_items', { siteId: SITE_ID, listId: LIST_ID, offset: 2, maxResults: 10 });

    expect(page.data.items.map((i: { id: string }) => i.id)).toEqual(['2', '3']);
    expect(page.data.hasMore).toBe(false);
    expect(page.data.nextOffset).toBe(4);
  });
});
