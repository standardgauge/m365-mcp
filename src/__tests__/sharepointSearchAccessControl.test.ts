/**
 * End-to-end access-control tests for SharePoint search.
 *
 * search_sharepoint was disclosing file names, paths, sizes, and timestamps from
 * sites NOT in the allow-list (other sites + personal OneDrive) and from folders
 * hidden by the deny list, even though list_folders / read_file correctly reject
 * them. These tests drive BOTH search surfaces — the MCP `search_sharepoint`
 * tool and the HTTP `/api/sharepoint/search` route — through the FULL chain with
 * the REAL deny-list service and REAL policy wrapper; only the external
 * boundaries (Microsoft Graph, Azure Table Storage) are mocked.
 *
 *   request → policy/dispatch (real) → search handler (real)
 *           → Graph /search/query (mocked responder, records the request body)
 *           → denyList.filterDeniedSearchHits (REAL) → Table Storage (in-memory)
 *
 * Acceptance criteria proven together:
 *   1. Results never include a site outside the allow-list (incl. `-my` OneDrive).
 *   2. Items under a denied folder are stripped — deny entry as a path form OR a
 *      bare folder name.
 *   3. The query is constrained at the source (contentSources) and requests
 *      parentReference; the MCP `siteId` arg is honored.
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

// ── Site fixtures ────────────────────────────────────────────────────────────

const TENANT = 'example-tenant';
const USER = 'user-e2e';
const HOST = 'example.sharepoint.com';
const MYHOST = 'exampleequitypartners-my.sharepoint.com';

// Allow-list = root site + it-hub (mirrors the reported example config).
const ROOT_ID = `${HOST},11111111-1111-1111-1111-111111111111,aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa`;
const ITHUB_ID = `${HOST},22222222-2222-2222-2222-222222222222,bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb`;
const ARIC_ID = `${HOST},33333333-3333-3333-3333-333333333333,cccccccc-cccc-cccc-cccc-cccccccccccc`;
const ONEDRIVE_ID = `${MYHOST},44444444-4444-4444-4444-444444444444,dddddddd-dddd-dddd-dddd-dddddddddddd`;
const ALLOWED_SITES = [
  { id: ROOT_ID, name: 'Root' },
  { id: ITHUB_ID, name: 'IT Hub' },
];

let allowedSites: Array<{ id: string; name: string }> = [];

// ── Graph mock ───────────────────────────────────────────────────────────────

interface GraphCall { path: string; body?: unknown }
const graphCalls: GraphCall[] = [];
let searchHits: unknown[] = [];

const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => ({
    post: (body: unknown) => {
      graphCalls.push({ path, body });
      return Promise.resolve({ value: [{ hitsContainers: [{ hits: searchHits }] }] });
    },
  }),
}));

/** Build a driveItem search hit for a given site + webUrl. */
function hit(opts: { id: string; name: string; siteId: string; webUrl: string; size?: number }) {
  return {
    resource: {
      id: opts.id,
      name: opts.name,
      webUrl: opts.webUrl,
      size: opts.size ?? 1024,
      createdDateTime: '2026-01-01T00:00:00Z',
      lastModifiedDateTime: '2026-02-01T00:00:00Z',
      parentReference: { siteId: opts.siteId },
    },
  };
}

// ── Boundary mocks (everything except denyList) ──────────────────────────────

const SESSION = { userId: USER, tenantId: TENANT, email: 'e2e@example.test' };

jest.mock('../services/telemetry.js', () => ({}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () =>
    Promise.resolve({ userId: USER, session: SESSION } as unknown as AuthResult),
  checkGlobalAdmin: () => Promise.resolve(false),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => Promise.resolve(['sharepoint']),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => Promise.resolve(allowedSites),
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
  createGraphClient: (...args: unknown[]) =>
    (mockCreateGraphClient as (...a: unknown[]) => unknown)(...args),
}));

jest.mock('../services/sharepointFilter.js', () => ({
  filterAndDisambiguateSites: (sites: unknown) => sites,
}));

// assertOpaqueId(s) no-op so composite site IDs (which contain commas) pass.
jest.mock('../services/opaqueId.js', () => ({
  assertOpaqueId: jest.fn(),
  assertOpaqueIds: jest.fn(),
  ValidationError: class ValidationError extends Error {},
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

// ── Import handlers after mocks ──────────────────────────────────────────────

import { app } from '@azure/functions';
import '../functions/sharepoint/searchSharepoint.js';
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;

function handlerFor(name: string) {
  const reg = httpMock.mock.calls.find((call) => call[0] === name);
  if (!reg) throw new Error(`${name} handler was not registered`);
  return reg[1].handler;
}

const httpSearch = handlerFor('searchSharepoint');
const mcpHandler = handlerFor('mcpEndpoint');

const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;

// ── Request helpers ──────────────────────────────────────────────────────────

function httpRequest(query: Record<string, string>): HttpRequest {
  return {
    method: 'GET',
    query: { get: (k: string) => query[k] ?? null },
    params: {},
    headers: new Map<string, string>([['x-forwarded-for', '203.0.113.9']]),
    json: () => Promise.resolve({}),
  } as unknown as HttpRequest;
}

async function callHttpSearch(query: Record<string, string>) {
  const res = await httpSearch(httpRequest(query), ctx);
  return res.jsonBody as { results: Array<{ name: string; siteId: string | null }>; count: number };
}

async function callMcpSearch(args: Record<string, unknown>) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name: 'search_sharepoint', arguments: args } }),
  } as unknown as HttpRequest;
  const res = await mcpHandler(req, ctx);
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  const text = body.result?.content?.[0]?.text ?? '{}';
  // search_sharepoint returns a truncation-aware envelope { results, count, limit, truncated }.
  const parsed = JSON.parse(text) as { results?: Array<{ name: string; siteId?: string }>; truncated?: boolean; count?: number; limit?: number };
  return { items: parsed.results ?? [], truncated: parsed.truncated, count: parsed.count, isError: body.result?.isError ?? false };
}

function pushGlobalDeny(path: string) {
  globalEntries.push({ partitionKey: `${TENANT}:sharepoint`, rowKey: Buffer.from(path).toString('base64'), path });
}

function reset() {
  globalEntries.length = 0;
  userEntries.length = 0;
  graphCalls.length = 0;
  searchHits = [];
  allowedSites = [];
  jest.clearAllMocks();
  delete process.env.DEFAULT_SHAREPOINT_DENY_PATHS;
}

beforeEach(reset);

// A cross-boundary result set: allowed sites, a foreign site, and OneDrive.
function crossBoundaryHits() {
  return [
    hit({ id: 'f1', name: 'roadmap.docx', siteId: ROOT_ID, webUrl: `https://${HOST}/Shared%20Documents/General/roadmap.docx` }),
    hit({ id: 'f2', name: 'runbook.md', siteId: ITHUB_ID, webUrl: `https://${HOST}/sites/it-hub/Shared%20Documents/runbook.md` }),
    hit({ id: 'f3', name: 'ProjectAlpha.xlsx', siteId: ARIC_ID, webUrl: `https://${HOST}/sites/AricPublic/Shared%20Documents/ProjectAlpha.xlsx` }),
    hit({ id: 'f4', name: 'salary.xlsx', siteId: ONEDRIVE_ID, webUrl: `https://${MYHOST}/personal/ceo_example_test/Documents/salary.xlsx` }),
  ];
}

// ─────────────────────────────────────────────────────────────────────────────
// 1. Allow-list enforcement
// ─────────────────────────────────────────────────────────────────────────────

describe('E2E: search enforces the site allow-list', () => {
  it('MCP: drops hits from non-allow-listed sites and personal OneDrive', async () => {
    allowedSites = ALLOWED_SITES;
    searchHits = crossBoundaryHits();

    const { items, isError } = await callMcpSearch({ q: 'project' });

    expect(isError).toBe(false);
    expect(items.map((i) => i.name).sort()).toEqual(['roadmap.docx', 'runbook.md']);
    // AricPublic + OneDrive metadata never leaves the boundary.
    expect(JSON.stringify(items)).not.toContain('ProjectAlpha');
    expect(JSON.stringify(items)).not.toContain('salary');
  });

  it('HTTP: drops hits from non-allow-listed sites and personal OneDrive', async () => {
    allowedSites = ALLOWED_SITES;
    searchHits = crossBoundaryHits();

    const body = await callHttpSearch({ q: 'project' });

    expect(body.count).toBe(2);
    expect(body.results.map((r) => r.name).sort()).toEqual(['roadmap.docx', 'runbook.md']);
  });

  it('never sends contentSources (invalid for driveItem,) and requests parentReference', async () => {
    allowedSites = ALLOWED_SITES;
    searchHits = crossBoundaryHits();

    await callMcpSearch({ q: 'project' });

    const req = graphCalls.find((c) => c.path === '/search/query');
    const request = (req?.body as any).requests[0];
    // contentSources errored on every driveItem query — it must never be sent.
    expect(request.contentSources).toBeUndefined();
    expect(request.entityTypes).toEqual(['driveItem']);
    expect(request.fields).toContain('parentReference');
    // With a scope active the request over-fetches beyond maxResults so the
    // post-filter can still fill the page.
    expect(request.size).toBeGreaterThanOrEqual(50);
  });

  it('empty allow-list requests exactly maxResults (no over-fetch) — nothing dropped by site', async () => {
    allowedSites = [];
    searchHits = crossBoundaryHits();

    const { items } = await callMcpSearch({ q: 'project', maxResults: 25 });

    const req = graphCalls.find((c) => c.path === '/search/query');
    const request = (req?.body as any).requests[0];
    expect(request.contentSources).toBeUndefined();
    expect(request.size).toBe(25);
    expect(items).toHaveLength(4); // no allow-list → site filter is a no-op
  });

  it('MCP: honors the siteId arg by post-filtering to that single site', async () => {
    allowedSites = ALLOWED_SITES;
    // Two allowed-site hits; siteId should narrow the result to just it-hub.
    searchHits = [
      hit({ id: 'f1', name: 'roadmap.docx', siteId: ROOT_ID, webUrl: `https://${HOST}/Shared%20Documents/General/roadmap.docx` }),
      hit({ id: 'f2', name: 'runbook.md', siteId: ITHUB_ID, webUrl: `https://${HOST}/sites/it-hub/Shared%20Documents/runbook.md` }),
    ];

    const { items } = await callMcpSearch({ q: 'runbook', siteId: ITHUB_ID });

    const req = graphCalls.find((c) => c.path === '/search/query');
    // No source scoping; siteId narrows via the fail-closed post-filter.
    expect((req?.body as any).requests[0].contentSources).toBeUndefined();
    expect(items.map((i) => i.name)).toEqual(['runbook.md']);
  });
});

// ─────────────────────────────────────────────────────────────────────────────
// 2. Deny-list enforcement over search hits
// ─────────────────────────────────────────────────────────────────────────────

describe('E2E: search enforces the folder deny-list', () => {
  // All hits below are on the allow-listed root site, so only the deny list
  // can remove them — isolating the deny-list behavior.
  function financeAndIrHits() {
    return [
      hit({ id: 'g1', name: 'ok.docx', siteId: ROOT_ID, webUrl: `https://${HOST}/Shared%20Documents/General/ok.docx` }),
      hit({ id: 'g2', name: 'deal-memo.xlsx', siteId: ROOT_ID, webUrl: `https://${HOST}/Shared%20Documents/Finance/deal-memo.xlsx` }),
      hit({ id: 'g3', name: 'q3-8k.pdf', siteId: ROOT_ID, webUrl: `https://${HOST}/Shared%20Documents/IR/q3-8k.pdf` }),
    ];
  }

  it('MCP: strips items under a path-form denied folder (/Shared Documents/Finance)', async () => {
    allowedSites = ALLOWED_SITES;
    pushGlobalDeny('/Shared Documents/Finance');
    searchHits = financeAndIrHits();

    const { items } = await callMcpSearch({ q: 'q3' });

    expect(items.map((i) => i.name).sort()).toEqual(['ok.docx', 'q3-8k.pdf']);
    expect(JSON.stringify(items)).not.toContain('deal-memo');
  });

  it('MCP: strips items under a bare-folder-name denied entry (IR)', async () => {
    allowedSites = ALLOWED_SITES;
    pushGlobalDeny('IR');
    searchHits = financeAndIrHits();

    const { items } = await callMcpSearch({ q: 'q3' });

    expect(items.map((i) => i.name).sort()).toEqual(['deal-memo.xlsx', 'ok.docx']);
    expect(JSON.stringify(items)).not.toContain('q3-8k');
  });

  it('HTTP: enforces both path-form and bare-name deny entries together', async () => {
    allowedSites = ALLOWED_SITES;
    pushGlobalDeny('/Shared Documents/Finance');
    pushGlobalDeny('IR');
    searchHits = financeAndIrHits();

    const body = await callHttpSearch({ q: 'q3' });

    expect(body.results.map((r) => r.name)).toEqual(['ok.docx']);
    expect(body.count).toBe(1);
  });

  it('enforces the config-driven DEFAULT_SHAREPOINT_DENY_PATHS with an empty admin table', async () => {
    allowedSites = ALLOWED_SITES;
    process.env.DEFAULT_SHAREPOINT_DENY_PATHS = 'Finance,IR';
    searchHits = financeAndIrHits();

    const body = await callHttpSearch({ q: 'q3' });

    expect(body.results.map((r) => r.name)).toEqual(['ok.docx']);
  });

  it('does not over-block a folder whose name merely contains a denied token', async () => {
    allowedSites = ALLOWED_SITES;
    pushGlobalDeny('IR');
    searchHits = [
      hit({ id: 'h1', name: 'notes.docx', siteId: ROOT_ID, webUrl: `https://${HOST}/Shared%20Documents/Investor%20IR%20Notes/notes.docx` }),
    ];

    const body = await callHttpSearch({ q: 'notes' });

    // "Investor IR Notes" is a distinct segment from "IR" — kept.
    expect(body.results.map((r) => r.name)).toEqual(['notes.docx']);
  });
});
