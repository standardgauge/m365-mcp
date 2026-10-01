/**
 * End-to-end tests for the Outlook mail deny-list enforcement.
 *
 * Unlike the per-handler unit tests, these exercise the FULL request chain with
 * the REAL deny-list service and REAL policy-enforcement wrapper in the loop —
 * only the external I/O boundaries are mocked:
 *
 *   HTTP request
 *     → withSecurity            (real)
 *     → withPolicyEnforcement   (real — auth, service-enabled, deny-list gate)
 *     → mail handler            (real)
 *     → Microsoft Graph         (mocked — in-memory responder)
 *     → containerResolver       (mocked — folderId → display-name map)
 *     → denyList.isPathDenied / filterDeniedPaths  (REAL)
 *       → Azure Table Storage   (mocked — in-memory tables)
 *
 * The point is to prove two acceptance criteria together:
 *   1. Tool calls actually reach the Graph API (we assert the Graph mock was hit).
 *   2. Denied folders are stripped / blocked — including the config-driven
 *      DEFAULT_MAIL_DENY_FOLDERS defaults, enforced even when the admin deny
 *      table is empty (the Example "no access to Finance/HR/Legal/IR" guarantee).
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

// denyList reads this at import time — set before any handler (which imports it).
process.env.AZURE_STORAGE_CONNECTION_STRING = 'UseDevelopmentStorage=true';

// ── Graph client mock (records that tool calls reach the Graph API) ─────────────

interface GraphCall { path: string; op: 'get' | 'post'; body?: unknown }
const graphCalls: GraphCall[] = [];
const mockGraphGet = jest.fn<(path: string) => Promise<unknown>>();
const mockGraphPost = jest.fn<(path: string, body: unknown) => Promise<unknown>>();

const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    const builder = {
      select: () => builder,
      search: () => builder,
      filter: () => builder,
      top: () => builder,
      orderby: () => builder,
      get: () => {
        graphCalls.push({ path, op: 'get' });
        return mockGraphGet(path);
      },
      post: (body: unknown) => {
        graphCalls.push({ path, op: 'post', body });
        return mockGraphPost(path, body);
      },
    };
    return builder;
  },
}));

// folderId → display name, controlled per test.
const folderNames = new Map<string, string>();
const mockResolveMailFolderName = jest.fn(
  (_graph: unknown, folderId: string) => Promise.resolve(folderNames.get(folderId) ?? null),
);

// ── Boundary mocks (everything except denyList + policyEnforcement) ─────────────

const TENANT = 'example-tenant';
const USER = 'user-e2e';
const SESSION = { userId: USER, tenantId: TENANT };

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () =>
    Promise.resolve({ userId: USER, session: SESSION } as unknown as AuthResult),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => Promise.resolve(['mail']),
  getAllowedSites: () => Promise.resolve([]),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  isServiceDisabledForUser: () => Promise.resolve(false),
}));

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: () => Promise.resolve(false),
}));

jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: () => Promise.resolve(TENANT),
  getTenantIdFromSession: (session: { tenantId?: string }) => session?.tenantId ?? TENANT,
  getValidAccessTokenForSession: () => Promise.resolve('fake-token'),
  getValidAccessToken: () => Promise.resolve('fake-token'),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: (...args: unknown[]) =>
    (mockCreateGraphClient as (...a: unknown[]) => unknown)(...args),
}));

jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: (graph: unknown, folderId: string) =>
    mockResolveMailFolderName(graph, folderId),
}));

// assertOpaqueId is a no-op here so tests can use plain string IDs.
jest.mock('../services/opaqueId.js', () => ({
  assertOpaqueId: jest.fn(),
  assertOpaqueIds: jest.fn(),
  ValidationError: class ValidationError extends Error {},
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
// Import all four handlers so their app.http registrations are captured.
import '../functions/mail/listFoldersMail.js';
import '../functions/mail/searchMail.js';
import '../functions/mail/readMessage.js';
import '../functions/mail/moveMessage.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;

function handlerFor(name: string) {
  const reg = httpMock.mock.calls.find((call) => call[0] === name);
  if (!reg) throw new Error(`${name} handler was not registered`);
  return reg[1].handler;
}

const listFoldersHandler = handlerFor('listFoldersMail');
const searchHandler = handlerFor('searchMail');
const readHandler = handlerFor('readMessage');
const moveHandler = handlerFor('moveMessage');

// ── Request helper ──────────────────────────────────────────────────────────────

function makeRequest(opts: {
  query?: Record<string, string>;
  params?: Record<string, string>;
  json?: unknown;
}): HttpRequest {
  const query = opts.query ?? {};
  return {
    query: { get: (k: string) => query[k] ?? null },
    params: opts.params ?? {},
    json: () => Promise.resolve(opts.json ?? {}),
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

const ctx = { error: jest.fn(), log: jest.fn() } as unknown as InvocationContext;

function reset() {
  globalEntries.length = 0;
  userEntries.length = 0;
  graphCalls.length = 0;
  folderNames.clear();
  jest.clearAllMocks();
  delete process.env.DEFAULT_MAIL_DENY_FOLDERS;
}

function pushGlobalMail(path: string) {
  globalEntries.push({ partitionKey: `${TENANT}:mail`, rowKey: Buffer.from(path).toString('base64'), path });
}

beforeEach(reset);

// ── list_folders_mail ───────────────────────────────────────────────────────────

describe('E2E: list_folders_mail strips denied folders', () => {
  function graphReturnsFolders(names: string[]) {
    mockGraphGet.mockResolvedValue({
      value: names.map((displayName, i) => ({
        id: `folder-${i}`,
        displayName,
        totalItemCount: 0,
        unreadItemCount: 0,
        childFolderCount: 0,
      })),
    });
  }

  it('removes DEFAULT_MAIL_DENY_FOLDERS even with an empty admin table', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance,HR,Legal,IR,Management';
    graphReturnsFolders(['Inbox', 'Finance', 'Sent Items', 'HR', 'Legal', 'IR', 'Management', 'Archive']);

    const res = await listFoldersHandler(makeRequest({}), ctx);

    expect(res.status).toBe(200);
    // The tool call actually reached Graph.
    expect(graphCalls.some((c) => c.op === 'get' && c.path.endsWith('/mailFolders'))).toBe(true);
    const body = res.jsonBody as { folders: Array<{ name: string }>; count: number };
    expect(body.folders.map((f) => f.name)).toEqual(['Inbox', 'Sent Items', 'Archive']);
    expect(body.count).toBe(3);
  });

  it('removes an admin-table-denied folder on top of the defaults', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance';
    pushGlobalMail('Board Reports');
    graphReturnsFolders(['Inbox', 'Finance', 'Board Reports']);

    const res = await listFoldersHandler(makeRequest({}), ctx);

    const body = res.jsonBody as { folders: Array<{ name: string }> };
    expect(body.folders.map((f) => f.name)).toEqual(['Inbox']);
  });
});

// ── search_mail (global) ─────────────────────────────────────────────────────────

describe('E2E: search_mail strips messages in denied folders', () => {
  it('drops results whose parent folder is default-denied', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance,HR';
    folderNames.set('f-inbox', 'Inbox');
    folderNames.set('f-finance', 'Finance');
    folderNames.set('f-hr', 'HR');
    mockGraphGet.mockResolvedValue({
      value: [
        { id: 'm1', subject: 'Team lunch', parentFolderId: 'f-inbox' },
        { id: 'm2', subject: 'Q3 budget', parentFolderId: 'f-finance' },
        { id: 'm3', subject: 'Schedule', parentFolderId: 'f-inbox' },
        { id: 'm4', subject: 'Comp review', parentFolderId: 'f-hr' },
      ],
    });

    const res = await searchHandler(makeRequest({ query: { q: 'meeting' } }), ctx);

    expect(res.status).toBe(200);
    expect(graphCalls.some((c) => c.op === 'get')).toBe(true); // reached Graph
    const body = res.jsonBody as { results: Array<{ id: string }>; count: number };
    expect(body.results.map((r) => r.id)).toEqual(['m1', 'm3']);
    expect(body.count).toBe(2);
  });
});

// ── read_message ─────────────────────────────────────────────────────────────────

describe('E2E: read_message blocks a message in a denied folder', () => {
  it('returns 403 when the message lives in a default-denied folder', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance';
    folderNames.set('f-finance', 'Finance');
    mockGraphGet.mockResolvedValue({ id: 'm9', subject: 'Payroll', parentFolderId: 'f-finance' });

    const res = await readHandler(makeRequest({ params: { messageId: 'm9' } }), ctx);

    expect(res.status).toBe(403);
  });

  it('returns 200 for a message in an allowed folder', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance';
    folderNames.set('f-inbox', 'Inbox');
    mockGraphGet.mockResolvedValue({
      id: 'm10',
      subject: 'Hello',
      parentFolderId: 'f-inbox',
      body: { content: 'hi', contentType: 'text' },
    });

    const res = await readHandler(makeRequest({ params: { messageId: 'm10' } }), ctx);

    expect(res.status).toBe(200);
    expect(graphCalls.some((c) => c.op === 'get' && c.path.includes('/messages/m10'))).toBe(true);
    expect((res.jsonBody as { id: string }).id).toBe('m10');
  });
});

// ── move_message ─────────────────────────────────────────────────────────────────

describe('E2E: move_message enforces deny list on the destination', () => {
  it('returns 403 when moving into a default-denied folder (no Graph move call)', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance';
    folderNames.set('f-inbox', 'Inbox');
    folderNames.set('f-finance', 'Finance');
    // First Graph get() resolves the source folder of the message.
    mockGraphGet.mockResolvedValue({ parentFolderId: 'f-inbox' });

    const res = await moveHandler(
      makeRequest({ params: { messageId: 'm11' }, json: { destinationFolderId: 'f-finance' } }),
      ctx,
    );

    expect(res.status).toBe(403);
    // The move must never be issued to Graph.
    expect(graphCalls.some((c) => c.op === 'post')).toBe(false);
  });

  it('allows a move into a permitted folder and issues the Graph move', async () => {
    process.env.DEFAULT_MAIL_DENY_FOLDERS = 'Finance';
    folderNames.set('f-inbox', 'Inbox');
    folderNames.set('f-archive', 'Archive');
    mockGraphGet.mockResolvedValue({ parentFolderId: 'f-inbox' });
    mockGraphPost.mockResolvedValue({ id: 'm12', subject: 'x', parentFolderId: 'f-archive' });

    const res = await moveHandler(
      makeRequest({ params: { messageId: 'm12' }, json: { destinationFolderId: 'f-archive' } }),
      ctx,
    );

    expect(res.status).toBe(200);
    expect(graphCalls.some((c) => c.op === 'post' && c.path.endsWith('/move'))).toBe(true);
  });
});
