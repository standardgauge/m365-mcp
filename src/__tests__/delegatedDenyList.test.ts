/**
 * Delegated mailbox and calendar access through the MCP tools/call surface is
 * checked against the data owner's per-user (tier 2) deny list, not only the
 * caller's (threat model §9.5).
 *
 * The REAL denyList and mailboxOwner services are in the loop; Azure Table
 * Storage, Graph, auth and the container resolver are mocked. The REST routes
 * get the same coverage in mailDenyListE2E.test.ts.
 */
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── In-memory Azure Table Storage (backs the REAL denyList service) ─────────

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
        listEntities: jest.fn(({ queryOptions }: { queryOptions: { filter: string } }) => {
          const source = tableName === 'UserDenyList' ? userEntries : [];
          const pk = queryOptions.filter.match(/PartitionKey eq '([^']+)'/)?.[1] ?? '';
          return makeAsyncIterable(source.filter((e) => e.partitionKey === pk));
        }),
      })),
    },
    odata: (strings: TemplateStringsArray, ...values: unknown[]) =>
      strings.reduce((acc, str, i) => `${acc}${str}${i < values.length ? `'${values[i]}'` : ''}`, ''),
  };
});
process.env.AZURE_STORAGE_CONNECTION_STRING = 'UseDevelopmentStorage=true';

// ── Graph mock ──────────────────────────────────────────────────────────────

const CALLER = 'delegate-oid';
const OWNER = 'owner-oid';
const OWNER_UPN = 'owner@example.com';

interface Call { path: string; op: 'get' | 'post' | 'patch' | 'delete' }
const calls: Call[] = [];
let respond: (path: string) => unknown = () => ({});

function makeChain(path: string) {
  const chain: Record<string, unknown> = {};
  for (const m of ['select', 'expand', 'header', 'top', 'filter', 'orderby', 'search']) chain[m] = () => chain;
  const record = (op: Call['op']) => () => {
    calls.push({ path, op });
    const r = respond(path);
    return r instanceof Error ? Promise.reject(r) : Promise.resolve(r);
  };
  chain.get = record('get');
  chain.post = record('post');
  chain.patch = record('patch');
  chain.delete = record('delete');
  return chain;
}

const folderNames = new Map<string, string>();

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () =>
    Promise.resolve({ userId: CALLER, session: { userId: CALLER, tenantId: 'tenant-1', accessToken: 'fake', sessionToken: 'sess' } }),
}));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => Promise.resolve('access-token'),
  getTenantIdFromSession: () => 'tenant-1',
  getTenantId: () => Promise.resolve('tenant-1'),
}));
jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => ({ api: (path: string) => makeChain(path) }),
}));
jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: (_g: unknown, id: string) => Promise.resolve(folderNames.get(id) ?? ''),
  resolveDefaultCalendarId: () => Promise.resolve(null),
  resolveCalendarName: () => Promise.resolve(''),
  resolveContactParentFolder: () => Promise.resolve(null),
  resolveDefaultContactFolder: () => Promise.resolve(null),
  resolveSectionNotebook: () => Promise.resolve(null),
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
}));
jest.mock('../services/sharepointFilter.js', () => ({ filterAndDisambiguateSites: (sites: unknown) => sites }));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => Promise.resolve(['mail', 'calendar']),
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
import { clearMailboxOwnerCache } from '../services/mailboxOwner.js';
import '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint');
if (!registration) throw new Error('mcpEndpoint handler was not registered');
const handler = registration[1].handler;

function callTool(name: string, args: Record<string, unknown>) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

function toolResult(res: { jsonBody?: unknown }): { isError: boolean; text: string } {
  const body = res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } };
  return { isError: body.result?.isError ?? false, text: body.result?.content?.[0]?.text ?? '' };
}

function hide(userId: string, type: 'mail' | 'calendar', path: string) {
  userEntries.push({ partitionKey: `${userId}:${type}`, rowKey: Buffer.from(path).toString('base64'), path });
}

const OWNER_LOOKUP = `/users/${encodeURIComponent(OWNER_UPN)}`;
const OWNER_BASE = `/users/${OWNER_UPN}`;

beforeEach(() => {
  userEntries.length = 0;
  calls.length = 0;
  folderNames.clear();
  clearMailboxOwnerCache();
  folderNames.set('f-pay', 'Payroll');
  folderNames.set('f-inbox', 'Inbox');
  respond = (path) => {
    if (path === OWNER_LOOKUP) return { id: OWNER };
    if (path.endsWith('/calendar')) return { id: 'owner-default-cal', name: 'Calendar' };
    if (path.endsWith('/mailFolders')) {
      return { value: [{ id: 'f-inbox', displayName: 'Inbox' }, { id: 'f-pay', displayName: 'Payroll' }] };
    }
    return { id: 'm1', parentFolderId: 'f-pay', subject: 'Salaries', body: { content: 'secret' } };
  };
});

describe('read_message', () => {
  it("is refused in the owner's mailbox when the owner hid the folder", async () => {
    hide(OWNER, 'mail', 'Payroll');
    const { isError, text } = toolResult(await callTool('read_message', { messageId: 'm1', mailboxId: OWNER_UPN }));
    expect(isError).toBe(true);
    expect(text).toContain('deny list');
    expect(text).not.toContain('secret');
  });

  it('is allowed when neither the owner nor the delegate hid the folder', async () => {
    const { isError, text } = toolResult(await callTool('read_message', { messageId: 'm1', mailboxId: OWNER_UPN }));
    expect(isError).toBe(false);
    expect(text).toContain('Salaries');
  });

  it("keeps applying the delegate's own list in the owner's mailbox", async () => {
    hide(CALLER, 'mail', 'Payroll');
    const { isError } = toolResult(await callTool('read_message', { messageId: 'm1', mailboxId: OWNER_UPN }));
    expect(isError).toBe(true);
  });

  it("does not apply another user's list, or look anyone up, in the caller's own mailbox", async () => {
    hide(OWNER, 'mail', 'Payroll');
    const { isError } = toolResult(await callTool('read_message', { messageId: 'm1' }));
    expect(isError).toBe(false);
    expect(calls.some((c) => c.path.startsWith('/users/'))).toBe(false);
  });

  it('fails closed, before reading the message, when the owner cannot be identified', async () => {
    respond = (path) => (path === OWNER_LOOKUP ? new Error('Resource not found') : { id: 'm1', parentFolderId: 'f-inbox' });
    const { isError } = toolResult(await callTool('read_message', { messageId: 'm1', mailboxId: OWNER_UPN }));
    expect(isError).toBe(true);
    expect(calls.some((c) => c.path.includes('/messages/'))).toBe(false);
  });
});

describe('list_folders_mail', () => {
  it("drops the owner's hidden folder from a delegated listing", async () => {
    hide(OWNER, 'mail', 'Payroll');
    const { isError, text } = toolResult(await callTool('list_folders_mail', { mailboxId: OWNER_UPN }));
    expect(isError).toBe(false);
    const names = (JSON.parse(text).items as Array<{ displayName: string }>).map((f) => f.displayName);
    expect(names).toEqual(['Inbox']);
  });
});

describe('mail writes in a delegated mailbox', () => {
  it("move_message refuses the owner's hidden folder as a destination", async () => {
    hide(OWNER, 'mail', 'Payroll');
    respond = (path) => (path === OWNER_LOOKUP ? { id: OWNER } : { parentFolderId: 'f-inbox' });
    const { isError } = toolResult(
      await callTool('move_message', { messageId: 'm1', destinationFolderId: 'f-pay', mailboxId: OWNER_UPN }),
    );
    expect(isError).toBe(true);
    expect(calls.some((c) => c.op === 'post')).toBe(false);
  });

  it("create_mail_folder refuses a name the owner hid", async () => {
    hide(OWNER, 'mail', 'Payroll');
    const { isError } = toolResult(await callTool('create_mail_folder', { displayName: 'Payroll', mailboxId: OWNER_UPN }));
    expect(isError).toBe(true);
    expect(calls.some((c) => c.op === 'post')).toBe(false);
  });
});

describe('respond_to_event', () => {
  it("is refused when the owner hid their default calendar", async () => {
    hide(OWNER, 'calendar', 'owner-default-cal');
    const { isError, text } = toolResult(
      await callTool('respond_to_event', { messageOrEventId: 'ev-1', response: 'accept', mailboxId: OWNER_UPN }),
    );
    expect(isError).toBe(true);
    expect(text).toContain('deny list');
    expect(calls.some((c) => c.op === 'post')).toBe(false);
    expect(calls.some((c) => c.path === `${OWNER_BASE}/calendar`)).toBe(true);
  });
});
