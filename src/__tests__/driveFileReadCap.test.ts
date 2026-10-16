/**
 * The readable-size cap on SharePoint and OneDrive file reads.
 *
 * `read_file` and `read_onedrive_file` (MCP) and their REST routes return the whole
 * file inline. Before the cap they streamed any file into memory regardless of size,
 * so a large file meant high memory, slow responses and function timeouts. Covered:
 *   - a file whose metadata size is over MAX_READ_FILE_BYTES is refused on all four
 *     surfaces with an error naming the size and the cap, and getStream() is never called;
 *   - a file at the cap is still read;
 *   - stale metadata (small declared size, large body) is stopped while streaming;
 *   - the MAX_READ_FILE_BYTES parser falls back to the default on bad values.
 */

import { jest } from '@jest/globals';
import { Readable } from 'node:stream';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

const TENANT = 'cap-tenant';
const USER = 'user-cap';
const SITE_ID = 'contoso.sharepoint.com,11111111-1111-1111-1111-111111111111,aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const SESSION = { userId: USER, tenantId: TENANT, email: 'cap@example.com' };

// ── Graph mock: one file whose declared size and body the tests set ────────────

let declaredSize: number | undefined;
let body: Buffer;
const getStreamCalls: string[] = [];

function graphApi(path: string) {
  const chain = {
    select: () => chain,
    get: async () => ({
      id: 'item-1',
      name: 'big.bin',
      webUrl: 'https://example.com/big.bin',
      ...(declaredSize !== undefined ? { size: declaredSize } : {}),
      file: { mimeType: 'application/octet-stream' },
      parentReference: { path: '/drive/root:/Shared' },
    }),
    getStream: async () => {
      getStreamCalls.push(path);
      // Two chunks so the streaming cap can trip part-way.
      const half = Math.floor(body.length / 2);
      return Readable.from([body.subarray(0, half), body.subarray(half)]);
    },
  };
  return chain;
}

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => Promise.resolve({ userId: USER, session: SESSION } as unknown as AuthResult),
  checkGlobalAdmin: () => Promise.resolve(false),
}));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => Promise.resolve(['sharepoint', 'onedrive']),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => Promise.resolve([]),
}));
jest.mock('../services/userServiceOverrides.js', () => ({
  isServiceDisabledForUser: () => Promise.resolve(false),
  getUserServiceOverrides: () => Promise.resolve([]),
}));
jest.mock('../services/userMailConfig.js', () => ({ isMailIndexingDisabled: () => Promise.resolve(false) }));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: () => Promise.resolve({ emailOutputMode: 'draft' }),
  setUserEmailSettings: () => Promise.resolve(undefined),
}));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('../services/denyList.js', () => ({
  isPathDenied: () => Promise.resolve(false),
  filterDeniedPaths: (_t: string, _u: string, _s: string, items: unknown[]) => Promise.resolve(items),
}));
jest.mock('../services/tokenCache.js', () => ({
  getTenantId: () => Promise.resolve(TENANT),
  getTenantIdFromSession: () => TENANT,
  getValidAccessTokenForSession: () => Promise.resolve('fake-token'),
}));
jest.mock('../services/graphClient.js', () => ({ createGraphClient: () => ({ api: graphApi }) }));
jest.mock('../services/opaqueId.js', () => ({
  assertOpaqueId: jest.fn(),
  assertOpaqueIds: jest.fn(),
  ValidationError: class ValidationError extends Error {},
}));
jest.mock('../services/securityHeaders.js', () => ({ withSecurity: (handler: unknown) => handler }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';
import '../functions/sharepoint/readFile.js';
import '../functions/onedrive/readOneDriveFile.js';
import { DEFAULT_MAX_READ_FILE_BYTES, maxReadFileBytes } from '../services/driveFileRead.js';

type Handler = (req: HttpRequest, ctx: InvocationContext, auth?: AuthResult) => Promise<{ status: number; jsonBody?: unknown }>;
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: { handler: Handler }) => void>;
// Captured at load: beforeEach's clearAllMocks wipes the registration calls.
const routes = new Map(httpMock.mock.calls.map((c) => [c[0], c[1].handler]));
const route = (name: string): Handler => {
  const h = routes.get(name);
  if (!h) throw new Error(`${name} not registered`);
  return h;
};

const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
const AUTH = { userId: USER, session: SESSION } as unknown as AuthResult;

async function callTool(name: string, args: Record<string, unknown>) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } }),
    text: () => Promise.resolve(JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'tools/call', params: { name, arguments: args } })),
  } as unknown as HttpRequest;
  const res = await route('mcpEndpoint')(req, ctx);
  const result = (res.jsonBody as { result?: { content?: Array<{ text: string }>; isError?: boolean } }).result;
  return { text: result?.content?.[0]?.text ?? '', isError: result?.isError ?? false };
}

function restRequest(query: Record<string, string>): HttpRequest {
  return {
    query: new URLSearchParams(query),
    params: { itemId: 'item-1' },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

const CAP = 1024;

const surfaces: Array<{ name: string; read: () => Promise<{ ok: boolean; status?: number; text: string }> }> = [
  {
    name: 'MCP read_file',
    read: async () => {
      const r = await callTool('read_file', { siteId: SITE_ID, itemId: 'item-1' });
      return { ok: !r.isError, text: r.text };
    },
  },
  {
    name: 'MCP read_onedrive_file',
    read: async () => {
      const r = await callTool('read_onedrive_file', { itemId: 'item-1' });
      return { ok: !r.isError, text: r.text };
    },
  },
  {
    name: 'REST GET /api/sharepoint/files/{itemId}',
    read: async () => {
      const r = await route('readFile')(restRequest({ siteId: SITE_ID }), ctx, AUTH);
      return { ok: r.status === 200, status: r.status, text: JSON.stringify(r.jsonBody) };
    },
  },
  {
    name: 'REST GET /api/onedrive/files/{itemId}',
    read: async () => {
      const r = await route('readOneDriveFile')(restRequest({}), ctx, AUTH);
      return { ok: r.status === 200, status: r.status, text: JSON.stringify(r.jsonBody) };
    },
  },
];

beforeEach(() => {
  jest.clearAllMocks();
  getStreamCalls.length = 0;
  process.env.MAX_READ_FILE_BYTES = String(CAP);
});

afterAll(() => {
  delete process.env.MAX_READ_FILE_BYTES;
});

describe.each(surfaces)('$name', ({ read }) => {
  it('refuses a file whose metadata size is over the cap without calling getStream()', async () => {
    declaredSize = 5 * 1024 * 1024 * 1024;
    body = Buffer.alloc(0);

    const r = await read();

    expect(r.ok).toBe(false);
    if (r.status !== undefined) expect(r.status).toBe(413);
    expect(r.text).toContain(`${declaredSize} bytes`);
    expect(r.text).toContain(`${CAP}-byte limit`);
    expect(getStreamCalls).toEqual([]);
  });

  it('reads a file exactly at the cap', async () => {
    declaredSize = CAP;
    body = Buffer.alloc(CAP, 7);

    const r = await read();

    expect(r.ok).toBe(true);
    expect(r.text).toContain(body.toString('base64').slice(0, 64));
    expect(getStreamCalls).toHaveLength(1);
  });

  it('stops streaming when stale metadata understates the size', async () => {
    declaredSize = 10;
    body = Buffer.alloc(CAP * 4, 1);

    const r = await read();

    expect(r.ok).toBe(false);
    if (r.status !== undefined) expect(r.status).toBe(413);
    expect(r.text).toContain(`${CAP}-byte limit`);
  });
});

describe('maxReadFileBytes', () => {
  afterEach(() => {
    process.env.MAX_READ_FILE_BYTES = String(CAP);
  });

  it('defaults to 150 MB when unset or empty', () => {
    delete process.env.MAX_READ_FILE_BYTES;
    expect(maxReadFileBytes()).toBe(DEFAULT_MAX_READ_FILE_BYTES);
    expect(DEFAULT_MAX_READ_FILE_BYTES).toBe(150 * 1024 * 1024);
    process.env.MAX_READ_FILE_BYTES = '  ';
    expect(maxReadFileBytes()).toBe(DEFAULT_MAX_READ_FILE_BYTES);
  });

  it('honours a positive integer override', () => {
    process.env.MAX_READ_FILE_BYTES = '5242880';
    expect(maxReadFileBytes()).toBe(5242880);
  });

  it.each(['0', '-1', '10MB', '1.5', '1e9', '99999999999999999999'])('falls back to the default for %p', (raw) => {
    const spy = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    process.env.MAX_READ_FILE_BYTES = raw;
    expect(maxReadFileBytes()).toBe(DEFAULT_MAX_READ_FILE_BYTES);
    spy.mockRestore();
  });
});
