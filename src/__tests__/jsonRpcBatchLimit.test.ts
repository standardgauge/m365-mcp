/**
 * A JSON-RPC batch is capped at MAX_JSONRPC_BATCH messages. An oversized or
 * empty batch is refused with -32600 before the request is authenticated, so
 * it costs no session lookup.
 */
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => mockAuthenticateRequest(),
}));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: () => Promise.resolve('access-token'),
  getTenantIdFromSession: () => 't',
}));
jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => ({ api: () => ({}) }),
}));
jest.mock('../services/denyList.js', () => ({
  filterDeniedPaths: (_t: unknown, _u: unknown, _s: unknown, items: unknown) => Promise.resolve(items),
  filterDeniedSearchHits: (_t: unknown, _u: unknown, items: unknown) => Promise.resolve(items),
  isPathDenied: () => Promise.resolve(false),
  canonicalizePath: (p: string) => (p.startsWith('/') ? p : `/${p}`).replace(/\/+$/, '').toLowerCase(),
}));
jest.mock('../services/calendarAccess.js', () => ({
  checkCalendarAccess: () => Promise.resolve(null),
}));
jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => Promise.resolve(null),
  resolveContactParentFolder: () => Promise.resolve(null),
  resolveDefaultContactFolder: () => Promise.resolve(null),
  resolveSectionNotebook: () => Promise.resolve(null),
  resolveDefaultCalendarId: () => Promise.resolve('cal'),
  resolveCalendarName: () => Promise.resolve('Calendar'),
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id),
}));
jest.mock('../services/sharepointFilter.js', () => ({
  filterAndDisambiguateSites: (sites: unknown) => sites,
}));
jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => Promise.resolve(['mail']),
  getReadOnlyServices: () => Promise.resolve([]),
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
jest.mock('../services/mailboxTimeZone.js', () => ({
  resolveMailboxTimeZone: () => Promise.resolve('Pacific Standard Time'),
}));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import { MAX_JSONRPC_BATCH } from '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const handler = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint')![1].handler;

function post(body: unknown) {
  const req = {
    method: 'POST',
    headers: new Map<string, string>(),
    json: () => Promise.resolve(body),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

const ping = (id: number) => ({ jsonrpc: '2.0', id, method: 'ping' });

beforeEach(() => {
  mockAuthenticateRequest.mockReset();
  mockAuthenticateRequest.mockResolvedValue(null);
});

describe('JSON-RPC batch size', () => {
  test('a batch at the cap is served', async () => {
    const res = await post(Array.from({ length: MAX_JSONRPC_BATCH }, (_, i) => ping(i)));
    expect(res.status).toBe(200);
    expect(res.jsonBody).toHaveLength(MAX_JSONRPC_BATCH);
  });

  test('a batch over the cap is refused without authenticating', async () => {
    const res = await post(Array.from({ length: MAX_JSONRPC_BATCH + 1 }, (_, i) => ping(i)));
    expect(res.status).toBe(400);
    expect(res.jsonBody).toEqual({
      jsonrpc: '2.0',
      id: null,
      error: { code: -32600, message: `Invalid Request: batch exceeds ${MAX_JSONRPC_BATCH} messages` },
    });
    expect(mockAuthenticateRequest).not.toHaveBeenCalled();
  });

  test('an empty batch is refused', async () => {
    const res = await post([]);
    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: { code: number } }).error.code).toBe(-32600);
    expect(mockAuthenticateRequest).not.toHaveBeenCalled();
  });

  test('a single message is unaffected', async () => {
    const res = await post(ping(1));
    expect(res.status).toBe(200);
    expect(mockAuthenticateRequest).toHaveBeenCalledTimes(1);
  });
});
