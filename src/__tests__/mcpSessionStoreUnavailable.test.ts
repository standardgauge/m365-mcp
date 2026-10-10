/**
 * The MCP endpoint answers 503 when session storage is down.
 *
 * authenticateRequest throws SessionStoreUnavailableError when storage cannot
 * say whether the token is a session. The endpoint must not treat that as an
 * unauthenticated call, which tells the client to re-authenticate.
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
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/mcp/mcpEndpoint.js';
import { SessionStoreUnavailableError } from '../services/sessionStoreError.js';

interface HttpRegistration {
  handler: (
    req: HttpRequest,
    context: InvocationContext,
  ) => Promise<{ status: number; jsonBody?: unknown; headers?: Record<string, string> }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const handler = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint')![1].handler;

function call(body: unknown) {
  const req = {
    method: 'POST',
    url: 'https://mcp.example.com/api/mcp',
    headers: new Map<string, string>([['x-forwarded-for', '203.0.113.7']]),
    json: () => Promise.resolve(body),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return handler(req, ctx);
}

beforeEach(() => {
  mockAuthenticateRequest.mockReset();
  jest.spyOn(console, 'error').mockImplementation(() => {});
});
afterEach(() => jest.restoreAllMocks());

describe('MCP endpoint, session store unavailable', () => {
  it('returns 503 with a JSON-RPC error instead of an unauthenticated result', async () => {
    mockAuthenticateRequest.mockRejectedValue(new SessionStoreUnavailableError(new Error('ServerBusy')));

    const out = await call({ jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} });

    expect(out.status).toBe(503);
    expect(out.headers).toMatchObject({ 'Retry-After': '5' });
    expect(out.jsonBody).toEqual({
      jsonrpc: '2.0',
      id: null,
      error: { code: -32603, message: 'Session store unavailable, retry shortly' },
    });
  });

  it('still serves an unauthenticated call when storage answers that there is no session', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);

    const out = await call({ jsonrpc: '2.0', id: 1, method: 'tools/list', params: {} });

    expect(out.status).toBe(200);
  });
});
