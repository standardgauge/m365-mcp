/**
 * The MCP endpoint refuses an oversized request body before it parses it and
 * before it authenticates, so an unauthenticated caller inside the rate limit
 * cannot make a replica buffer and parse a large payload.
 *
 * Covers:
 *   - a declared Content-Length over the cap is refused on the header alone:
 *     413, body never read, no authentication
 *   - an undeclared (or lying) oversized body is refused after it is read and
 *     before it is parsed, still without authentication
 *   - a body that is not JSON is a -32700 parse error, not a 500
 *   - a body under the cap flows through as before
 *   - MCP_MAX_BODY_BYTES overrides the default and ignores nonsense values
 */
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const mockAuthenticateRequest = jest.fn<() => Promise<unknown>>();

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({ authenticateRequest: () => mockAuthenticateRequest() }));
jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: async () => 'at',
  getTenantIdFromSession: () => 'tenant-a',
}));
jest.mock('../services/graphClient.js', () => ({ createGraphClient: () => ({ api: () => ({ get: async () => ({}) }) }) }));
jest.mock('../services/denyList.js', () => ({ isPathDenied: async () => false, filterDeniedPaths: async (_t: unknown, _u: unknown, _s: unknown, i: unknown) => i }));
jest.mock('../services/serviceSettings.js', () => ({ ALL_SERVICE_KEYS: [], getEnabledServices: async () => [], getReadOnlyServices: async () => [], getAllowedSites: async () => [] }));
jest.mock('../services/userServiceOverrides.js', () => ({ getUserServiceOverrides: async () => [] }));
jest.mock('../services/userMailConfig.js', () => ({ isMailIndexingDisabled: async () => false }));
jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('../services/securityHeaders.js', () => ({ withSecurity: (h: unknown) => h }));
jest.mock('../services/rateLimit.js', () => ({ withRateLimit: (_r: unknown, h: unknown) => h }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import { DEFAULT_MAX_BODY_BYTES, maxBodyBytes } from '../functions/mcp/mcpEndpoint.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const handler = httpMock.mock.calls.find((call) => call[0] === 'mcpEndpoint')![1].handler;

function post(text: string, headers: Record<string, string> = {}) {
  const textRead = jest.fn(() => Promise.resolve(text));
  const req = {
    method: 'POST',
    headers: new Map<string, string>(Object.entries(headers)),
    text: textRead,
    json: () => Promise.resolve(JSON.parse(text)),
  } as unknown as HttpRequest;
  const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
  return { res: handler(req, ctx), textRead };
}

const ping = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'ping' });

beforeEach(() => {
  mockAuthenticateRequest.mockReset();
  mockAuthenticateRequest.mockResolvedValue(null);
  delete process.env.MCP_MAX_BODY_BYTES;
});

describe('MCP request body cap', () => {
  it('refuses a declared oversize body on the header alone, without reading or authenticating', async () => {
    const { res, textRead } = post(ping, { 'content-length': String(DEFAULT_MAX_BODY_BYTES + 1) });
    const r = await res;
    expect(r.status).toBe(413);
    expect(r.jsonBody).toEqual({ jsonrpc: '2.0', id: null, error: { code: -32600, message: `Invalid Request: body exceeds ${DEFAULT_MAX_BODY_BYTES} bytes` } });
    expect(textRead).not.toHaveBeenCalled();
    expect(mockAuthenticateRequest).not.toHaveBeenCalled();
  });

  it('refuses an undeclared oversize body after reading it and before parsing or authenticating', async () => {
    process.env.MCP_MAX_BODY_BYTES = '64';
    const big = JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'ping', params: { pad: 'x'.repeat(200) } });
    const { res, textRead } = post(big);
    const r = await res;
    expect(r.status).toBe(413);
    expect(textRead).toHaveBeenCalledTimes(1);
    expect(mockAuthenticateRequest).not.toHaveBeenCalled();
  });

  it('counts bytes, not characters', async () => {
    // Seven two-byte characters: the string is under the cap by character
    // count and over it by byte count, and bytes are what the cap means.
    const multi = JSON.stringify({ jsonrpc: '2.0', id: 1, method: '\u00e9'.repeat(7) });
    const cap = multi.length + 2;
    expect(Buffer.byteLength(multi, 'utf8')).toBeGreaterThan(cap);
    process.env.MCP_MAX_BODY_BYTES = String(cap);
    const r = await post(multi).res;
    expect(r.status).toBe(413);
  });

  it('turns a body that is not JSON into a -32700 parse error, not a 500', async () => {
    const r = await post('{"jsonrpc": "2.0", "id": 1, "method": ').res;
    expect(r.status).toBe(400);
    expect(r.jsonBody).toEqual({ jsonrpc: '2.0', id: null, error: { code: -32700, message: 'Parse error: body is not valid JSON' } });
    expect(mockAuthenticateRequest).not.toHaveBeenCalled();
  });

  it('serves a body under the cap as before', async () => {
    const r = await post(ping, { 'content-length': String(Buffer.byteLength(ping)) }).res;
    expect(r.status).toBe(200);
    expect(mockAuthenticateRequest).toHaveBeenCalledTimes(1);
  });

  it('reads the cap from MCP_MAX_BODY_BYTES and ignores nonsense', () => {
    expect(maxBodyBytes({})).toBe(DEFAULT_MAX_BODY_BYTES);
    expect(maxBodyBytes({ MCP_MAX_BODY_BYTES: '1048576' })).toBe(1048576);
    expect(maxBodyBytes({ MCP_MAX_BODY_BYTES: '0' })).toBe(DEFAULT_MAX_BODY_BYTES);
    expect(maxBodyBytes({ MCP_MAX_BODY_BYTES: '-5' })).toBe(DEFAULT_MAX_BODY_BYTES);
    expect(maxBodyBytes({ MCP_MAX_BODY_BYTES: 'lots' })).toBe(DEFAULT_MAX_BODY_BYTES);
    expect(maxBodyBytes({ MCP_MAX_BODY_BYTES: ' 2048 ' })).toBe(2048);
  });
});
