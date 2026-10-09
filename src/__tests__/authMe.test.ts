/**
 * Tests for the /api/auth/me endpoint.
 *
 * This endpoint is the admin SPA's session probe: it resolves identity from
 * the server session (mcp_session cookie / bearer) rather than a client Graph
 * token, and reports isGlobalAdmin so the SPA can unlock the admin view.
 *
 * Covers:
 *   - No session → 401 with a loginUrl for the SPA to redirect to
 *   - Authenticated non-admin → 200 with identity + isGlobalAdmin:false
 *   - Authenticated global admin → isGlobalAdmin:true
 *   - Console session present → consoleSession:true and the console cookie is
 *     re-issued with its idle expiry pushed out, issuedAt kept
 *   - Session without a console session → consoleSession:false, no cookie
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import { randomBytes } from 'crypto';
import type { AuthResult, ConsoleAuthResult } from '../services/authMiddleware.js';
import { CONSOLE_IDLE_MS, verifyConsoleToken } from '../services/consoleSession.js';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockAuthenticateConsoleRequest = jest.fn<(req: HttpRequest) => Promise<ConsoleAuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  authenticateConsoleRequest: (req: unknown) => mockAuthenticateConsoleRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/auth/me.js';

interface HttpRegistration {
  handler: (
    req: HttpRequest,
    context: InvocationContext,
  ) => Promise<{ status: number; jsonBody?: unknown; cookies?: Array<{ name: string; value: string; maxAge?: number }> }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const reg = httpMock.mock.calls.find((c) => c[0] === 'authMe');
if (!reg) throw new Error('authMe handler was not registered');
const handler = reg[1].handler;

const USER = 'user-123';
const AUTH = {
  userId: USER,
  session: { userId: USER, displayName: 'Test User', email: 'nate@example.com' },
} as unknown as AuthResult;
const ctx = { error: jest.fn(), warn: jest.fn() } as unknown as InvocationContext;

function req(): HttpRequest {
  return { method: 'GET', headers: new Map<string, string>() } as unknown as HttpRequest;
}

beforeEach(() => {
  jest.clearAllMocks();
  mockAuthenticateConsoleRequest.mockResolvedValue(null);
});

describe('authMe — /api/auth/me', () => {
  it('returns 401 with a loginUrl when there is no session', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);
    const res = await handler(req(), ctx);
    expect(res.status).toBe(401);
    expect(res.jsonBody).toEqual({ authenticated: false, loginUrl: '/api/auth/login' });
    expect(mockCheckGlobalAdmin).not.toHaveBeenCalled();
  });

  it('returns identity with isGlobalAdmin:false for a non-admin session', async () => {
    mockAuthenticateRequest.mockResolvedValue(AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(false);
    const res = await handler(req(), ctx);
    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({
      authenticated: true,
      userId: USER,
      displayName: 'Test User',
      email: 'nate@example.com',
      isGlobalAdmin: false,
      consoleSession: false,
    });
    expect(res.cookies).toBeUndefined();
  });

  it('reports isGlobalAdmin:true for a global admin session', async () => {
    mockAuthenticateRequest.mockResolvedValue(AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(true);
    const res = await handler(req(), ctx);
    expect(res.status).toBe(200);
    expect((res.jsonBody as { isGlobalAdmin: boolean }).isGlobalAdmin).toBe(true);
    expect(mockCheckGlobalAdmin).toHaveBeenCalledWith(USER);
  });

  it('renews the console cookie, keeping issuedAt, when a console session is present', async () => {
    const sessionToken = 'a'.repeat(64);
    const issuedAt = Date.now() - 60 * 60 * 1000; // signed in an hour ago
    mockAuthenticateConsoleRequest.mockResolvedValue({
      ...AUTH,
      sessionToken,
      console: { issuedAt, expiresAt: Date.now() + 60_000 },
    } as ConsoleAuthResult);
    mockCheckGlobalAdmin.mockResolvedValue(true);

    const before = Date.now();
    const res = await handler(req(), ctx);
    expect(res.status).toBe(200);
    const body = res.jsonBody as { consoleSession: boolean; consoleExpiresAt: number };
    expect(body.consoleSession).toBe(true);
    expect(body.consoleExpiresAt).toBeGreaterThanOrEqual(before + CONSOLE_IDLE_MS);

    const cookie = res.cookies?.find((c) => c.name === 'mcp_console');
    const claims = verifyConsoleToken(cookie?.value, sessionToken);
    expect(claims).toEqual({ issuedAt, expiresAt: body.consoleExpiresAt });
    expect(mockAuthenticateRequest).not.toHaveBeenCalled();
  });
});
