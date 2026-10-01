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
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/auth/me.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
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
    });
  });

  it('reports isGlobalAdmin:true for a global admin session', async () => {
    mockAuthenticateRequest.mockResolvedValue(AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(true);
    const res = await handler(req(), ctx);
    expect(res.status).toBe(200);
    expect((res.jsonBody as { isGlobalAdmin: boolean }).isGlobalAdmin).toBe(true);
    expect(mockCheckGlobalAdmin).toHaveBeenCalledWith(USER);
  });
});
