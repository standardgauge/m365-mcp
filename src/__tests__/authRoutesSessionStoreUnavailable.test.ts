/**
 * /api/auth/logout and /api/auth/refresh when session storage is down.
 *
 * Neither may treat an unanswered session lookup as "no session". Refresh
 * answers 503 rather than a 401 with a loginUrl, so the client retries instead
 * of signing in again. Logout answers 503 and leaves the cookies alone: it
 * could not identify the session, so it could not end it, and redirecting as
 * if signed out would leave a live server session behind.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockAuthenticateRequestAllowExpired = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockDeleteAllUserSessions = jest.fn<(userId: string) => Promise<void>>();
const mockStoreSession = jest.fn<(session: unknown) => Promise<void>>();
const mockAcquireTokenSilent = jest.fn<(homeAccountId: string) => Promise<unknown>>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  authenticateRequestAllowExpired: (req: unknown) => mockAuthenticateRequestAllowExpired(req as HttpRequest),
}));
jest.mock('../services/tokenCache.js', () => ({
  deleteAllUserSessions: (userId: unknown) => mockDeleteAllUserSessions(userId as string),
  storeSession: (session: unknown) => mockStoreSession(session),
}));
jest.mock('../services/graphClient.js', () => ({
  acquireTokenSilent: (id: unknown) => mockAcquireTokenSilent(id as string),
}));
jest.mock('../services/consoleSession.js', () => ({
  checkBrowserOrigin: () => ({ ok: true }),
  expiredConsoleCookie: () => ({ name: 'mcp_console', value: '', maxAge: 0 }),
}));
jest.mock('../services/frontendUrl.js', () => ({
  resolveFrontendUrl: () => 'https://app.example.com/',
}));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/auth/logout.js';
import '../functions/auth/refresh.js';
import { SessionStoreUnavailableError } from '../services/sessionStoreError.js';

type Response = {
  status: number;
  jsonBody?: unknown;
  headers?: Record<string, string>;
  cookies?: Array<{ name: string }>;
};
interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<Response>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
function handlerFor(name: string) {
  const reg = httpMock.mock.calls.find((c) => c[0] === name);
  if (!reg) throw new Error(`${name} handler was not registered`);
  return reg[1].handler;
}
const logout = handlerFor('authLogout');
const refresh = handlerFor('refresh');

const AUTH = {
  userId: 'user-1',
  session: { userId: 'user-1', homeAccountId: 'home-1' },
} as unknown as AuthResult;

function req(): HttpRequest {
  return { method: 'POST', headers: new Map<string, string>() } as unknown as HttpRequest;
}
function ctx(): InvocationContext {
  return { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
}
const outage = () => new SessionStoreUnavailableError(new Error('ServerBusy'));

beforeEach(() => {
  jest.clearAllMocks();
  mockDeleteAllUserSessions.mockResolvedValue(undefined);
  mockStoreSession.mockResolvedValue(undefined);
});

describe('POST /api/auth/logout', () => {
  it('returns 503 with the cookies untouched when the session store is down', async () => {
    mockAuthenticateRequest.mockRejectedValue(outage());

    const res = await logout(req(), ctx());

    expect(res.status).toBe(503);
    expect(res.headers).toMatchObject({ 'Retry-After': '5' });
    expect(res.headers?.Location).toBeUndefined();
    expect(res.cookies).toBeUndefined();
    expect(mockDeleteAllUserSessions).not.toHaveBeenCalled();
  });

  it('still signs out and clears cookies for a session storage knows', async () => {
    mockAuthenticateRequest.mockResolvedValue(AUTH);

    const res = await logout(req(), ctx());

    expect(res.status).toBe(303);
    expect(mockDeleteAllUserSessions).toHaveBeenCalledWith('user-1');
    expect(res.cookies?.map((c) => c.name)).toContain('mcp_session');
  });

  it('still clears cookies when storage answers that there is no session', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);

    const res = await logout(req(), ctx());

    expect(res.status).toBe(303);
    expect(res.cookies?.map((c) => c.name)).toContain('mcp_session');
  });
});

describe('POST /api/auth/refresh', () => {
  it('returns 503 without a loginUrl when the session store is down', async () => {
    mockAuthenticateRequestAllowExpired.mockRejectedValue(outage());

    const res = await refresh(req(), ctx());

    expect(res.status).toBe(503);
    expect(res.headers).toMatchObject({ 'Retry-After': '5' });
    expect(res.jsonBody).toEqual({ error: 'Session store unavailable, retry shortly' });
    expect(mockAcquireTokenSilent).not.toHaveBeenCalled();
    expect(mockStoreSession).not.toHaveBeenCalled();
  });

  it('still returns 401 with a loginUrl when storage answers that there is no session', async () => {
    mockAuthenticateRequestAllowExpired.mockResolvedValue(null);

    const res = await refresh(req(), ctx());

    expect(res.status).toBe(401);
    expect(res.jsonBody).toEqual({ refreshed: false, loginUrl: '/api/auth/login' });
  });
});
