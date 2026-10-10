/**
 * Tests for the /api/mail/settings endpoint.
 *
 * Covers:
 *   - Authentication gating (401 when unauthenticated)
 *   - GET: returns own email output mode
 *   - GET: admin can query another user
 *   - GET: non-admin cannot query another user (403)
 *   - GET: defaults userId to caller when not provided
 *   - POST: user can update their own mode
 *   - POST: admin can update another user's mode
 *   - POST: non-admin cannot update another user (403)
 *   - POST: rejects invalid emailOutputMode (400)
 *   - POST: uses caller's userId when body omits userId
 *
 * Enforced draft mode:
 *   - GET: surfaces enforced / enforcedBy alongside the effective mode
 *   - POST: refused with 403 for the user and for an admin while a policy applies,
 *     nothing is written, and the refusal is logged as denied
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
type Settings = {
  emailOutputMode: string;
  preferredEmailOutputMode?: string;
  enforced?: boolean;
  enforcedBy?: 'tenant' | 'user' | null;
};
const mockGetUserEmailSettings = jest.fn<(tenantId: string, userId: string) => Promise<Settings>>();
const mockLogAccess = jest.fn<(entry: Record<string, unknown>) => void>();
const mockSetUserEmailSettings = jest.fn<
  (tenantId: string, userId: string, settings: { emailOutputMode: string }) => Promise<void>
>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: (userId: unknown) => mockGetTenantId(userId as string),
}));

jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: (tenantId: unknown, userId: unknown) =>
    mockGetUserEmailSettings(tenantId as string, userId as string),
  setUserEmailSettings: (tenantId: unknown, userId: unknown, settings: unknown) =>
    mockSetUserEmailSettings(
      tenantId as string,
      userId as string,
      settings as { emailOutputMode: string },
    ),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

jest.mock('../services/securityHeaders.js', () => ({
  withSecurity: (handler: unknown) => handler,
}));

jest.mock('../services/auditLog.js', () => ({
  logAccess: (entry: unknown) => mockLogAccess(entry as Record<string, unknown>),
}));

import { app } from '@azure/functions';
import '../functions/mail/manageEmailSettings.js';
import { SessionStoreUnavailableError } from '../services/sessionStoreError.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'manageEmailSettings');
if (!registration) throw new Error('manageEmailSettings handler was not registered');
const handler = registration[1].handler;

const TENANT = 'test-tenant';
const CALLER_USER = 'caller-user';
const OTHER_USER = 'other-user';

const CALLER_AUTH: AuthResult = {
  userId: CALLER_USER,
  session: {
    userId: CALLER_USER,
    homeAccountId: 'home-caller',
    displayName: 'Caller',
    email: 'caller@example.com',
    tenantId: TENANT,
    accessToken: 'fake-token',
    expiresAt: Date.now() + 3_600_000,
    sessionToken: 'fake-session',
    sessionCreatedAt: Date.now(),
  },
};

function makeGetRequest(query?: Record<string, string>): HttpRequest {
  const queryMap = new Map(Object.entries(query ?? {}));
  return {
    method: 'GET',
    query: {
      get: (k: string) => queryMap.get(k) ?? null,
      has: (k: string) => queryMap.has(k),
    },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

function makePostRequest(body: unknown): HttpRequest {
  return {
    method: 'POST',
    query: {
      get: () => null,
      has: () => false,
    },
    headers: new Map<string, string>(),
    json: () => Promise.resolve(body),
  } as unknown as HttpRequest;
}

function makeContext(): InvocationContext {
  return { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
}

beforeEach(() => {
  jest.clearAllMocks();
  mockGetTenantId.mockResolvedValue(TENANT);
  mockAuthenticateRequest.mockResolvedValue(CALLER_AUTH);
  mockCheckGlobalAdmin.mockResolvedValue(false);
  mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', preferredEmailOutputMode: 'draft', enforced: false, enforcedBy: null });
});

describe('Authentication', () => {
  it('returns 401 when not authenticated', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);

    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(401);
    expect((res.jsonBody as { error: string }).error).toContain('Authentication required');
  });

  it('returns 503, not 401, when the session store cannot be reached', async () => {
    mockAuthenticateRequest.mockRejectedValue(
      new SessionStoreUnavailableError(new Error('AuthorizationFailure: storage-account-detail')),
    );

    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(503);
    expect(res.jsonBody).toEqual({ error: 'Session store unavailable, retry shortly' });
    expect(mockCheckGlobalAdmin).not.toHaveBeenCalled();
    expect(mockGetUserEmailSettings).not.toHaveBeenCalled();
  });
});

describe('GET /api/mail/settings', () => {
  it('returns own email output mode when no userId param', async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft' });

    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { userId: string; emailOutputMode: string };
    expect(body.userId).toBe(CALLER_USER);
    expect(body.emailOutputMode).toBe('draft');
    expect(mockGetUserEmailSettings).toHaveBeenCalledWith(TENANT, CALLER_USER);
  });

  it('returns own email output mode when userId param matches caller', async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });

    const res = await handler(makeGetRequest({ userId: CALLER_USER }), makeContext());

    expect(res.status).toBe(200);
    expect((res.jsonBody as { emailOutputMode: string }).emailOutputMode).toBe('send');
  });

  it('returns 403 when non-admin queries another user', async () => {
    const res = await handler(makeGetRequest({ userId: OTHER_USER }), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('your own');
    expect(mockGetUserEmailSettings).not.toHaveBeenCalled();
  });

  it('allows admin to query another user', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'send' });

    const res = await handler(makeGetRequest({ userId: OTHER_USER }), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { userId: string; emailOutputMode: string };
    expect(body.userId).toBe(OTHER_USER);
    expect(mockGetUserEmailSettings).toHaveBeenCalledWith(TENANT, OTHER_USER);
  });
});

describe('POST /api/mail/settings', () => {
  it('allows user to set their own mode to "send"', async () => {
    mockSetUserEmailSettings.mockResolvedValue(undefined);

    const res = await handler(makePostRequest({ emailOutputMode: 'send' }), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { ok: boolean; userId: string; emailOutputMode: string };
    expect(body.ok).toBe(true);
    expect(body.userId).toBe(CALLER_USER);
    expect(body.emailOutputMode).toBe('send');
    expect(mockSetUserEmailSettings).toHaveBeenCalledWith(TENANT, CALLER_USER, { emailOutputMode: 'send' });
  });

  it('allows user to set their own mode to "draft"', async () => {
    mockSetUserEmailSettings.mockResolvedValue(undefined);

    const res = await handler(makePostRequest({ emailOutputMode: 'draft' }), makeContext());

    expect(res.status).toBe(200);
    expect(mockSetUserEmailSettings).toHaveBeenCalledWith(TENANT, CALLER_USER, { emailOutputMode: 'draft' });
  });

  it('uses body.userId when provided and caller is admin', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockSetUserEmailSettings.mockResolvedValue(undefined);

    const res = await handler(
      makePostRequest({ userId: OTHER_USER, emailOutputMode: 'send' }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    expect(mockSetUserEmailSettings).toHaveBeenCalledWith(TENANT, OTHER_USER, { emailOutputMode: 'send' });
  });

  it('returns 403 when non-admin targets another user', async () => {
    const res = await handler(
      makePostRequest({ userId: OTHER_USER, emailOutputMode: 'send' }),
      makeContext(),
    );

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('your own');
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
  });

  it('returns 400 for invalid emailOutputMode', async () => {
    const res = await handler(makePostRequest({ emailOutputMode: 'immediate' }), makeContext());

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('"draft" or "send"');
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
  });

  it('returns 400 when emailOutputMode is missing', async () => {
    const res = await handler(makePostRequest({}), makeContext());

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('"draft" or "send"');
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
  });
});

describe('Enforced draft mode', () => {
  it('GET surfaces enforced and enforcedBy with the effective mode', async () => {
    mockGetUserEmailSettings.mockResolvedValue({
      emailOutputMode: 'draft', preferredEmailOutputMode: 'send', enforced: true, enforcedBy: 'tenant',
    });

    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({
      userId: CALLER_USER,
      emailOutputMode: 'draft',
      preferredEmailOutputMode: 'send',
      enforced: true,
      enforcedBy: 'tenant',
    });
  });

  it('GET reports enforced=false and enforcedBy=null when no policy applies', async () => {
    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(200);
    expect(res.jsonBody).toMatchObject({ enforced: false, enforcedBy: null });
  });

  it.each([['tenant'], ['user']] as const)('POST by the user is refused with 403 under a %s policy, writes nothing, logs denied', async (enforcedBy) => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: true, enforcedBy });

    const res = await handler(makePostRequest({ emailOutputMode: 'send' }), makeContext());

    expect(res.status).toBe(403);
    expect(res.jsonBody).toMatchObject({ enforced: true, enforcedBy });
    expect((res.jsonBody as { error: string }).error).toContain('enforced to draft');
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      tenantId: TENANT,
      userId: CALLER_USER,
      operation: 'set_email_output_mode',
      resource: CALLER_USER,
      result: 'denied',
      source: 'http',
      reason: expect.stringContaining(`enforced by ${enforcedBy} policy`),
    }));
  });

  it('POST by an admin targeting an enforced user is refused too — the policy endpoint is the only way out', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: true, enforcedBy: 'user' });

    const res = await handler(makePostRequest({ userId: OTHER_USER, emailOutputMode: 'send' }), makeContext());

    expect(res.status).toBe(403);
    expect(mockGetUserEmailSettings).toHaveBeenCalledWith(TENANT, OTHER_USER);
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ result: 'denied', resource: OTHER_USER }));
  });

  it("POST of 'draft' is refused while enforced as well — nothing is written under a policy", async () => {
    mockGetUserEmailSettings.mockResolvedValue({ emailOutputMode: 'draft', enforced: true, enforcedBy: 'tenant' });

    const res = await handler(makePostRequest({ emailOutputMode: 'draft' }), makeContext());

    expect(res.status).toBe(403);
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
  });

  it('POST validates the body before consulting the policy', async () => {
    const res = await handler(makePostRequest({ emailOutputMode: 'immediate' }), makeContext());

    expect(res.status).toBe(400);
    expect(mockGetUserEmailSettings).not.toHaveBeenCalled();
  });

  it('POST writes and does not log a denial when no policy applies', async () => {
    mockSetUserEmailSettings.mockResolvedValue(undefined);

    const res = await handler(makePostRequest({ emailOutputMode: 'send' }), makeContext());

    expect(res.status).toBe(200);
    expect(mockSetUserEmailSettings).toHaveBeenCalledWith(TENANT, CALLER_USER, { emailOutputMode: 'send' });
    expect(mockLogAccess).not.toHaveBeenCalled();
  });
});
