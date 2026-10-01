/**
 * Tests for the /api/manage/user-services admin endpoint.
 *
 * Covers:
 *   - Authentication gating (401 when unauthenticated)
 *   - Global Admin requirement (403 for non-admins)
 *   - GET: returns disabled services for target user
 *   - GET: 400 when userId query param missing
 *   - POST: sets disabled services for a user
 *   - POST: validates service keys (400 on invalid)
 *   - POST: 400 when body missing required fields
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetUserServiceOverrides = jest.fn<(tenantId: string, userId: string) => Promise<string[]>>();
const mockSetUserServiceOverrides = jest.fn<(tenantId: string, userId: string, disabled: string[]) => Promise<void>>();

// ── Wire up mocks ────────────────────────────────────────────────────────────

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: (userId: unknown) => mockGetTenantId(userId as string),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: (tenantId: unknown, userId: unknown) =>
    mockGetUserServiceOverrides(tenantId as string, userId as string),
  setUserServiceOverrides: (tenantId: unknown, userId: unknown, disabled: unknown) =>
    mockSetUserServiceOverrides(tenantId as string, userId as string, disabled as string[]),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import { app } from '@azure/functions';
import '../functions/admin/manageUserServices.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'manageUserServices');
if (!registration) throw new Error('manageUserServices handler was not registered');
const handler = registration[1].handler;

// ── Test fixtures ────────────────────────────────────────────────────────────

const TENANT = 'test-tenant';
const ADMIN_USER = 'admin-user';
const TARGET_USER = 'target-user';

const ADMIN_AUTH: AuthResult = {
  userId: ADMIN_USER,
  session: {
    userId: ADMIN_USER,
    homeAccountId: 'home-admin',
    displayName: 'Admin',
    email: 'admin@example.com',
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
  return {
    error: jest.fn(),
    warn: jest.fn(),
    log: jest.fn(),
  } as unknown as InvocationContext;
}

// ── Tests ────────────────────────────────────────────────────────────────────

beforeEach(() => {
  jest.clearAllMocks();
  mockGetTenantId.mockResolvedValue(TENANT);
});

describe('Authentication', () => {
  it('returns 401 when not authenticated', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);

    const res = await handler(makeGetRequest({ userId: TARGET_USER }), makeContext());

    expect(res.status).toBe(401);
    expect((res.jsonBody as { error: string }).error).toContain('Authentication required');
  });

  it('returns 403 when non-admin queries a different user', async () => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(makeGetRequest({ userId: TARGET_USER }), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('your own');
  });

  it('allows non-admin to query their own overrides', async () => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(false);
    mockGetUserServiceOverrides.mockResolvedValue(['mail']);

    // Query self — userId matches auth.userId
    const res = await handler(makeGetRequest({ userId: ADMIN_USER }), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { disabledServices: string[] };
    expect(body.disabledServices).toEqual(['mail']);
  });
});

describe('GET /api/manage/user-services', () => {
  beforeEach(() => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(true);
  });

  it('returns disabled services for the target user', async () => {
    mockGetUserServiceOverrides.mockResolvedValue(['mail', 'calendar']);

    const res = await handler(makeGetRequest({ userId: TARGET_USER }), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { userId: string; disabledServices: string[] };
    expect(body.userId).toBe(TARGET_USER);
    expect(body.disabledServices).toEqual(['mail', 'calendar']);
    expect(mockGetUserServiceOverrides).toHaveBeenCalledWith(TENANT, TARGET_USER);
  });

  it('returns empty array when user has no overrides', async () => {
    mockGetUserServiceOverrides.mockResolvedValue([]);

    const res = await handler(makeGetRequest({ userId: TARGET_USER }), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { disabledServices: string[] };
    expect(body.disabledServices).toEqual([]);
  });

  it('returns 400 when userId query param is missing', async () => {
    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('Missing userId');
  });
});

describe('POST /api/manage/user-services', () => {
  beforeEach(() => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockSetUserServiceOverrides.mockResolvedValue(undefined);
  });

  it('sets disabled services for a user', async () => {
    const res = await handler(
      makePostRequest({ userId: TARGET_USER, disabledServices: ['mail'] }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    const body = res.jsonBody as { ok: boolean; userId: string; disabledServices: string[] };
    expect(body.ok).toBe(true);
    expect(body.userId).toBe(TARGET_USER);
    expect(body.disabledServices).toEqual(['mail']);
    expect(mockSetUserServiceOverrides).toHaveBeenCalledWith(TENANT, TARGET_USER, ['mail']);
  });

  it('clears overrides when empty array is passed', async () => {
    const res = await handler(
      makePostRequest({ userId: TARGET_USER, disabledServices: [] }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    expect(mockSetUserServiceOverrides).toHaveBeenCalledWith(TENANT, TARGET_USER, []);
  });

  it('returns 400 when userId is missing from body', async () => {
    const res = await handler(
      makePostRequest({ disabledServices: ['mail'] }),
      makeContext(),
    );

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('Missing userId');
  });

  it('returns 400 when disabledServices is missing from body', async () => {
    const res = await handler(
      makePostRequest({ userId: TARGET_USER }),
      makeContext(),
    );

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('Missing disabledServices');
  });

  it('returns 400 when disabledServices contains invalid keys', async () => {
    const res = await handler(
      makePostRequest({ userId: TARGET_USER, disabledServices: ['mail', 'bogus', 'fakeService'] }),
      makeContext(),
    );

    expect(res.status).toBe(400);
    const body = res.jsonBody as { error: string };
    expect(body.error).toContain('Invalid service key');
    expect(body.error).toContain('bogus');
    expect(body.error).toContain('fakeService');
    expect(mockSetUserServiceOverrides).not.toHaveBeenCalled();
  });

  it('returns 403 when non-admin tries to modify a different user', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(
      makePostRequest({ userId: TARGET_USER, disabledServices: ['mail'] }),
      makeContext(),
    );

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('your own');
    expect(mockSetUserServiceOverrides).not.toHaveBeenCalled();
  });

  it('allows non-admin to modify their own overrides (self-service)', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(
      makePostRequest({ userId: ADMIN_USER, disabledServices: ['calendar'] }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    expect(mockSetUserServiceOverrides).toHaveBeenCalledWith(TENANT, ADMIN_USER, ['calendar']);
  });
});
