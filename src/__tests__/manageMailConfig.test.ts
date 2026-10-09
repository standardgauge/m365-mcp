/**
 * Tests for the /api/manage/mail-config admin endpoint.
 *
 * Covers:
 *   - 401 when unauthenticated
 *   - 403 when caller is not a Global Admin (GET and POST)
 *   - GET: returns mail config for target user
 *   - GET: 400 when userId query param missing
 *   - POST: sets disable_mail_indexing flag
 *   - POST: clears disable_mail_indexing flag (false)
 *   - POST: 400 when body missing userId or disable_mail_indexing
 *   - POST: 400 when disable_mail_indexing is not a boolean
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';
import type { UserMailConfig } from '../services/userMailConfig.js';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockAuditAdminRefusal = jest.fn<(operation: string) => void>();
const mockLogAccess = jest.fn<(entry: Record<string, unknown>) => void>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetMailConfig = jest.fn<(tenantId: string, userId: string) => Promise<UserMailConfig>>();
const mockSetMailConfig = jest.fn<
  (tenantId: string, userId: string, patch: Partial<UserMailConfig>, setBy: string) => Promise<void>
>();

// ── Wire up mocks ────────────────────────────────────────────────────────────

jest.mock('../services/authMiddleware.js', () => ({
  authenticateConsoleRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
  authorizeAdmin: async (auth: unknown, operation: unknown) => {
    const isAdmin = await mockCheckGlobalAdmin((auth as AuthResult).userId);
    if (!isAdmin) mockAuditAdminRefusal(operation as string);
    return isAdmin;
  },
}));

jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (entry: unknown) => mockLogAccess(entry as Record<string, unknown>),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: (userId: unknown) => mockGetTenantId(userId as string),
}));

jest.mock('../services/userMailConfig.js', () => ({
  getMailConfig: (tenantId: unknown, userId: unknown) =>
    mockGetMailConfig(tenantId as string, userId as string),
  setMailConfig: (tenantId: unknown, userId: unknown, patch: unknown, setBy: unknown) =>
    mockSetMailConfig(tenantId as string, userId as string, patch as Partial<UserMailConfig>, setBy as string),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import { app } from '@azure/functions';
import '../functions/admin/manageMailConfig.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'manageMailConfig');
if (!registration) throw new Error('manageMailConfig handler was not registered');
const handler = registration[1].handler;

// ── Fixtures ────────────────────────────────────────────────────────────────

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

  it('returns 403 when caller is not a Global Admin (GET)', async () => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(makeGetRequest({ userId: TARGET_USER }), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('Global Administrator');
  });

  it('returns 403 when caller is not a Global Admin (POST)', async () => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(
      makePostRequest({ userId: TARGET_USER, disable_mail_indexing: true }),
      makeContext(),
    );

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('Global Administrator');
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('policy.mail_config.set');
    expect(mockSetMailConfig).not.toHaveBeenCalled();
  });
});

describe('GET /api/manage/mail-config', () => {
  beforeEach(() => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(true);
  });

  it('returns mail config for target user', async () => {
    mockGetMailConfig.mockResolvedValue({
      disable_mail_indexing: true,
      updated_at: '2026-01-01T00:00:00.000Z',
      updated_by: ADMIN_USER,
    });

    const res = await handler(makeGetRequest({ userId: TARGET_USER }), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { userId: string; disable_mail_indexing: boolean };
    expect(body.userId).toBe(TARGET_USER);
    expect(body.disable_mail_indexing).toBe(true);
    expect(mockGetMailConfig).toHaveBeenCalledWith(TENANT, TARGET_USER);
  });

  it('returns default config (flag=false) when no config exists', async () => {
    mockGetMailConfig.mockResolvedValue({ disable_mail_indexing: false });

    const res = await handler(makeGetRequest({ userId: TARGET_USER }), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { disable_mail_indexing: boolean };
    expect(body.disable_mail_indexing).toBe(false);
  });

  it('returns 400 when userId query param is missing', async () => {
    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('Missing userId');
  });
});

describe('POST /api/manage/mail-config', () => {
  beforeEach(() => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockSetMailConfig.mockResolvedValue(undefined);
    mockGetMailConfig.mockResolvedValue({ disable_mail_indexing: false });
  });

  it('audits the change with the flag before and after', async () => {
    await handler(makePostRequest({ userId: TARGET_USER, disable_mail_indexing: true }), makeContext());

    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      tenantId: TENANT,
      userId: ADMIN_USER,
      operation: 'policy.mail_config.set',
      resource: `user:${TARGET_USER}`,
      result: 'allowed',
      before: '{"disable_mail_indexing":false}',
      after: '{"disable_mail_indexing":true}',
    }));
  });

  it('sets disable_mail_indexing to true', async () => {
    const res = await handler(
      makePostRequest({ userId: TARGET_USER, disable_mail_indexing: true }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    const body = res.jsonBody as { ok: boolean; userId: string; disable_mail_indexing: boolean };
    expect(body.ok).toBe(true);
    expect(body.userId).toBe(TARGET_USER);
    expect(body.disable_mail_indexing).toBe(true);
    expect(mockSetMailConfig).toHaveBeenCalledWith(
      TENANT,
      TARGET_USER,
      { disable_mail_indexing: true },
      ADMIN_USER,
    );
  });

  it('clears disable_mail_indexing (set to false)', async () => {
    const res = await handler(
      makePostRequest({ userId: TARGET_USER, disable_mail_indexing: false }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    const body = res.jsonBody as { ok: boolean; disable_mail_indexing: boolean };
    expect(body.ok).toBe(true);
    expect(body.disable_mail_indexing).toBe(false);
    expect(mockSetMailConfig).toHaveBeenCalledWith(
      TENANT,
      TARGET_USER,
      { disable_mail_indexing: false },
      ADMIN_USER,
    );
  });

  it('returns 400 when userId is missing from body', async () => {
    const res = await handler(
      makePostRequest({ disable_mail_indexing: true }),
      makeContext(),
    );

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('Missing userId');
    expect(mockSetMailConfig).not.toHaveBeenCalled();
  });

  it('returns 400 when disable_mail_indexing is missing from body', async () => {
    const res = await handler(
      makePostRequest({ userId: TARGET_USER }),
      makeContext(),
    );

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('disable_mail_indexing');
    expect(mockSetMailConfig).not.toHaveBeenCalled();
  });

  it('returns 400 when disable_mail_indexing is not a boolean (string)', async () => {
    const res = await handler(
      makePostRequest({ userId: TARGET_USER, disable_mail_indexing: 'yes' }),
      makeContext(),
    );

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('boolean');
    expect(mockSetMailConfig).not.toHaveBeenCalled();
  });
});
