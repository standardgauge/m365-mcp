/**
 * Tests for the /api/manage/email-output-policy endpoint.
 *
 * Covers:
 *   - Authentication gating (401 when unauthenticated)
 *   - Global Admin gating on both GET and POST (403 for non-admins; nothing read or written)
 *   - GET without userId: returns the tenant-wide policy only
 *   - GET with userId: returns tenant + user policies and the resolved enforcement
 *   - POST scope=tenant: sets / clears the tenant policy, stamps the admin, logs it
 *   - POST scope=user: sets / clears a per-user policy
 *   - POST validation: bad scope, missing userId for scope=user, non-boolean enforceDraft
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

type Policy = { enforceDraft: boolean; updatedAt?: string; updatedBy?: string };
type Scope = { scope: 'tenant' } | { scope: 'user'; userId: string };

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockAuditAdminRefusal = jest.fn<(operation: string, resource?: string) => void>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetEmailOutputModePolicy = jest.fn<(tenantId: string, target: Scope) => Promise<Policy>>();
const mockGetEmailOutputModeEnforcement = jest.fn<
  (tenantId: string, userId: string) => Promise<{ enforced: boolean; enforcedBy: 'tenant' | 'user' | null; tenant: Policy; user: Policy }>
>();
const mockSetEmailOutputModePolicy = jest.fn<
  (tenantId: string, target: Scope, enforceDraft: boolean, setBy: string) => Promise<Policy>
>();
const mockLogAccess = jest.fn<(entry: Record<string, unknown>) => void>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateConsoleRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
  auditAdminRefusal: (_auth: unknown, operation: unknown, resource?: unknown) =>
    mockAuditAdminRefusal(operation as string, resource as string | undefined),
  authorizeAdmin: async (auth: unknown, operation: unknown, resource?: unknown) => {
    const isAdmin = await mockCheckGlobalAdmin((auth as AuthResult).userId);
    if (!isAdmin) mockAuditAdminRefusal(operation as string, resource as string | undefined);
    return isAdmin;
  },
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: (userId: unknown) => mockGetTenantId(userId as string),
}));

jest.mock('../services/userEmailSettings.js', () => ({
  getEmailOutputModePolicy: (tenantId: unknown, target: unknown) =>
    mockGetEmailOutputModePolicy(tenantId as string, target as Scope),
  getEmailOutputModeEnforcement: (tenantId: unknown, userId: unknown) =>
    mockGetEmailOutputModeEnforcement(tenantId as string, userId as string),
  setEmailOutputModePolicy: (tenantId: unknown, target: unknown, enforceDraft: unknown, setBy: unknown) =>
    mockSetEmailOutputModePolicy(tenantId as string, target as Scope, enforceDraft as boolean, setBy as string),
}));

jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (entry: unknown) => mockLogAccess(entry as Record<string, unknown>),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

jest.mock('../services/securityHeaders.js', () => ({
  withSecurity: (handler: unknown) => handler,
}));

import { app } from '@azure/functions';
import '../functions/admin/manageEmailOutputPolicy.js';

interface HttpRegistration {
  methods: string[];
  route: string;
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'manageEmailOutputPolicy');
if (!registration) throw new Error('manageEmailOutputPolicy handler was not registered');
const handler = registration[1].handler;

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
    query: { get: (k: string) => queryMap.get(k) ?? null, has: (k: string) => queryMap.has(k) },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

function makePostRequest(body: unknown): HttpRequest {
  return {
    method: 'POST',
    query: { get: () => null, has: () => false },
    headers: new Map<string, string>(),
    json: () => Promise.resolve(body),
  } as unknown as HttpRequest;
}

function makeContext(): InvocationContext {
  return { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
}

beforeEach(() => {
  jest.clearAllMocks();
  jest.spyOn(console, 'log').mockImplementation(() => undefined);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
  mockCheckGlobalAdmin.mockResolvedValue(true);
  mockGetEmailOutputModePolicy.mockResolvedValue({ enforceDraft: false });
  mockGetEmailOutputModeEnforcement.mockResolvedValue({
    enforced: false, enforcedBy: null, tenant: { enforceDraft: false }, user: { enforceDraft: false },
  });
  mockSetEmailOutputModePolicy.mockImplementation(async (_t, _target, enforceDraft, setBy) => ({
    enforceDraft, updatedAt: '2026-09-28T12:00:00.000Z', updatedBy: setBy,
  }));
});

afterEach(() => {
  jest.restoreAllMocks();
});

describe('registration', () => {
  it('is mounted at api/manage/email-output-policy for GET and POST', () => {
    expect(registration![1].route).toBe('api/manage/email-output-policy');
    expect(registration![1].methods).toEqual(['GET', 'POST']);
  });
});

describe('gating', () => {
  it('returns 401 when not authenticated', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);

    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(401);
    expect(mockGetEmailOutputModePolicy).not.toHaveBeenCalled();
  });

  it('returns 403 to a non-admin on GET, reading nothing', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('Global Administrator');
    expect(mockGetEmailOutputModePolicy).not.toHaveBeenCalled();
    expect(mockGetEmailOutputModeEnforcement).not.toHaveBeenCalled();
  });

  it('returns 403 to a non-admin on POST, writing nothing', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(makePostRequest({ scope: 'tenant', enforceDraft: false }), makeContext());

    expect(res.status).toBe(403);
    expect(mockSetEmailOutputModePolicy).not.toHaveBeenCalled();
    expect(mockLogAccess).not.toHaveBeenCalled();
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('set_email_output_policy', undefined);
  });
});

describe('GET /api/manage/email-output-policy', () => {
  it('returns the tenant-wide policy when no userId is given', async () => {
    mockGetEmailOutputModePolicy.mockResolvedValue({ enforceDraft: true, updatedAt: 'ts', updatedBy: ADMIN_USER });

    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({ tenant: { enforceDraft: true, updatedAt: 'ts', updatedBy: ADMIN_USER } });
    expect(mockGetEmailOutputModePolicy).toHaveBeenCalledWith(TENANT, { scope: 'tenant' });
  });

  it('returns tenant + user policies and the resolution for a userId', async () => {
    mockGetEmailOutputModeEnforcement.mockResolvedValue({
      enforced: true, enforcedBy: 'user', tenant: { enforceDraft: false }, user: { enforceDraft: true, updatedBy: ADMIN_USER },
    });

    const res = await handler(makeGetRequest({ userId: TARGET_USER }), makeContext());

    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({
      userId: TARGET_USER,
      tenant: { enforceDraft: false },
      user: { enforceDraft: true, updatedBy: ADMIN_USER },
      enforced: true,
      enforcedBy: 'user',
    });
    expect(mockGetEmailOutputModeEnforcement).toHaveBeenCalledWith(TENANT, TARGET_USER);
  });
});

describe('POST /api/manage/email-output-policy', () => {
  it('enforces draft mode tenant-wide, stamped with the admin, and logs the change', async () => {
    const res = await handler(makePostRequest({ scope: 'tenant', enforceDraft: true }), makeContext());

    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({
      ok: true, scope: 'tenant', enforceDraft: true, updatedAt: '2026-09-28T12:00:00.000Z', updatedBy: ADMIN_USER,
    });
    expect(mockSetEmailOutputModePolicy).toHaveBeenCalledWith(TENANT, { scope: 'tenant' }, true, ADMIN_USER);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      tenantId: TENANT,
      userId: ADMIN_USER,
      userEmail: 'admin@example.com',
      operation: 'set_email_output_policy',
      resource: 'tenant',
      result: 'allowed',
      source: 'http',
      before: '{"enforceDraft":false}',
      after: '{"enforceDraft":true}',
    }));
  });

  it('clears the tenant policy with enforceDraft=false', async () => {
    mockGetEmailOutputModePolicy.mockResolvedValue({ enforceDraft: true, updatedAt: 'ts', updatedBy: ADMIN_USER });
    mockSetEmailOutputModePolicy.mockResolvedValue({ enforceDraft: false, updatedAt: 'ts2', updatedBy: ADMIN_USER });
    const res = await handler(makePostRequest({ scope: 'tenant', enforceDraft: false }), makeContext());

    expect(res.status).toBe(200);
    expect(mockSetEmailOutputModePolicy).toHaveBeenCalledWith(TENANT, { scope: 'tenant' }, false, ADMIN_USER);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      resource: 'tenant',
      before: '{"enforceDraft":true}',
      after: '{"enforceDraft":false}',
    }));
  });

  it('enforces draft mode for one user', async () => {
    const res = await handler(makePostRequest({ scope: 'user', userId: TARGET_USER, enforceDraft: true }), makeContext());

    expect(res.status).toBe(200);
    expect(res.jsonBody).toMatchObject({ ok: true, scope: 'user', userId: TARGET_USER, enforceDraft: true });
    expect(mockSetEmailOutputModePolicy).toHaveBeenCalledWith(TENANT, { scope: 'user', userId: TARGET_USER }, true, ADMIN_USER);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ resource: `user:${TARGET_USER}` }));
  });

  it.each([
    ['bad scope', { scope: 'everyone', enforceDraft: true }, 'scope must be'],
    ['missing scope', { enforceDraft: true }, 'scope must be'],
    ['scope=user without userId', { scope: 'user', enforceDraft: true }, 'Missing userId'],
    ['scope=user with empty userId', { scope: 'user', userId: '', enforceDraft: true }, 'Missing userId'],
    ['string enforceDraft', { scope: 'tenant', enforceDraft: 'true' }, 'enforceDraft'],
    ['missing enforceDraft', { scope: 'tenant' }, 'enforceDraft'],
  ])('rejects %s with 400 and writes nothing', async (_label, body, errorText) => {
    const res = await handler(makePostRequest(body), makeContext());

    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain(errorText);
    expect(mockSetEmailOutputModePolicy).not.toHaveBeenCalled();
    expect(mockLogAccess).not.toHaveBeenCalled();
  });

  it('returns 500 and does not log when the write fails', async () => {
    mockSetEmailOutputModePolicy.mockRejectedValue(new Error('storage down'));

    const res = await handler(makePostRequest({ scope: 'tenant', enforceDraft: true }), makeContext());

    expect(res.status).toBe(500);
    expect((res.jsonBody as { error: string }).error).toBe('storage down');
    expect(mockLogAccess).not.toHaveBeenCalled();
  });
});
