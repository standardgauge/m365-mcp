/**
 * Integration tests for the withPolicyEnforcement wrapper.
 *
 * All external dependencies (auth, tenant settings, deny list, token cache)
 * are mocked so these tests run without cloud infrastructure. The focus is
 * on verifying the wrapper's enforcement chain:
 *   1. Authentication (401 when invalid/expired)
 *   2. Service enablement (403 when service not in enabledServices)
 *   3. Allowed sites — SharePoint (403 when siteId not in allowedSites)
 *   4. Deny list (403 when path is denied)
 *   5. Inner handler invocation when all checks pass
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockGetEnabledServices = jest.fn<(tenantId: string) => Promise<string[]>>();
const mockGetAllowedSites = jest.fn<(tenantId: string) => Promise<Array<{ id: string; name: string }>>>();
const mockIsPathDenied = jest.fn<(tenantId: string, userId: string, service: string, path: string) => Promise<boolean>>();
const mockIsServiceDisabledForUser = jest.fn<(tenantId: string, userId: string, service: string) => Promise<boolean>>();

const mockIsMailIndexingDisabled = jest.fn<(tenantId: string, userId: string) => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();

// ── Wire up mocks before importing the module under test ─────────────────────

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (...args: unknown[]) => mockAuthenticateRequest(args[0] as HttpRequest),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: (...args: unknown[]) => mockGetEnabledServices(args[0] as string),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: (...args: unknown[]) => mockGetAllowedSites(args[0] as string),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  isServiceDisabledForUser: (...args: unknown[]) =>
    mockIsServiceDisabledForUser(args[0] as string, args[1] as string, args[2] as string),
}));

jest.mock('../services/denyList.js', () => ({
  isPathDenied: (...args: unknown[]) =>
    mockIsPathDenied(args[0] as string, args[1] as string, args[2] as string, args[3] as string),
}));

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: (...args: unknown[]) =>
    mockIsMailIndexingDisabled(args[0] as string, args[1] as string),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: (...args: unknown[]) => mockGetTenantId(args[0] as string),
  getTenantIdFromSession: (session: { tenantId?: string }) => {
    if (!session?.tenantId) throw new Error('No tenantId in session');
    return session.tenantId;
  },
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import {
  withPolicyEnforcement,
  checkServiceEnabled,
  checkAllowedSite,
  checkDenyList,
  checkMailIndexing,
} from '../services/policyEnforcement.js';
import type { PolicyHandler } from '../services/policyEnforcement.js';
import { SessionStoreUnavailableError } from '../services/sessionStoreError.js';

// ── Test helpers ─────────────────────────────────────────────────────────────

const TENANT = 'test-tenant-id';
const USER = 'user-abc';

const FAKE_AUTH: AuthResult = {
  userId: USER,
  session: {
    userId: USER,
    homeAccountId: 'home-abc',
    displayName: 'Test User',
    email: 'test@example.com',
    tenantId: TENANT,
    accessToken: 'fake-access-token',
    expiresAt: Date.now() + 3_600_000,
    sessionToken: 'fake-session-token',
    sessionCreatedAt: Date.now(),
  },
};

/** Build a minimal mock HttpRequest with optional query params and route params. */
function makeRequest(opts?: {
  query?: Record<string, string>;
  params?: Record<string, string>;
}): HttpRequest {
  const queryMap = new Map(Object.entries(opts?.query ?? {}));
  return {
    query: {
      get: (key: string) => queryMap.get(key) ?? null,
      has: (key: string) => queryMap.has(key),
    },
    params: opts?.params ?? {},
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

/** Build a minimal mock InvocationContext. */
function makeContext(): InvocationContext {
  return {
    error: jest.fn(),
    warn: jest.fn(),
    log: jest.fn(),
    trace: jest.fn(),
  } as unknown as InvocationContext;
}

/** Set up "happy path" mocks so all policy checks pass. */
function setupPassingMocks(enabledServices: string[] = ['mail', 'sharepoint']) {
  mockAuthenticateRequest.mockResolvedValue(FAKE_AUTH);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockGetEnabledServices.mockResolvedValue(enabledServices);
  mockGetAllowedSites.mockResolvedValue([]);
  mockIsPathDenied.mockResolvedValue(false);
  mockIsServiceDisabledForUser.mockResolvedValue(false);

  mockIsMailIndexingDisabled.mockResolvedValue(false);
}

// ── Test suite ───────────────────────────────────────────────────────────────

beforeEach(() => {
  jest.clearAllMocks();
});

// ──────────────────────────────────────────────────────────────────────────────
// Service enablement
// ──────────────────────────────────────────────────────────────────────────────

describe('Authentication', () => {
  it('returns 503, not 401, when the session store cannot be reached', async () => {
    setupPassingMocks();
    mockAuthenticateRequest.mockRejectedValue(new SessionStoreUnavailableError(new Error('ServerBusy')));

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const res = await withPolicyEnforcement('mail', handler)(makeRequest(), makeContext());

    expect(res.status).toBe(503);
    expect(res.headers).toMatchObject({ 'Retry-After': '5' });
    expect((res.jsonBody as { error: string }).error).toBe('Session store unavailable, retry shortly');
    expect(handler).not.toHaveBeenCalled();
  });
});

describe('Service enablement', () => {
  it('1. returns 403 when the service is disabled for the tenant', async () => {
    setupPassingMocks(['mail']); // sharepoint NOT enabled

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('sharepoint', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('sharepoint');
    expect((res.jsonBody as { error: string }).error).toContain('not enabled');
    expect(handler).not.toHaveBeenCalled();
  });

  it('2. passes through to the inner handler when the service is enabled', async () => {
    setupPassingMocks(['mail', 'sharepoint']);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({
      status: 200,
      jsonBody: { ok: true },
    });
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(200);
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it('3. treats an empty enabledServices list as blocking all services (strict enforcement)', async () => {
    // When getEnabledServices returns [] (admin explicitly disabled everything),
    // no service passes through — this is strict, not permissive.
    // The "permissive default" is handled by getEnabledServices returning
    // DEFAULT_SERVICES when no config row exists.
    setupPassingMocks([]);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(403);
    expect(handler).not.toHaveBeenCalled();
  });
});

// ──────────────────────────────────────────────────────────────────────────────
// Per-user service overrides
// ──────────────────────────────────────────────────────────────────────────────

describe('Per-user service overrides', () => {
  it('returns 403 when the service is disabled for the specific user', async () => {
    setupPassingMocks(['mail', 'sharepoint']);
    mockIsServiceDisabledForUser.mockResolvedValue(true);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('disabled for your account');
    expect(handler).not.toHaveBeenCalled();
  });

  it('passes through when user has no overrides', async () => {
    setupPassingMocks(['mail', 'sharepoint']);
    mockIsServiceDisabledForUser.mockResolvedValue(false);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({
      status: 200,
      jsonBody: { ok: true },
    });
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(200);
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it('tenant-level disable takes precedence over user overrides (tenant wins)', async () => {
    // Service disabled at tenant level — user override is irrelevant
    setupPassingMocks(['sharepoint']); // mail NOT enabled at tenant level
    mockIsServiceDisabledForUser.mockResolvedValue(false); // user has no override

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('not enabled for this tenant');
    // User override check should NOT have been called — tenant check short-circuits
    expect(mockIsServiceDisabledForUser).not.toHaveBeenCalled();
    expect(handler).not.toHaveBeenCalled();
  });
});

// ──────────────────────────────────────────────────────────────────────────────
// Mail indexing disabled
// ──────────────────────────────────────────────────────────────────────────────

describe('Mail indexing disabled', () => {
  it('returns 403 when mail indexing is disabled for a mail operation', async () => {
    setupPassingMocks(['mail', 'sharepoint']);
    mockIsMailIndexingDisabled.mockResolvedValue(true);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('disabled for your account');
    expect(handler).not.toHaveBeenCalled();
  });

  it('does NOT block a non-mail service when mail indexing flag is set', async () => {
    setupPassingMocks(['sharepoint']);
    mockIsMailIndexingDisabled.mockResolvedValue(true);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({
      status: 200,
      jsonBody: { ok: true },
    });
    const wrapped = withPolicyEnforcement('sharepoint', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(200);
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it('short-circuits before deny-list when mail indexing is disabled', async () => {
    setupPassingMocks(['mail']);
    mockIsMailIndexingDisabled.mockResolvedValue(true);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('mail', handler, {
      getDenyListPaths: () => ['/Inbox'],
    });

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(403);
    expect(mockIsPathDenied).not.toHaveBeenCalled();
    expect(handler).not.toHaveBeenCalled();
  });

  it('passes through when mail indexing is NOT disabled', async () => {
    setupPassingMocks(['mail']);
    mockIsMailIndexingDisabled.mockResolvedValue(false);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({
      status: 200,
      jsonBody: { messages: [] },
    });
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(200);
    expect(handler).toHaveBeenCalledTimes(1);
  });
});

// ──────────────────────────────────────────────────────────────────────────────
// Allowed sites (SharePoint)
// ──────────────────────────────────────────────────────────────────────────────

describe('Allowed sites (SharePoint)', () => {
  it('4. returns 403 when siteId is NOT in allowedSites', async () => {
    setupPassingMocks(['sharepoint']);
    mockGetAllowedSites.mockResolvedValue([{ id: 'site-aaa', name: 'Allowed Site' }]);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('sharepoint', handler, {
      getSiteId: (req) => req.query.get('siteId') ?? undefined,
    });

    const res = await wrapped(
      makeRequest({ query: { siteId: 'site-zzz' } }),
      makeContext(),
    );

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('site-zzz');
    expect((res.jsonBody as { error: string }).error).toContain('not in the allowed sites');
    expect(handler).not.toHaveBeenCalled();
  });

  it('5. passes through when siteId IS in allowedSites', async () => {
    setupPassingMocks(['sharepoint']);
    mockGetAllowedSites.mockResolvedValue([{ id: 'site-aaa', name: 'Allowed Site' }]);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({
      status: 200,
      jsonBody: { files: [] },
    });
    const wrapped = withPolicyEnforcement('sharepoint', handler, {
      getSiteId: (req) => req.query.get('siteId') ?? undefined,
    });

    const res = await wrapped(
      makeRequest({ query: { siteId: 'site-aaa' } }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it('6. allows all sites when allowedSites is empty (permissive default)', async () => {
    setupPassingMocks(['sharepoint']);
    mockGetAllowedSites.mockResolvedValue([]); // no restriction

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({
      status: 200,
      jsonBody: { files: [] },
    });
    const wrapped = withPolicyEnforcement('sharepoint', handler, {
      getSiteId: (req) => req.query.get('siteId') ?? undefined,
    });

    const res = await wrapped(
      makeRequest({ query: { siteId: 'any-site-id' } }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    expect(handler).toHaveBeenCalledTimes(1);
  });
});

// ──────────────────────────────────────────────────────────────────────────────
// Deny list
// ──────────────────────────────────────────────────────────────────────────────

describe('Deny list', () => {
  it('7. returns 403 when a path is denied', async () => {
    setupPassingMocks(['sharepoint']);
    mockIsPathDenied.mockResolvedValue(true);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('sharepoint', handler, {
      getDenyListPaths: () => ['/sites/HR/Payroll'],
    });

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('deny list');
    expect(handler).not.toHaveBeenCalled();
  });

  it('8. passes through when the path is not denied', async () => {
    setupPassingMocks(['sharepoint']);
    mockIsPathDenied.mockResolvedValue(false);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({
      status: 200,
      jsonBody: { data: 'ok' },
    });
    const wrapped = withPolicyEnforcement('sharepoint', handler, {
      getDenyListPaths: () => ['/sites/Engineering/Docs'],
    });

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(200);
    expect(handler).toHaveBeenCalledTimes(1);
  });

  it('9. child paths of denied parent paths are also blocked (via isPathDenied)', async () => {
    setupPassingMocks(['sharepoint']);
    // isPathDenied handles the prefix matching internally — we just verify
    // the wrapper delegates properly and returns 403 when isPathDenied says true.
    mockIsPathDenied.mockResolvedValue(true);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('sharepoint', handler, {
      getDenyListPaths: () => ['/sites/HR/Payroll/Q1/salaries.xlsx'],
    });

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(403);
    expect(handler).not.toHaveBeenCalled();
    // Verify the child path was passed to isPathDenied
    expect(mockIsPathDenied).toHaveBeenCalledWith(
      TENANT,
      USER,
      'sharepoint',
      '/sites/HR/Payroll/Q1/salaries.xlsx',
    );
  });
});

// ──────────────────────────────────────────────────────────────────────────────
// Authentication
// ──────────────────────────────────────────────────────────────────────────────

describe('Authentication', () => {
  it('10. returns 401 when no valid session exists', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(401);
    expect((res.jsonBody as { error: string }).error).toContain('Authentication required');
    expect(handler).not.toHaveBeenCalled();
  });

  it('11. returns 401 when inner handler throws "Re-authentication required"', async () => {
    // authenticateRequest succeeds initially, but the handler throws when
    // it discovers the session is stale (e.g. token refresh failed).
    setupPassingMocks(['mail']);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockRejectedValue(
      new Error('Token refresh failed for user user-abc. Re-authentication required.'),
    );
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(401);
    expect((res.jsonBody as { error: string }).error).toContain('Re-authentication required');
  });
});

// ──────────────────────────────────────────────────────────────────────────────
// Combined — full pass-through
// ──────────────────────────────────────────────────────────────────────────────

describe('Combined — full pass-through', () => {
  it('12. request passes all checks and inner handler response is returned', async () => {
    setupPassingMocks(['sharepoint']);
    mockGetAllowedSites.mockResolvedValue([{ id: 'site-aaa', name: 'Project Site' }]);
    mockIsPathDenied.mockResolvedValue(false);

    const expectedBody = { files: [{ name: 'report.docx' }] };
    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({
      status: 200,
      jsonBody: expectedBody,
    });

    const wrapped = withPolicyEnforcement('sharepoint', handler, {
      getSiteId: (req) => req.query.get('siteId') ?? undefined,
      getDenyListPaths: () => ['/sites/Project/Documents/report.docx'],
    });

    const req = makeRequest({ query: { siteId: 'site-aaa' } });
    const ctx = makeContext();
    const res = await wrapped(req, ctx);

    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual(expectedBody);
    expect(handler).toHaveBeenCalledTimes(1);
    // Verify the handler received the auth result
    expect(handler).toHaveBeenCalledWith(req, ctx, FAKE_AUTH);
  });
});

// ──────────────────────────────────────────────────────────────────────────────
// Error handling
// ──────────────────────────────────────────────────────────────────────────────

describe('Error handling', () => {
  it('returns 500 for unexpected handler errors', async () => {
    setupPassingMocks(['mail']);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockRejectedValue(
      new Error('Unexpected Graph API failure'),
    );
    const wrapped = withPolicyEnforcement('mail', handler);

    const ctx = makeContext();
    const res = await wrapped(makeRequest(), ctx);

    expect(res.status).toBe(500);
    expect((res.jsonBody as { error: string }).error).toBe('Internal server error');
    expect(ctx.error).toHaveBeenCalled();
  });

  it('returns 500 when session has no tenantId', async () => {
    // Session-aware path: getTenantIdFromSession reads session.tenantId directly
    const authNoTenant = {
      userId: FAKE_AUTH.userId,
      session: { ...FAKE_AUTH.session, tenantId: '' },
    };
    mockAuthenticateRequest.mockResolvedValue(authNoTenant);
    mockGetEnabledServices.mockResolvedValue(['mail']);

    const handler: PolicyHandler = jest.fn<PolicyHandler>().mockResolvedValue({ status: 200 });
    const wrapped = withPolicyEnforcement('mail', handler);

    const res = await wrapped(makeRequest(), makeContext());

    expect(res.status).toBe(500);
    expect(handler).not.toHaveBeenCalled();
  });
});

// ──────────────────────────────────────────────────────────────────────────────
// Lower-level helpers — direct unit tests
// ──────────────────────────────────────────────────────────────────────────────

describe('checkServiceEnabled', () => {
  it('returns null when the service is in the enabled list', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockGetEnabledServices.mockResolvedValue(['mail', 'sharepoint', 'calendar']);
    mockIsServiceDisabledForUser.mockResolvedValue(false);

    const result = await checkServiceEnabled(USER, 'calendar');
    expect(result).toBeNull();
  });

  it('returns a 403 violation when the service is not in the enabled list', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockGetEnabledServices.mockResolvedValue(['mail']);

    const result = await checkServiceEnabled(USER, 'onedrive');
    expect(result).not.toBeNull();
    expect(result!.status).toBe(403);
    expect(result!.error).toContain('onedrive');
  });

  it('returns a 403 violation when the service is disabled for the user', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockGetEnabledServices.mockResolvedValue(['mail', 'sharepoint', 'calendar']);
    mockIsServiceDisabledForUser.mockResolvedValue(true);

    const result = await checkServiceEnabled(USER, 'calendar');
    expect(result).not.toBeNull();
    expect(result!.status).toBe(403);
    expect(result!.error).toContain('disabled for your account');
  });

  it('does not check user overrides when tenant disables the service', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockGetEnabledServices.mockResolvedValue(['mail']); // calendar not enabled

    const result = await checkServiceEnabled(USER, 'calendar');
    expect(result).not.toBeNull();
    expect(result!.error).toContain('not enabled for this tenant');
    expect(mockIsServiceDisabledForUser).not.toHaveBeenCalled();
  });
});

describe('checkAllowedSite', () => {
  it('returns null when allowedSites is empty (permissive)', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockGetAllowedSites.mockResolvedValue([]);

    const result = await checkAllowedSite(USER, 'any-site-id');
    expect(result).toBeNull();
  });

  it('returns null when siteId is in the allowed list', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockGetAllowedSites.mockResolvedValue([{ id: 'site-aaa', name: 'A' }]);

    const result = await checkAllowedSite(USER, 'site-aaa');
    expect(result).toBeNull();
  });

  it('returns a 403 violation when siteId is not in the allowed list', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockGetAllowedSites.mockResolvedValue([{ id: 'site-aaa', name: 'A' }]);

    const result = await checkAllowedSite(USER, 'site-zzz');
    expect(result).not.toBeNull();
    expect(result!.status).toBe(403);
    expect(result!.error).toContain('site-zzz');
  });
});

describe('checkDenyList', () => {
  it('returns null when path is not denied', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockIsPathDenied.mockResolvedValue(false);

    const result = await checkDenyList(USER, 'sharepoint', '/sites/Public');
    expect(result).toBeNull();
  });

  it('returns a 403 violation when path is denied', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockIsPathDenied.mockResolvedValue(true);

    const result = await checkDenyList(USER, 'sharepoint', '/sites/HR/Payroll');
    expect(result).not.toBeNull();
    expect(result!.status).toBe(403);
    expect(result!.error).toContain('deny list');
  });
});

describe('checkMailIndexing', () => {
  it('returns null when mail indexing is not disabled', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockIsMailIndexingDisabled.mockResolvedValue(false);

    const result = await checkMailIndexing(USER);
    expect(result).toBeNull();
  });

  it('returns a 403 violation when mail indexing is disabled', async () => {
    mockGetTenantId.mockResolvedValue(TENANT);
    mockIsMailIndexingDisabled.mockResolvedValue(true);

    const result = await checkMailIndexing(USER);
    expect(result).not.toBeNull();
    expect(result!.status).toBe(403);
    expect(result!.error).toContain('disabled for your account');
  });
});
