/**
 * Tests for the /api/manage/tenant-users admin endpoint.
 *
 * Covers:
 *   - Authentication gating (401 / 403)
 *   - Default filter: enabled + Member + licensed
 *   - includeAll=true bypasses filters
 *   - Sort order (displayName ascending)
 *   - "installed" badge cross-referenced with active sessions
 *   - Pagination via @odata.nextLink
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockGetValidAccessTokenForSession = jest.fn<(s: unknown) => Promise<string>>();
const mockListActiveSessions = jest.fn<() => Promise<Array<{ userId: string; email: string; displayName: string; expiresAt: number; sessionCreatedAt: number }>>>();
const SEVEN_DAYS_MS = 7 * 24 * 60 * 60 * 1000;

interface FakeGraphResponse { value?: unknown[]; '@odata.nextLink'?: string }
const mockGraphGet = jest.fn<() => Promise<FakeGraphResponse>>();

const fakeGraphChain = {
  header: () => fakeGraphChain,
  select: () => fakeGraphChain,
  top: () => fakeGraphChain,
  get: () => mockGraphGet(),
};

// ── Wire up mocks ────────────────────────────────────────────────────────────

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
}));

jest.mock('../services/tokenCache.js', () => ({
  getValidAccessTokenForSession: (s: unknown) => mockGetValidAccessTokenForSession(s),
  listActiveSessions: () => mockListActiveSessions(),
  SESSION_TTL_MS: SEVEN_DAYS_MS,
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => ({ api: () => fakeGraphChain }),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import { app } from '@azure/functions';
import '../functions/admin/getTenantUsers.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'getTenantUsers');
if (!registration) throw new Error('getTenantUsers handler was not registered');
const handler = registration[1].handler;

// ── Test fixtures ────────────────────────────────────────────────────────────

const TENANT = 'tenant-1';
const ADMIN_USER = 'admin-user';

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

const LICENSED_MEMBER = {
  id: 'u-1',
  displayName: 'Alice',
  userPrincipalName: 'alice@example.com',
  mail: 'alice@example.com',
  accountEnabled: true,
  userType: 'Member',
  assignedLicenses: [{ skuId: 'sku-1' }],
};

const DISABLED_MEMBER = {
  id: 'u-2',
  displayName: 'Bob',
  userPrincipalName: 'bob@example.com',
  mail: 'bob@example.com',
  accountEnabled: false,
  userType: 'Member',
  assignedLicenses: [{ skuId: 'sku-1' }],
};

const UNLICENSED_MEMBER = {
  id: 'u-3',
  displayName: 'Carol',
  userPrincipalName: 'carol@example.com',
  mail: 'carol@example.com',
  accountEnabled: true,
  userType: 'Member',
  assignedLicenses: [],
};

const GUEST = {
  id: 'u-4',
  displayName: 'Dan (External)',
  userPrincipalName: 'dan_external#EXT#@example.com',
  mail: 'dan@external.com',
  accountEnabled: true,
  userType: 'Guest',
  assignedLicenses: [{ skuId: 'sku-1' }],
};

function makeRequest(query?: Record<string, string>): HttpRequest {
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
  // Each test runs with a fresh tenant ID to bypass the in-memory cache.
  ADMIN_AUTH.session.tenantId = `tenant-${Math.random().toString(36).slice(2)}`;
  mockGetValidAccessTokenForSession.mockResolvedValue('fake-token');
  mockListActiveSessions.mockResolvedValue([]);
  mockGraphGet.mockReset();
});

describe('Authentication', () => {
  it('returns 401 when not authenticated', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);
    const res = await handler(makeRequest(), makeContext());
    expect(res.status).toBe(401);
  });

  it('returns 403 when caller is not Global Admin', async () => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(false);
    const res = await handler(makeRequest(), makeContext());
    expect(res.status).toBe(403);
  });
});

describe('Default filtering', () => {
  beforeEach(() => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(true);
  });

  it('keeps only enabled licensed Members by default', async () => {
    mockGraphGet.mockResolvedValueOnce({
      value: [LICENSED_MEMBER, DISABLED_MEMBER, UNLICENSED_MEMBER, GUEST],
    });

    const res = await handler(makeRequest(), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { users: Array<{ userId: string }>; totalCount: number; filteredCount: number };
    expect(body.totalCount).toBe(4);
    expect(body.filteredCount).toBe(1);
    expect(body.users.map((u) => u.userId)).toEqual(['u-1']);
  });

  it('returns all users when includeAll=true', async () => {
    mockGraphGet.mockResolvedValueOnce({
      value: [LICENSED_MEMBER, DISABLED_MEMBER, UNLICENSED_MEMBER, GUEST],
    });

    const res = await handler(makeRequest({ includeAll: 'true' }), makeContext());

    expect(res.status).toBe(200);
    const body = res.jsonBody as { users: Array<{ userId: string }> };
    expect(body.users.map((u) => u.userId).sort()).toEqual(['u-1', 'u-2', 'u-3', 'u-4']);
  });

  it('sorts users by displayName ascending', async () => {
    mockGraphGet.mockResolvedValueOnce({
      value: [
        { ...LICENSED_MEMBER, id: 'u-z', displayName: 'Zach' },
        { ...LICENSED_MEMBER, id: 'u-a', displayName: 'Alice' },
        { ...LICENSED_MEMBER, id: 'u-m', displayName: 'Mike' },
      ],
    });

    const res = await handler(makeRequest(), makeContext());

    const body = res.jsonBody as { users: Array<{ displayName: string }> };
    expect(body.users.map((u) => u.displayName)).toEqual(['Alice', 'Mike', 'Zach']);
  });
});

describe('Installed badge', () => {
  beforeEach(() => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(true);
  });

  it('marks users with active sessions as installed', async () => {
    mockGraphGet.mockResolvedValueOnce({
      value: [LICENSED_MEMBER, { ...LICENSED_MEMBER, id: 'u-9', displayName: 'Nine' }],
    });
    mockListActiveSessions.mockResolvedValue([
      { userId: 'u-9', email: 'nine@example.com', displayName: 'Nine', expiresAt: Date.now() + 1000, sessionCreatedAt: Date.now() },
    ]);

    const res = await handler(makeRequest(), makeContext());

    const body = res.jsonBody as { users: Array<{ userId: string; installed: boolean }> };
    const byId = Object.fromEntries(body.users.map((u) => [u.userId, u.installed]));
    expect(byId['u-9']).toBe(true);
    expect(byId['u-1']).toBe(false);
  });

  it('dedups when one user has multiple sessions', async () => {
    mockGraphGet.mockResolvedValueOnce({ value: [LICENSED_MEMBER] });
    mockListActiveSessions.mockResolvedValue([
      { userId: 'u-1', email: 'alice@example.com', displayName: 'Alice', expiresAt: 1, sessionCreatedAt: Date.now() - 1000 },
      { userId: 'u-1', email: 'alice@example.com', displayName: 'Alice', expiresAt: 2, sessionCreatedAt: Date.now() - 500 },
      { userId: 'u-1', email: 'alice@example.com', displayName: 'Alice', expiresAt: 3, sessionCreatedAt: Date.now() },
    ]);

    const res = await handler(makeRequest(), makeContext());

    const body = res.jsonBody as { users: Array<{ userId: string; installed: boolean }> };
    expect(body.users).toHaveLength(1);
    expect(body.users[0].installed).toBe(true);
  });

  it('does NOT mark stale sessions (past idle TTL) as installed', async () => {
    mockGraphGet.mockResolvedValueOnce({ value: [LICENSED_MEMBER] });
    mockListActiveSessions.mockResolvedValue([
      // sessionCreatedAt is older than 7-day SESSION_TTL_MS
      { userId: 'u-1', email: 'alice@example.com', displayName: 'Alice', expiresAt: Date.now() + 1000, sessionCreatedAt: Date.now() - SEVEN_DAYS_MS - 60_000 },
    ]);

    const res = await handler(makeRequest(), makeContext());

    const body = res.jsonBody as { users: Array<{ userId: string; installed: boolean }> };
    expect(body.users[0].installed).toBe(false);
  });

  it('does NOT mark legacy sessions with sessionCreatedAt=0 as installed', async () => {
    mockGraphGet.mockResolvedValueOnce({ value: [LICENSED_MEMBER] });
    mockListActiveSessions.mockResolvedValue([
      { userId: 'u-1', email: 'alice@example.com', displayName: 'Alice', expiresAt: Date.now() + 1000, sessionCreatedAt: 0 },
    ]);

    const res = await handler(makeRequest(), makeContext());

    const body = res.jsonBody as { users: Array<{ userId: string; installed: boolean }> };
    expect(body.users[0].installed).toBe(false);
  });
});

describe('Pagination', () => {
  beforeEach(() => {
    mockAuthenticateRequest.mockResolvedValue(ADMIN_AUTH);
    mockCheckGlobalAdmin.mockResolvedValue(true);
  });

  it('halts when the same @odata.nextLink is returned twice ( cycle guard)', async () => {
    const cyclic = 'https://graph/page-loop';
    mockGraphGet
      .mockResolvedValueOnce({ value: [LICENSED_MEMBER], '@odata.nextLink': cyclic })
      .mockResolvedValue({ value: [{ ...LICENSED_MEMBER, id: 'u-dup' }], '@odata.nextLink': cyclic });

    const res = await handler(makeRequest(), makeContext());

    expect(res.status).toBe(200);
    // First page processed, then one follow-up to cyclic link, then halt — 2 calls total.
    expect(mockGraphGet).toHaveBeenCalledTimes(2);
  });

  it('halts at MAX_PAGES even when nextLinks keep advancing', async () => {
    // Always return a fresh, non-cyclic nextLink so only the page cap can stop the loop.
    let counter = 0;
    mockGraphGet.mockImplementation(() =>
      Promise.resolve({
        value: [{ ...LICENSED_MEMBER, id: `u-${counter}` }],
        '@odata.nextLink': `https://graph/page-${counter++}`,
      }),
    );

    const res = await handler(makeRequest(), makeContext());

    expect(res.status).toBe(200);
    // Initial fetch + MAX_PAGES (50) follow-ups = 51 calls before halt.
    expect(mockGraphGet).toHaveBeenCalledTimes(51);
  });

  it('follows @odata.nextLink to fetch all pages', async () => {
    mockGraphGet
      .mockResolvedValueOnce({
        value: [LICENSED_MEMBER],
        '@odata.nextLink': 'https://graph/page2',
      })
      .mockResolvedValueOnce({
        value: [{ ...LICENSED_MEMBER, id: 'u-2b', displayName: 'Beth' }],
        '@odata.nextLink': 'https://graph/page3',
      })
      .mockResolvedValueOnce({
        value: [{ ...LICENSED_MEMBER, id: 'u-3c', displayName: 'Carl' }],
      });

    const res = await handler(makeRequest(), makeContext());

    const body = res.jsonBody as { users: Array<{ userId: string }>; totalCount: number };
    expect(body.totalCount).toBe(3);
    expect(body.users.map((u) => u.userId).sort()).toEqual(['u-1', 'u-2b', 'u-3c']);
    expect(mockGraphGet).toHaveBeenCalledTimes(3);
  });
});
