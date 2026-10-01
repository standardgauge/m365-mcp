/**
 * Tests for the /api/manage/deny-list/user endpoint.
 *
 * Covers the IDOR fix: the GET branch must apply the same admin/self check
 * as the POST/DELETE branch — any authenticated user was previously able to
 * read another user's personal deny list by supplying targetUserId.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';
import type { DenyListEntry } from '../services/denyList.js';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockListUserDenyEntries = jest.fn<(userId: string, type: string) => Promise<DenyListEntry[]>>();
const mockAddUserDenyEntry = jest.fn<() => Promise<void>>();
const mockRemoveUserDenyEntry = jest.fn<() => Promise<void>>();

// ── Wire up mocks ────────────────────────────────────────────────────────────

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: jest.fn<(userId: string) => Promise<string>>().mockResolvedValue('test-tenant'),
}));

jest.mock('../services/denyList.js', () => ({
  listUserDenyEntries: (userId: unknown, type: unknown) =>
    mockListUserDenyEntries(userId as string, type as string),
  addUserDenyEntry: (...args: unknown[]) => mockAddUserDenyEntry(...(args as [])),
  removeUserDenyEntry: (...args: unknown[]) => mockRemoveUserDenyEntry(...(args as [])),
  listGlobalDenyEntries: jest.fn<() => Promise<DenyListEntry[]>>().mockResolvedValue([]),
  addGlobalDenyEntry: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
  removeGlobalDenyEntry: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
  clearUserDenyList: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import { app } from '@azure/functions';
import '../functions/admin/getDenyList.js';

interface HttpRegistration {
  handler: (req: HttpRequest, ctx: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'denyListUser');
if (!registration) throw new Error('denyListUser handler was not registered');
const handler = registration[1].handler;

// ── Fixtures ─────────────────────────────────────────────────────────────────

const REQUESTER = 'user-requester';
const TARGET = 'user-target';

function makeAuth(userId: string): AuthResult {
  return {
    userId,
    session: {
      userId,
      homeAccountId: `home-${userId}`,
      displayName: 'Test User',
      email: `${userId}@example.com`,
      tenantId: 'test-tenant',
      accessToken: 'fake-token',
      expiresAt: Date.now() + 3_600_000,
      sessionToken: 'fake-session',
      sessionCreatedAt: Date.now(),
    },
  };
}

function makeGetRequest(query: Record<string, string> = {}): HttpRequest {
  const queryMap = new Map(Object.entries(query));
  return {
    method: 'GET',
    query: {
      get: (k: string) => queryMap.get(k) ?? null,
      has: (k: string) => queryMap.has(k),
    },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

function makePostRequest(body: unknown, query: Record<string, string> = {}): HttpRequest {
  const queryMap = new Map(Object.entries(query));
  return {
    method: 'POST',
    query: {
      get: (k: string) => queryMap.get(k) ?? null,
      has: (k: string) => queryMap.has(k),
    },
    headers: new Map<string, string>(),
    json: () => Promise.resolve(body),
  } as unknown as HttpRequest;
}

function makeDeleteRequest(body: unknown, query: Record<string, string> = {}): HttpRequest {
  const queryMap = new Map(Object.entries(query));
  return {
    method: 'DELETE',
    query: {
      get: (k: string) => queryMap.get(k) ?? null,
      has: (k: string) => queryMap.has(k),
    },
    headers: new Map<string, string>(),
    json: () => Promise.resolve(body),
  } as unknown as HttpRequest;
}

function makeContext(): InvocationContext {
  return { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
}

const SAMPLE_ENTRY: DenyListEntry = {
  partitionKey: `${TARGET}:sharepoint`,
  rowKey: 'c2l0ZXM',
  path: '/sites/private',
  description: '',
  addedBy: TARGET,
  addedAt: '2026-01-01T00:00:00.000Z',
};

// ── Tests ─────────────────────────────────────────────────────────────────────

beforeEach(() => {
  jest.clearAllMocks();
  mockListUserDenyEntries.mockResolvedValue([]);
  mockAddUserDenyEntry.mockResolvedValue(undefined);
  mockRemoveUserDenyEntry.mockResolvedValue(undefined);
});

describe('GET /api/manage/deny-list/user — authentication', () => {
  it('returns 401 when unauthenticated', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);

    const res = await handler(makeGetRequest({ targetUserId: TARGET }), makeContext());

    expect(res.status).toBe(401);
    expect((res.jsonBody as { error: string }).error).toContain('Authentication required');
  });
});

describe('GET /api/manage/deny-list/user — IDOR guard', () => {
  it('returns 403 when non-admin requests another user\'s deny list', async () => {
    mockAuthenticateRequest.mockResolvedValue(makeAuth(REQUESTER));
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(makeGetRequest({ targetUserId: TARGET }), makeContext());

    expect(res.status).toBe(403);
    expect((res.jsonBody as { error: string }).error).toContain('your own');
    expect(mockListUserDenyEntries).not.toHaveBeenCalled();
  });

  it('allows a global admin to read another user\'s deny list', async () => {
    mockAuthenticateRequest.mockResolvedValue(makeAuth(REQUESTER));
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockListUserDenyEntries.mockResolvedValue([SAMPLE_ENTRY]);

    const res = await handler(makeGetRequest({ targetUserId: TARGET }), makeContext());

    expect(res.status).toBe(200);
    expect((res.jsonBody as { entries: DenyListEntry[] }).entries).toHaveLength(1);
    expect(mockListUserDenyEntries).toHaveBeenCalledWith(TARGET, 'sharepoint');
  });

  it('allows a user to read their own deny list without admin check', async () => {
    mockAuthenticateRequest.mockResolvedValue(makeAuth(REQUESTER));
    mockListUserDenyEntries.mockResolvedValue([SAMPLE_ENTRY]);

    // targetUserId equals requesterId — should not call checkGlobalAdmin
    const res = await handler(makeGetRequest({ targetUserId: REQUESTER }), makeContext());

    expect(res.status).toBe(200);
    expect(mockCheckGlobalAdmin).not.toHaveBeenCalled();
    expect(mockListUserDenyEntries).toHaveBeenCalledWith(REQUESTER, 'sharepoint');
  });

  it('defaults to the requester\'s own list when targetUserId is omitted', async () => {
    mockAuthenticateRequest.mockResolvedValue(makeAuth(REQUESTER));
    mockListUserDenyEntries.mockResolvedValue([]);

    const res = await handler(makeGetRequest(), makeContext());

    expect(res.status).toBe(200);
    expect(mockCheckGlobalAdmin).not.toHaveBeenCalled();
    expect(mockListUserDenyEntries).toHaveBeenCalledWith(REQUESTER, 'sharepoint');
  });
});

describe('POST /api/manage/deny-list/user — IDOR guard (pre-existing)', () => {
  it('returns 403 when non-admin posts to another user\'s deny list', async () => {
    mockAuthenticateRequest.mockResolvedValue(makeAuth(REQUESTER));
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(
      makePostRequest({ targetUserId: TARGET, path: '/sites/private' }),
      makeContext(),
    );

    expect(res.status).toBe(403);
    expect(mockAddUserDenyEntry).not.toHaveBeenCalled();
  });

  it('allows a global admin to post to another user\'s deny list', async () => {
    mockAuthenticateRequest.mockResolvedValue(makeAuth(REQUESTER));
    mockCheckGlobalAdmin.mockResolvedValue(true);

    const res = await handler(
      makePostRequest({ targetUserId: TARGET, path: '/sites/private' }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    expect(mockAddUserDenyEntry).toHaveBeenCalledWith(TARGET, 'sharepoint', '/sites/private', undefined);
  });

  it('allows a user to add to their own deny list', async () => {
    mockAuthenticateRequest.mockResolvedValue(makeAuth(REQUESTER));

    const res = await handler(
      makePostRequest({ targetUserId: REQUESTER, path: '/sites/private' }),
      makeContext(),
    );

    expect(res.status).toBe(200);
    expect(mockCheckGlobalAdmin).not.toHaveBeenCalled();
    expect(mockAddUserDenyEntry).toHaveBeenCalledWith(REQUESTER, 'sharepoint', '/sites/private', undefined);
  });
});

describe('DELETE /api/manage/deny-list/user — IDOR guard (pre-existing)', () => {
  it('returns 403 when non-admin deletes from another user\'s deny list', async () => {
    mockAuthenticateRequest.mockResolvedValue(makeAuth(REQUESTER));
    mockCheckGlobalAdmin.mockResolvedValue(false);

    const res = await handler(
      makeDeleteRequest({ targetUserId: TARGET, path: '/sites/private' }),
      makeContext(),
    );

    expect(res.status).toBe(403);
    expect(mockRemoveUserDenyEntry).not.toHaveBeenCalled();
  });
});
