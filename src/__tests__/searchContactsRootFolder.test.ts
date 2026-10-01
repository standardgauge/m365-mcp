/**
 * Regression test for the contacts-root routing bug on the REST
 * search_contacts surface (Codex CR finding on PR #134).
 *
 * list_contact_folders reports the default contacts folder as the synthetic id
 * `contacts-root`. Before the fix, the REST search handler built
 * `/me/contactFolders/contacts-root/contacts` for that id instead of routing it
 * through contactsApiPath -> `/me/contacts`, so the advertised default-folder id
 * could not be searched. This verifies the REST handler now maps it correctly
 * while still routing an explicit non-root folder id to the folder path.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── mocks ─────────────────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockGetEnabledServices = jest.fn<(tenantId: string) => Promise<string[]>>();
const mockIsServiceDisabledForUser = jest.fn<() => Promise<boolean>>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetValidAccessToken = jest.fn<(userId: string) => Promise<string>>();

const lastGraphCall: { path: string | null } = { path: null };
const mockGraphGet = jest.fn<() => Promise<unknown>>();

const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    lastGraphCall.path = path;
    return {
      select: () => ({
        get: mockGraphGet,
        top: () => ({ filter: () => ({ get: mockGraphGet }), get: mockGraphGet }),
      }),
      get: mockGraphGet,
      top: () => ({ filter: () => ({ get: mockGraphGet }), get: mockGraphGet }),
      filter: () => ({ get: mockGraphGet }),
    };
  },
}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: (...args: unknown[]) => mockGetEnabledServices(args[0] as string),
  getReadOnlyServices: () => Promise.resolve([]),
  getAllowedSites: () => Promise.resolve([]),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  isServiceDisabledForUser: () => mockIsServiceDisabledForUser(),
}));

jest.mock('../services/denyList.js', () => ({
  isPathDenied: () => mockIsPathDenied(),
  filterDeniedPaths: (_t: string, _u: string, _s: string, items: unknown[]) => Promise.resolve(items),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: (...args: unknown[]) => mockGetTenantId(args[0] as string),
  getTenantIdFromSession: (session: { tenantId?: string }) => {
    if (!session?.tenantId) throw new Error('No tenantId in session');
    return session.tenantId;
  },
  getValidAccessToken: (...args: unknown[]) => mockGetValidAccessToken(args[0] as string),
  getValidAccessTokenForSession: (session: { userId?: string }) =>
    mockGetValidAccessToken(session?.userId as string),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: (...args: unknown[]) => (mockCreateGraphClient as (...a: unknown[]) => unknown)(...args),
}));

jest.mock('../services/containerResolver.js', () => ({
  normalizeContactFolderId: (_g: unknown, _u: unknown, id: string) => Promise.resolve(id || 'contacts-root'),
}));

jest.mock('../services/policyEnforcement.js', () => {
  const real = jest.requireActual('../services/policyEnforcement.js') as Record<string, unknown>;
  return real;
});

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<() => Promise<boolean>>().mockResolvedValue(false),
}));

jest.mock('../services/auditLog.js', () => ({ logAccess: jest.fn() }));
jest.mock('../services/tableStorage.js', () => ({ getTableClient: jest.fn() }));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

// ── import handler after mocks ──────────────────────────────────────────────────

import { app } from '@azure/functions';
import '../functions/contacts/searchContacts.js';

interface HttpRegistration {
  handler: (req: HttpRequest, ctx: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const reg = httpMock.mock.calls.find((c) => c[0] === 'searchContacts');
if (!reg) throw new Error("Handler 'searchContacts' was not registered");
const handler = reg[1].handler;

const TENANT = 'tenant-xyz';
const USER = 'user-abc';
const TOKEN = 'fake-graph-token';
const fakeContext = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;

const FAKE_AUTH = {
  userId: USER,
  session: { userId: USER, tenantId: TENANT, accessToken: TOKEN, sessionToken: 'sess' },
} as AuthResult;

function makeReq(query: Record<string, string>): HttpRequest {
  return {
    json: async () => ({}),
    params: {},
    query: { get: (k: string) => (k in query ? query[k] : null) },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

beforeEach(() => {
  jest.clearAllMocks();
  lastGraphCall.path = null;
  mockAuthenticateRequest.mockResolvedValue(FAKE_AUTH);
  mockGetEnabledServices.mockResolvedValue(['contacts']);
  mockIsServiceDisabledForUser.mockResolvedValue(false);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockGetValidAccessToken.mockResolvedValue(TOKEN);
  mockGraphGet.mockResolvedValue({ value: [] });
});

describe('searchContacts (REST) — contacts-root routing', () => {
  it('maps the synthetic contacts-root folderId to /me/contacts', async () => {
    const res = await handler(makeReq({ folderId: 'contacts-root' }), fakeContext);
    expect(res.status).toBe(200);
    expect(lastGraphCall.path).toBe('/me/contacts');
  });

  it('routes an explicit non-root folderId to /me/contactFolders/{id}/contacts', async () => {
    const res = await handler(makeReq({ folderId: 'folder-xyz' }), fakeContext);
    expect(res.status).toBe(200);
    expect(lastGraphCall.path).toBe('/me/contactFolders/folder-xyz/contacts');
  });

  it('routes an omitted folderId to /me/contacts', async () => {
    const res = await handler(makeReq({}), fakeContext);
    expect(res.status).toBe(200);
    expect(lastGraphCall.path).toBe('/me/contacts');
  });
});
