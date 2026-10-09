/**
 * Tests for the createMailFolder HTTP handler.
 *
 * Verifies:
 *   1. Root-level folder creation → POST /me/mailFolders with displayName
 *   2. Child-folder creation → POST /me/mailFolders/{parentId}/childFolders
 *   3. Custom mailboxId routes to /users/{mailboxId}/...
 *   4. Missing displayName → 400 (no Graph call)
 *   5. Deny-listed displayName → 403 (no Graph call)
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Mock declarations ────────────────────────────────────────────────────────

interface GraphFolderResponse {
  id: string;
  displayName: string;
  parentFolderId: string;
}

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockGetEnabledServices = jest.fn<(tenantId: string) => Promise<string[]>>();
const mockIsServiceDisabledForUser = jest.fn<() => Promise<boolean>>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetValidAccessToken = jest.fn<(userId: string) => Promise<string>>();

const mockGraphPost = jest.fn<(body: unknown) => Promise<GraphFolderResponse>>();
const lastGraphCall: { path: string | null } = { path: null };
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    lastGraphCall.path = path;
    return { post: mockGraphPost };
  },
}));

// Routing tests: the delegated-mailbox owner lookup (its own Graph call) is
// covered by mailboxOwner.test.ts and delegatedDenyList.test.ts.
jest.mock('../services/mailboxOwner.js', () => ({
  resolveDenySubject: (_g: unknown, callerId: string) => Promise.resolve(callerId),
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

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<(tenantId: string, userId: string) => Promise<boolean>>().mockResolvedValue(false),
}));

jest.mock('../services/auditLog.js', () => ({
  logAccess: jest.fn(),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

// ── Import the module under test (after mocks) ──────────────────────────────

import { app } from '@azure/functions';
import '../functions/mail/createMailFolder.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'createMailFolder');
if (!registration) throw new Error('createMailFolder handler was not registered');
const wrappedHandler = registration[1].handler;

// ── Test fixtures ────────────────────────────────────────────────────────────

const TENANT = 'tenant-xyz';
const USER = 'user-abc';
const TOKEN = 'fake-graph-token';

function makeRequest(body: unknown): HttpRequest {
  return {
    json: async () => body,
    query: { get: () => null },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

const fakeContext = {
  error: jest.fn(),
} as unknown as InvocationContext;

beforeEach(() => {
  jest.clearAllMocks();
  lastGraphCall.path = null;
  mockAuthenticateRequest.mockResolvedValue({
    userId: USER,
    session: {
      userId: USER,
      homeAccountId: 'home-abc',
      displayName: 'Test',
      email: 'test@example.com',
      tenantId: TENANT,
      accessToken: TOKEN,
      expiresAt: Date.now() + 3_600_000,
      sessionToken: 'fake-session',
      sessionCreatedAt: Date.now(),
    },
  } as AuthResult);
  mockGetEnabledServices.mockResolvedValue(['mail']);
  mockIsServiceDisabledForUser.mockResolvedValue(false);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockGetValidAccessToken.mockResolvedValue(TOKEN);
  mockGraphPost.mockResolvedValue({
    id: 'folder-id-1',
    displayName: '_Notifications',
    parentFolderId: 'msgroot',
  });
});

// ── Tests ───────────────────────────────────────────────────────────────────

describe('createMailFolder', () => {
  it('creates a root-level folder via POST /me/mailFolders', async () => {
    const res = await wrappedHandler(makeRequest({ displayName: '_Notifications' }), fakeContext);

    expect(res.status).toBe(201);
    expect(res.jsonBody).toEqual({
      id: 'folder-id-1',
      displayName: '_Notifications',
      parentFolderId: 'msgroot',
      status: 'created',
    });
    expect(lastGraphCall.path).toBe('/me/mailFolders');
    expect(mockGraphPost).toHaveBeenCalledWith({ displayName: '_Notifications' });
  });

  it('creates a child folder under parentFolderId', async () => {
    await wrappedHandler(
      makeRequest({ displayName: 'Sub', parentFolderId: 'parent-123' }),
      fakeContext,
    );

    expect(lastGraphCall.path).toBe('/me/mailFolders/parent-123/childFolders');
  });

  it('routes to /users/{mailboxId} when a non-me mailboxId is provided', async () => {
    await wrappedHandler(
      makeRequest({ displayName: 'Shared', mailboxId: 'other@example.com' }),
      fakeContext,
    );

    expect(lastGraphCall.path).toBe('/users/other@example.com/mailFolders');
  });

  it('treats mailboxId="me" the same as omitted', async () => {
    await wrappedHandler(makeRequest({ displayName: 'X', mailboxId: 'me' }), fakeContext);
    expect(lastGraphCall.path).toBe('/me/mailFolders');
  });

  it('returns 400 and skips Graph when displayName is missing', async () => {
    const res = await wrappedHandler(makeRequest({}), fakeContext);

    expect(res.status).toBe(400);
    expect(res.jsonBody).toEqual({ error: 'displayName is required' });
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('returns 403 and skips Graph when the displayName is on the deny list', async () => {
    mockIsPathDenied.mockResolvedValue(true);

    const res = await wrappedHandler(makeRequest({ displayName: '_Sensitive' }), fakeContext);

    expect(res.status).toBe(403);
    expect(res.jsonBody).toEqual({ error: 'Access restricted by deny list' });
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('returns 403 when mail service is disabled for the tenant', async () => {
    mockGetEnabledServices.mockResolvedValue(['sharepoint']);

    const res = await wrappedHandler(makeRequest({ displayName: 'X' }), fakeContext);

    expect(res.status).toBe(403);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('returns 401 when authentication fails', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);

    const res = await wrappedHandler(makeRequest({ displayName: 'X' }), fakeContext);

    expect(res.status).toBe(401);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});
