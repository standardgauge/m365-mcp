/**
 * Tests for the moveMailFolder HTTP handler.
 *
 * Verifies:
 *   1. Move → POST /me/mailFolders/{folderId}/move with { destinationId }
 *   2. Custom mailboxId routes to /users/{mailboxId}/...
 *   3. Missing folderId → 400 (no Graph call)
 *   4. Missing destinationParentFolderId → 400 (no Graph call)
 *   5. Deny-listed moved/destination folder → 403 (no Graph call)
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

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
const mockResolveMailFolderName = jest.fn<() => Promise<string | null>>();

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

jest.mock('../services/containerResolver.js', () => ({
  resolveMailFolderName: () => mockResolveMailFolderName(),
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

import { app } from '@azure/functions';
import '../functions/mail/moveMailFolder.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'moveMailFolder');
if (!registration) throw new Error('moveMailFolder handler was not registered');
const wrappedHandler = registration[1].handler;

const TENANT = 'tenant-xyz';
const USER = 'user-abc';
const TOKEN = 'fake-graph-token';

function makeRequest(body: unknown, folderId?: string): HttpRequest {
  return {
    json: async () => body,
    params: folderId ? { folderId } : {},
    query: { get: () => null },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

const fakeContext = { error: jest.fn() } as unknown as InvocationContext;

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
  mockResolveMailFolderName.mockResolvedValue(null);
  mockGraphPost.mockResolvedValue({ id: 'folder-1', displayName: 'github', parentFolderId: 'archive-id' });
});

describe('moveMailFolder', () => {
  it('moves a folder via POST /me/mailFolders/{folderId}/move', async () => {
    const res = await wrappedHandler(
      makeRequest({ destinationParentFolderId: 'archive-id' }, 'folder-1'),
      fakeContext,
    );

    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({
      id: 'folder-1',
      displayName: 'github',
      parentFolderId: 'archive-id',
      status: 'moved',
    });
    expect(lastGraphCall.path).toBe('/me/mailFolders/folder-1/move');
    expect(mockGraphPost).toHaveBeenCalledWith({ destinationId: 'archive-id' });
  });

  it('routes to /users/{mailboxId} when a non-me mailboxId is provided', async () => {
    await wrappedHandler(
      makeRequest({ destinationParentFolderId: 'archive-id', mailboxId: 'other@example.com' }, 'folder-1'),
      fakeContext,
    );
    expect(lastGraphCall.path).toBe('/users/other@example.com/mailFolders/folder-1/move');
  });

  it('returns 400 and skips Graph when folderId is missing', async () => {
    const res = await wrappedHandler(makeRequest({ destinationParentFolderId: 'archive-id' }), fakeContext);
    expect(res.status).toBe(400);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('returns 400 and skips Graph when destinationParentFolderId is missing', async () => {
    const res = await wrappedHandler(makeRequest({}, 'folder-1'), fakeContext);
    expect(res.status).toBe(400);
    expect(res.jsonBody).toEqual({ error: 'destinationParentFolderId is required' });
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('returns 403 and skips Graph when a folder is deny-listed', async () => {
    mockResolveMailFolderName.mockResolvedValue('_Sensitive');
    mockIsPathDenied.mockResolvedValue(true);
    const res = await wrappedHandler(
      makeRequest({ destinationParentFolderId: 'archive-id' }, 'folder-1'),
      fakeContext,
    );
    expect(res.status).toBe(403);
    expect(mockGraphPost).not.toHaveBeenCalled();
  });
});
