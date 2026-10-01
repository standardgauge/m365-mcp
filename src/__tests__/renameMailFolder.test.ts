/**
 * Tests for the renameMailFolder HTTP handler.
 *
 * Verifies:
 *   1. Rename → PATCH /me/mailFolders/{folderId} with { displayName }
 *   2. Custom mailboxId routes to /users/{mailboxId}/...
 *   3. Missing folderId → 400 (no Graph call)
 *   4. Missing displayName → 400 (no Graph call)
 *   5. Deny-listed target name → 403 (no Graph call)
 *   6. Deny-listed source folder → 403 (no Graph call)
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

const mockGraphPatch = jest.fn<(body: unknown) => Promise<GraphFolderResponse>>();
const lastGraphCall: { path: string | null } = { path: null };
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    lastGraphCall.path = path;
    return { patch: mockGraphPatch };
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
import '../functions/mail/renameMailFolder.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'renameMailFolder');
if (!registration) throw new Error('renameMailFolder handler was not registered');
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
  mockGraphPatch.mockResolvedValue({ id: 'folder-1', displayName: 'github', parentFolderId: 'archive-id' });
});

describe('renameMailFolder', () => {
  it('renames a folder via PATCH /me/mailFolders/{folderId}', async () => {
    const res = await wrappedHandler(makeRequest({ displayName: 'github' }, 'folder-1'), fakeContext);

    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({
      id: 'folder-1',
      displayName: 'github',
      parentFolderId: 'archive-id',
      status: 'renamed',
    });
    expect(lastGraphCall.path).toBe('/me/mailFolders/folder-1');
    expect(mockGraphPatch).toHaveBeenCalledWith({ displayName: 'github' });
  });

  it('routes to /users/{mailboxId} when a non-me mailboxId is provided', async () => {
    await wrappedHandler(
      makeRequest({ displayName: 'github', mailboxId: 'other@example.com' }, 'folder-1'),
      fakeContext,
    );
    expect(lastGraphCall.path).toBe('/users/other@example.com/mailFolders/folder-1');
  });

  it('returns 400 and skips Graph when folderId is missing', async () => {
    const res = await wrappedHandler(makeRequest({ displayName: 'github' }), fakeContext);
    expect(res.status).toBe(400);
    expect(mockGraphPatch).not.toHaveBeenCalled();
  });

  it('returns 400 and skips Graph when displayName is missing', async () => {
    const res = await wrappedHandler(makeRequest({}, 'folder-1'), fakeContext);
    expect(res.status).toBe(400);
    expect(res.jsonBody).toEqual({ error: 'displayName is required' });
    expect(mockGraphPatch).not.toHaveBeenCalled();
  });

  it('returns 403 and skips Graph when the target name is deny-listed', async () => {
    mockIsPathDenied.mockResolvedValue(true);
    const res = await wrappedHandler(makeRequest({ displayName: '_Sensitive' }, 'folder-1'), fakeContext);
    expect(res.status).toBe(403);
    expect(mockGraphPatch).not.toHaveBeenCalled();
  });

  it('returns 403 and skips Graph when the source folder is deny-listed', async () => {
    mockResolveMailFolderName.mockResolvedValue('_Sensitive');
    mockIsPathDenied.mockResolvedValue(true);
    const res = await wrappedHandler(makeRequest({ displayName: 'github' }, 'folder-1'), fakeContext);
    expect(res.status).toBe(403);
    expect(mockGraphPatch).not.toHaveBeenCalled();
  });
});
