/**
 * Tests for the deleteMailFolder HTTP handler.
 *
 * deleteMailFolder is destructive and guarded by confirm=true. It is deliberately
 * NOT exposed as an MCP bridge tool.
 *
 * Verifies:
 *   1. confirm=true → DELETE /me/mailFolders/{folderId}
 *   2. Missing/!=true confirm → 400 (no Graph call)
 *   3. Missing folderId → 400 (no Graph call)
 *   4. Custom mailboxId routes to /users/{mailboxId}/...
 *   5. Deny-listed folder → 403 (no Graph call)
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockGetEnabledServices = jest.fn<(tenantId: string) => Promise<string[]>>();
const mockIsServiceDisabledForUser = jest.fn<() => Promise<boolean>>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetValidAccessToken = jest.fn<(userId: string) => Promise<string>>();
const mockResolveMailFolderName = jest.fn<() => Promise<string | null>>();

const mockGraphDelete = jest.fn<() => Promise<unknown>>();
const lastGraphCall: { path: string | null } = { path: null };
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    lastGraphCall.path = path;
    return { delete: mockGraphDelete };
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
import '../functions/mail/deleteMailFolder.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'deleteMailFolder');
if (!registration) throw new Error('deleteMailFolder handler was not registered');
const wrappedHandler = registration[1].handler;

const TENANT = 'tenant-xyz';
const USER = 'user-abc';
const TOKEN = 'fake-graph-token';

function makeRequest(query: Record<string, string>, folderId?: string): HttpRequest {
  return {
    json: async () => ({}),
    params: folderId ? { folderId } : {},
    query: { get: (k: string) => (k in query ? query[k] : null) },
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
  mockGraphDelete.mockResolvedValue(undefined);
});

describe('deleteMailFolder', () => {
  it('deletes a folder via DELETE /me/mailFolders/{folderId} when confirm=true', async () => {
    const res = await wrappedHandler(makeRequest({ confirm: 'true' }, 'folder-1'), fakeContext);

    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({ id: 'folder-1', status: 'deleted' });
    expect(lastGraphCall.path).toBe('/me/mailFolders/folder-1');
    expect(mockGraphDelete).toHaveBeenCalled();
  });

  it('returns 400 and skips Graph when confirm is not "true"', async () => {
    const res = await wrappedHandler(makeRequest({}, 'folder-1'), fakeContext);
    expect(res.status).toBe(400);
    expect(mockGraphDelete).not.toHaveBeenCalled();
  });

  it('returns 400 and skips Graph when folderId is missing', async () => {
    const res = await wrappedHandler(makeRequest({ confirm: 'true' }), fakeContext);
    expect(res.status).toBe(400);
    expect(mockGraphDelete).not.toHaveBeenCalled();
  });

  it('routes to /users/{mailboxId} when a non-me mailboxId is provided', async () => {
    await wrappedHandler(
      makeRequest({ confirm: 'true', mailboxId: 'other@example.com' }, 'folder-1'),
      fakeContext,
    );
    expect(lastGraphCall.path).toBe('/users/other@example.com/mailFolders/folder-1');
  });

  it('returns 403 and skips Graph when the folder is deny-listed', async () => {
    mockResolveMailFolderName.mockResolvedValue('_Sensitive');
    mockIsPathDenied.mockResolvedValue(true);
    const res = await wrappedHandler(makeRequest({ confirm: 'true' }, 'folder-1'), fakeContext);
    expect(res.status).toBe(403);
    expect(mockGraphDelete).not.toHaveBeenCalled();
  });
});
