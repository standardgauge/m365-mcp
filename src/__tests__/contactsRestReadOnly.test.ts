/**
 * Read-only enforcement for the new contacts REST write routes.
 *
 * codex review (PR #134): `POST /api/contacts/folders` and
 * `POST /api/contacts/batch` were registered without `{ mutating: true }`, so a
 * tenant with the `contacts` service in read-only mode could still create
 * folders and batch-create contacts through REST even though the MCP tools are
 * gated via `WRITE_TOOLS`. These tests drive each registered handler through the
 * real `withPolicyEnforcement` wrapper and assert:
 *   1. read-only mode → 403 and no Graph call (the regression guard: dropping
 *      `mutating: true` again makes the handler proceed and this fails), and
 *   2. read-only OFF → the write goes through, so the flag doesn't over-block.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Mock declarations ────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockGetEnabledServices = jest.fn<() => Promise<string[]>>();
const mockGetReadOnlyServices = jest.fn<() => Promise<string[]>>();
const mockIsServiceDisabledForUser = jest.fn<() => Promise<boolean>>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetTenantId = jest.fn<() => Promise<string>>();
const mockGetValidAccessToken = jest.fn<() => Promise<string>>();
const mockResolveDefaultContactFolder = jest.fn<() => Promise<string | null>>();

interface GraphCall { path: string; method: string; body?: unknown }
const graphCalls: GraphCall[] = [];
const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => ({
    post: (body?: unknown) => {
      graphCalls.push({ path, method: 'POST', body });
      if (path === '/$batch') {
        const requests = ((body as { requests?: Array<{ id: string }> })?.requests) ?? [];
        return Promise.resolve({
          responses: requests.map((r) => ({ id: r.id, status: 201, body: { id: `new-${r.id}` } })),
        });
      }
      return Promise.resolve({ id: 'folder-1', displayName: 'Imported', parentFolderId: null });
    },
  }),
}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: () => mockGetEnabledServices(),
  getReadOnlyServices: () => mockGetReadOnlyServices(),
  getAllowedSites: () => Promise.resolve([]),
}));

jest.mock('../services/userServiceOverrides.js', () => ({
  isServiceDisabledForUser: () => mockIsServiceDisabledForUser(),
}));

jest.mock('../services/denyList.js', () => ({
  isPathDenied: () => mockIsPathDenied(),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: () => mockGetTenantId(),
  getTenantIdFromSession: (session: { tenantId?: string }) => {
    if (!session?.tenantId) throw new Error('No tenantId in session');
    return session.tenantId;
  },
  getValidAccessTokenForSession: () => mockGetValidAccessToken(),
}));

jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => mockCreateGraphClient(),
}));

jest.mock('../services/containerResolver.js', () => ({
  resolveDefaultContactFolder: () => mockResolveDefaultContactFolder(),
}));

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<() => Promise<boolean>>().mockResolvedValue(false),
}));

jest.mock('../services/auditLog.js', () => ({
  logAccess: jest.fn(),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

// ── Import the route modules under test (after mocks) ───────────────────────

import { app } from '@azure/functions';
import '../functions/contacts/createContactFolder.js';
import '../functions/contacts/createContactsBatch.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;

function handlerFor(name: string) {
  const registration = httpMock.mock.calls.find((call) => call[0] === name);
  if (!registration) throw new Error(`${name} handler was not registered`);
  return registration[1].handler;
}
const folderHandler = handlerFor('createContactFolder');
const batchHandler = handlerFor('createContactsBatch');

// ── Fixtures ─────────────────────────────────────────────────────────────────

const TENANT = 'tenant-xyz';
const USER = 'user-abc';

function makeRequest(body: unknown): HttpRequest {
  return {
    method: 'POST',
    json: async () => body,
    query: { get: () => null },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}
const fakeContext = { error: jest.fn() } as unknown as InvocationContext;

beforeEach(() => {
  jest.clearAllMocks();
  graphCalls.length = 0;
  mockAuthenticateRequest.mockResolvedValue({
    userId: USER,
    session: {
      userId: USER,
      tenantId: TENANT,
      email: 'test@example.com',
      accessToken: 'fake-token',
      sessionToken: 'fake-session',
    },
  } as unknown as AuthResult);
  mockGetEnabledServices.mockResolvedValue(['contacts']);
  mockGetReadOnlyServices.mockResolvedValue([]);
  mockIsServiceDisabledForUser.mockResolvedValue(false);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockGetValidAccessToken.mockResolvedValue('fake-token');
  mockResolveDefaultContactFolder.mockResolvedValue(null);
});

// ── Tests ───────────────────────────────────────────────────────────────────

describe('POST /api/contacts/folders — read-only enforcement', () => {
  it('returns 403 and skips Graph when contacts is read-only', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['contacts']);

    const res = await folderHandler(makeRequest({ displayName: 'Imported' }), fakeContext);

    expect(res.status).toBe(403);
    expect(graphCalls).toHaveLength(0);
  });

  it('creates the folder when contacts is not read-only', async () => {
    const res = await folderHandler(makeRequest({ displayName: 'Imported' }), fakeContext);

    expect(res.status).toBe(201);
    expect(graphCalls).toEqual([
      { path: '/me/contactFolders', method: 'POST', body: { displayName: 'Imported' } },
    ]);
  });
});

describe('POST /api/contacts/batch — read-only enforcement', () => {
  const contacts = [{ givenName: 'Jane', surname: 'Doe' }];

  it('returns 403 and skips Graph when contacts is read-only', async () => {
    mockGetReadOnlyServices.mockResolvedValue(['contacts']);

    const res = await batchHandler(makeRequest({ contacts }), fakeContext);

    expect(res.status).toBe(403);
    expect(graphCalls).toHaveLength(0);
  });

  it('runs the batch when contacts is not read-only', async () => {
    const res = await batchHandler(makeRequest({ contacts }), fakeContext);

    expect(res.status).toBe(200);
    expect(res.jsonBody).toMatchObject({ created: 1, failed: 0 });
    expect(graphCalls.some((c) => c.path === '/$batch')).toBe(true);
  });
});
