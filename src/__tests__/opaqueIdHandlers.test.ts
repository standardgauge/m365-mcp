/**
 * Regression tests for opaque-ID injection protection on handlers that were
 * originally missed in's first pass (Codex CR finding).
 *
 * Covered handlers:
 *   - mail/updateMessage      (messageId route param, mailboxId query param)
 *   - contacts/updateContact  (contactId route param)
 *   - contacts/deleteContact  (contactId route param)
 *   - onedrive/deleteOneDriveItem  (itemId route param)
 *   - onedrive/readOneDriveFile    (itemId route param)
 *   - onedrive/listOneDriveFolders (parentId query param)
 *   - onedrive/createFolder        (parentId body param)
 *   - onedrive/moveOneDriveItem    (itemId + destinationFolderId body params)
 *
 * Each test verifies that a path-injection payload is rejected with HTTP 400
 * before any Graph API call is made.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

// ── shared mocks ──────────────────────────────────────────────────────────────

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockGetEnabledServices = jest.fn<(tenantId: string) => Promise<string[]>>();
const mockIsServiceDisabledForUser = jest.fn<() => Promise<boolean>>();
const mockIsPathDenied = jest.fn<() => Promise<boolean>>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetValidAccessToken = jest.fn<(userId: string) => Promise<string>>();
const mockResolveMailFolderName = jest.fn<() => Promise<string | null>>();
const mockResolveContactParentFolder = jest.fn<() => Promise<string | null>>();

const lastGraphCall: { path: string | null } = { path: null };
const mockGraphGet = jest.fn<() => Promise<unknown>>();
const mockGraphDelete = jest.fn<() => Promise<unknown>>();
const mockGraphPatch = jest.fn<() => Promise<unknown>>();
const mockGraphPost = jest.fn<() => Promise<unknown>>();
const mockGraphGetStream = jest.fn<() => Promise<unknown>>();

const mockCreateGraphClient = jest.fn(() => ({
  api: (path: string) => {
    lastGraphCall.path = path;
    return {
      select: () => ({
        get: mockGraphGet,
        top: () => ({ filter: () => ({ get: mockGraphGet }), get: mockGraphGet }),
      }),
      get: mockGraphGet,
      delete: mockGraphDelete,
      patch: mockGraphPatch,
      post: mockGraphPost,
      getStream: mockGraphGetStream,
      header: () => ({ put: jest.fn() }),
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
  filterDeniedPaths: (_tenantId: string, _userId: string, _service: string, items: unknown[]) =>
    Promise.resolve(items),
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
  resolveContactParentFolder: () => mockResolveContactParentFolder(),
}));

jest.mock('../services/policyEnforcement.js', () => {
  const real = jest.requireActual('../services/policyEnforcement.js') as Record<string, unknown>;
  return real;
});

jest.mock('../services/userMailConfig.js', () => ({
  isMailIndexingDisabled: jest.fn<() => Promise<boolean>>().mockResolvedValue(false),
}));

jest.mock('../services/auditLog.js', () => ({
  logAccess: jest.fn(),
}));

jest.mock('../services/tableStorage.js', () => ({
  getTableClient: jest.fn(),
}));

jest.mock('@azure/functions', () => ({
  app: { http: jest.fn() },
}));

// ── import handlers (triggers app.http registrations) ─────────────────────────

import { app } from '@azure/functions';
import '../functions/mail/updateMessage.js';
import '../functions/contacts/updateContact.js';
import '../functions/contacts/deleteContact.js';
import '../functions/onedrive/deleteOneDriveItem.js';
import '../functions/onedrive/readOneDriveFile.js';
import '../functions/onedrive/listOneDriveFolders.js';
import '../functions/onedrive/createFolder.js';
import '../functions/onedrive/moveOneDriveItem.js';

// ── helpers ───────────────────────────────────────────────────────────────────

interface HttpRegistration {
  handler: (req: HttpRequest, ctx: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;

function getHandler(name: string) {
  const reg = httpMock.mock.calls.find((c) => c[0] === name);
  if (!reg) throw new Error(`Handler '${name}' was not registered`);
  return reg[1].handler;
}

const TENANT = 'tenant-xyz';
const USER = 'user-abc';
const TOKEN = 'fake-graph-token';
const fakeContext = { error: jest.fn() } as unknown as InvocationContext;

const FAKE_AUTH: AuthResult = {
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
} as AuthResult;

function makeReq(opts: {
  params?: Record<string, string>;
  query?: Record<string, string>;
  body?: unknown;
}): HttpRequest {
  return {
    json: async () => opts.body ?? {},
    params: opts.params ?? {},
    query: { get: (k: string) => (opts.query && k in opts.query ? opts.query[k] : null) },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

/** A representative injection payload. */
const INJECT = 'legit-id/../../../users/victim';

beforeEach(() => {
  jest.clearAllMocks();
  lastGraphCall.path = null;
  mockAuthenticateRequest.mockResolvedValue(FAKE_AUTH);
  mockGetEnabledServices.mockResolvedValue(['mail', 'contacts', 'onedrive']);
  mockIsServiceDisabledForUser.mockResolvedValue(false);
  mockIsPathDenied.mockResolvedValue(false);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockGetValidAccessToken.mockResolvedValue(TOKEN);
  mockResolveMailFolderName.mockResolvedValue(null);
  mockResolveContactParentFolder.mockResolvedValue(null);
  mockGraphGet.mockResolvedValue({ isDraft: true, parentFolderId: null, name: 'file.txt', parentReference: { path: '/drive/root' }, file: { mimeType: 'text/plain' }, value: [] });
  mockGraphDelete.mockResolvedValue(undefined);
  mockGraphPatch.mockResolvedValue({ id: 'msg-1', subject: 'x', webLink: '', lastModifiedDateTime: '' });
  mockGraphPost.mockResolvedValue({ id: 'folder-1', name: 'New Folder', webUrl: '' });
  mockGraphGetStream.mockResolvedValue((async function* () { yield Buffer.from('hello'); })());
});

// ── updateMessage ─────────────────────────────────────────────────────────────

describe('updateMessage — opaque ID injection protection', () => {
  const handler = getHandler('updateMessage');

  it('rejects injected messageId before Graph call', async () => {
    const req = makeReq({
      params: { messageId: INJECT },
      body: { subject: 'x' },
    });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('rejects injected mailboxId before Graph call', async () => {
    const req = makeReq({
      params: { messageId: 'AAMkAGE1valid' },
      query: { mailboxId: INJECT },
      body: { subject: 'x' },
    });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('accepts clean messageId and proceeds', async () => {
    const req = makeReq({
      params: { messageId: 'AAMkAGE1valid' },
      body: { subject: 'Updated Subject' },
    });
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).not.toBeNull();
  });
});

// ── updateContact ─────────────────────────────────────────────────────────────

describe('updateContact — opaque ID injection protection', () => {
  const handler = getHandler('updateContact');

  it('rejects injected contactId before Graph call', async () => {
    const req = makeReq({
      params: { contactId: INJECT },
      body: { givenName: 'Alice' },
    });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('accepts clean contactId and proceeds', async () => {
    const req = makeReq({
      params: { contactId: 'AAMkAGE1valid' },
      body: { givenName: 'Alice' },
    });
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).not.toBeNull();
  });
});

// ── deleteContact ─────────────────────────────────────────────────────────────

describe('deleteContact — opaque ID injection protection', () => {
  const handler = getHandler('deleteContact');

  it('rejects injected contactId before Graph call', async () => {
    const req = makeReq({ params: { contactId: INJECT } });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('accepts clean contactId and proceeds', async () => {
    const req = makeReq({ params: { contactId: 'AAMkAGE1valid' } });
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).not.toBeNull();
  });
});

// ── deleteOneDriveItem ────────────────────────────────────────────────────────

describe('deleteOneDriveItem — opaque ID injection protection', () => {
  const handler = getHandler('deleteOneDriveItem');

  it('rejects injected itemId before Graph call', async () => {
    const req = makeReq({ params: { itemId: INJECT } });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('accepts clean itemId and proceeds', async () => {
    const req = makeReq({ params: { itemId: '01BWKFZZAH3GYXNQMQIQ' } });
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).not.toBeNull();
  });
});

// ── readOneDriveFile ──────────────────────────────────────────────────────────

describe('readOneDriveFile — opaque ID injection protection', () => {
  const handler = getHandler('readOneDriveFile');

  it('rejects injected itemId before Graph call', async () => {
    const req = makeReq({ params: { itemId: INJECT } });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('accepts clean itemId and proceeds', async () => {
    const req = makeReq({ params: { itemId: '01BWKFZZAH3GYXNQMQIQ' } });
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).not.toBeNull();
  });
});

// ── listOneDriveFolders ───────────────────────────────────────────────────────

describe('listOneDriveFolders — opaque ID injection protection', () => {
  const handler = getHandler('listOneDriveItems');

  it('rejects injected parentId before Graph call', async () => {
    const req = makeReq({ query: { parentId: INJECT } });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('accepts no parentId (lists root) and proceeds', async () => {
    const req = makeReq({});
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).toContain('/me/drive/root');
  });

  it('accepts clean parentId and proceeds', async () => {
    const req = makeReq({ query: { parentId: '01BWKFZZAH3GYXNQMQIQ' } });
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).toContain('01BWKFZZAH3GYXNQMQIQ');
  });
});

// ── createFolder ──────────────────────────────────────────────────────────────

describe('createFolder — opaque ID injection protection', () => {
  const handler = getHandler('createFolder');

  it('rejects injected parentId before Graph call', async () => {
    const req = makeReq({ body: { name: 'NewFolder', parentId: INJECT } });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('accepts no parentId (creates at root) and proceeds', async () => {
    const req = makeReq({ body: { name: 'NewFolder' } });
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).toContain('/me/drive/root/children');
  });

  it('accepts clean parentId and proceeds', async () => {
    const req = makeReq({ body: { name: 'NewFolder', parentId: '01BWKFZZAH3GYXNQMQIQ' } });
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).toContain('01BWKFZZAH3GYXNQMQIQ');
  });
});

// ── moveOneDriveItem ──────────────────────────────────────────────────────────

describe('moveOneDriveItem — opaque ID injection protection', () => {
  const handler = getHandler('moveOneDriveItem');

  it('rejects injected itemId before Graph call', async () => {
    const req = makeReq({ body: { itemId: INJECT, destinationFolderId: '01BWKFZZAH3GYXNQMQIQ' } });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('rejects injected destinationFolderId before Graph call', async () => {
    const req = makeReq({ body: { itemId: '01BWKFZZAH3GYXNQMQIQ', destinationFolderId: INJECT } });
    const res = await handler(req, fakeContext);
    expect(res.status).toBe(400);
    expect(lastGraphCall.path).toBeNull();
  });

  it('accepts clean itemId + destinationFolderId and proceeds', async () => {
    const req = makeReq({
      body: { itemId: '01BWKFZZAH3GYXNQMQIQ', destinationFolderId: '01BWKFZZAH3GYXNQMQIQ2' },
    });
    const res = await handler(req, fakeContext);
    expect(res.status).not.toBe(400);
    expect(lastGraphCall.path).not.toBeNull();
  });
});
