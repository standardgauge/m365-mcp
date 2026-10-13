/**
 * /api/manage/deny-list/user with `keyId`: one agent key's own deny entries.
 *
 * Covers:
 *   - the owner reads, adds and removes entries under the key's subject
 *   - another user is refused; a Global Administrator may read and add
 *     (narrow) but not remove (widen) another user's key's entries
 *   - an owner's removal is refused when the tenant allows narrowing only
 *   - a revoked or unknown key, and an unknown type, are refused
 *   - the key subject cannot be reached through targetUserId
 */
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';
import type { AgentKeyRecord, TenantAgentKeyPolicy } from '../services/agentKeys.js';
import type { DenyListEntry } from '../services/denyList.js';

const mockAuth = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockIsAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockRefusal = jest.fn<(operation: string, resource?: string) => void>();
const mockFindKey = jest.fn<(tenantId: string, keyId: string) => Promise<AgentKeyRecord | null>>();
const mockTenantPolicy = jest.fn<(tenantId: string) => Promise<TenantAgentKeyPolicy>>();
const mockList = jest.fn<(subject: string, type: string) => Promise<DenyListEntry[]>>();
const mockAdd = jest.fn<(subject: string, type: string, path: string, name?: string) => Promise<void>>();
const mockRemove = jest.fn<(subject: string, type: string, path: string) => Promise<void>>();
const mockLogAccess = jest.fn<(e: Record<string, unknown>) => void>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateConsoleRequest: (req: unknown) => mockAuth(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockIsAdmin(userId as string),
  auditAdminRefusal: (_auth: unknown, operation: unknown, resource?: unknown) =>
    mockRefusal(operation as string, resource as string | undefined),
  authorizeAdmin: async (auth: unknown, operation: unknown, resource?: unknown) => {
    const isAdmin = await mockIsAdmin((auth as AuthResult).userId);
    if (!isAdmin) mockRefusal(operation as string, resource as string | undefined);
    return isAdmin;
  },
}));
jest.mock('../services/tokenCache.js', () => ({
  getTenantId: async () => 'tenant-1',
  getTenantIdFromSession: () => 'tenant-1',
}));
jest.mock('../services/agentKeys.js', () => ({
  findAgentKey: (t: unknown, k: unknown) => mockFindKey(t as string, k as string),
  getTenantAgentKeyPolicy: (t: unknown) => mockTenantPolicy(t as string),
}));
jest.mock('../services/denyList.js', () => ({
  ...jest.requireActual<object>('../services/denyList.js'),
  listUserDenyEntries: (s: unknown, t: unknown) => mockList(s as string, t as string),
  addUserDenyEntry: (s: unknown, t: unknown, p: unknown, n?: unknown) => mockAdd(s as string, t as string, p as string, n as string | undefined),
  removeUserDenyEntry: (s: unknown, t: unknown, p: unknown) => mockRemove(s as string, t as string, p as string),
}));
jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (e: unknown) => mockLogAccess(e as Record<string, unknown>),
}));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/admin/getDenyList.js';

interface HttpRegistration {
  handler: (req: HttpRequest, ctx: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const handler = httpMock.mock.calls.find((c) => c[0] === 'denyListUser')![1].handler;

const OWNER = 'owner-oid';
const OTHER = 'other-oid';
const KEY_ID = 'abc123def456';
const SUBJECT = `key:${KEY_ID}`;

function auth(userId: string): AuthResult {
  return {
    userId,
    session: {
      userId, homeAccountId: `h-${userId}`, displayName: 'Test', email: `${userId}@example.com`, tenantId: 'tenant-1',
      accessToken: 'at', expiresAt: Date.now() + 3_600_000, sessionToken: 's', sessionCreatedAt: Date.now(),
    },
  };
}
function keyRecord(over: Partial<AgentKeyRecord> = {}): AgentKeyRecord {
  return {
    keyId: KEY_ID, tenantId: 'tenant-1', userId: OWNER, userEmail: 'owner@example.com', label: 'nightly',
    sessionStorageKey: 'row', policy: { services: ['mail', 'calendar'], readOnlyServices: [], allowedSites: [] },
    createdAt: Date.now() - 1000, expiresAt: Date.now() + 86_400_000, lastUsedAt: 0, revokedAt: 0, ...over,
  };
}
function req(method: string, query: Record<string, string>, body?: unknown): HttpRequest {
  const q = new Map(Object.entries(query));
  return {
    method,
    query: { get: (k: string) => q.get(k) ?? null, has: (k: string) => q.has(k) },
    headers: new Map<string, string>(),
    json: () => Promise.resolve(body),
  } as unknown as HttpRequest;
}
const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;
const policy = (ownersMayWiden: boolean): TenantAgentKeyPolicy => ({
  allowKeys: true, disabledServices: [], readOnlyServices: [], maxKeyLifetimeDays: 365, ownersMayWiden,
});

beforeEach(() => {
  jest.clearAllMocks();
  mockAuth.mockResolvedValue(auth(OWNER));
  mockIsAdmin.mockResolvedValue(false);
  mockFindKey.mockResolvedValue(keyRecord());
  mockTenantPolicy.mockResolvedValue(policy(true));
  mockList.mockResolvedValue([]);
  mockAdd.mockResolvedValue();
  mockRemove.mockResolvedValue();
});

describe('as the key owner', () => {
  it("reads the key's entries, not their own", async () => {
    const entry = { partitionKey: `${SUBJECT}:mail`, rowKey: 'x', path: 'Board', description: '', addedBy: SUBJECT, addedAt: '' };
    mockList.mockResolvedValue([entry]);
    const res = await handler(req('GET', { type: 'mail', keyId: KEY_ID }), ctx);
    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({ entries: [entry] });
    expect(mockFindKey).toHaveBeenCalledWith('tenant-1', KEY_ID);
    expect(mockList).toHaveBeenCalledWith(SUBJECT, 'mail');
    expect(mockIsAdmin).not.toHaveBeenCalled();
  });

  it('adds an entry under the key subject and audits it', async () => {
    const res = await handler(req('POST', {}, { type: 'calendar', path: 'Family', keyId: KEY_ID, addedByName: 'Owner' }), ctx);
    expect(res.status).toBe(200);
    expect(mockAdd).toHaveBeenCalledWith(SUBJECT, 'calendar', 'Family', 'Owner');
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      operation: 'policy.deny_list.agent_key.add', resource: `key:${KEY_ID}/calendar:Family`, result: 'allowed',
    }));
  });

  it('removes an entry when owners may widen', async () => {
    const res = await handler(req('DELETE', {}, { type: 'mail', path: 'Board', keyId: KEY_ID }), ctx);
    expect(res.status).toBe(200);
    expect(mockRemove).toHaveBeenCalledWith(SUBJECT, 'mail', 'Board');
  });

  it('is refused a removal when the tenant allows narrowing only', async () => {
    mockTenantPolicy.mockResolvedValue(policy(false));
    const res = await handler(req('DELETE', {}, { type: 'mail', path: 'Board', keyId: KEY_ID }), ctx);
    expect(res.status).toBe(403);
    expect(mockRemove).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      operation: 'policy.deny_list.agent_key.remove', result: 'denied',
    }));
  });

  it('may still add when the tenant allows narrowing only', async () => {
    mockTenantPolicy.mockResolvedValue(policy(false));
    const res = await handler(req('POST', {}, { type: 'mail', path: 'Board', keyId: KEY_ID }), ctx);
    expect(res.status).toBe(200);
  });

  it('is refused on a revoked key', async () => {
    mockFindKey.mockResolvedValue(keyRecord({ revokedAt: Date.now() - 5 }));
    const res = await handler(req('POST', {}, { type: 'mail', path: 'Board', keyId: KEY_ID }), ctx);
    expect(res.status).toBe(409);
    expect(mockAdd).not.toHaveBeenCalled();
  });

  it('gets 404 for a key not in the tenant', async () => {
    mockFindKey.mockResolvedValue(null);
    const res = await handler(req('GET', { type: 'mail', keyId: 'nope' }), ctx);
    expect(res.status).toBe(404);
  });

  it('gets 400 for an unknown type', async () => {
    const res = await handler(req('POST', {}, { type: 'tasks', path: 'x', keyId: KEY_ID }), ctx);
    expect(res.status).toBe(400);
    expect(mockAdd).not.toHaveBeenCalled();
  });
});

describe('as someone else', () => {
  beforeEach(() => mockAuth.mockResolvedValue(auth(OTHER)));

  it("is refused another user's key", async () => {
    const res = await handler(req('GET', { type: 'mail', keyId: KEY_ID }), ctx);
    expect(res.status).toBe(403);
    expect(mockList).not.toHaveBeenCalled();
    const post = await handler(req('POST', {}, { type: 'mail', path: 'Board', keyId: KEY_ID }), ctx);
    expect(post.status).toBe(403);
    expect(mockAdd).not.toHaveBeenCalled();
  });

  it('as a Global Administrator, may read and add', async () => {
    mockIsAdmin.mockResolvedValue(true);
    expect((await handler(req('GET', { type: 'mail', keyId: KEY_ID }), ctx)).status).toBe(200);
    expect((await handler(req('POST', {}, { type: 'mail', path: 'Board', keyId: KEY_ID }), ctx)).status).toBe(200);
    expect(mockAdd).toHaveBeenCalledWith(SUBJECT, 'mail', 'Board', undefined);
  });

  it('as a Global Administrator, may not remove', async () => {
    mockIsAdmin.mockResolvedValue(true);
    const res = await handler(req('DELETE', {}, { type: 'mail', path: 'Board', keyId: KEY_ID }), ctx);
    expect(res.status).toBe(403);
    expect(mockRemove).not.toHaveBeenCalled();
    expect(mockRefusal).toHaveBeenCalledWith('policy.deny_list.agent_key.remove', `key:${KEY_ID}/mail:Board`);
  });
});

describe('the key subject through targetUserId', () => {
  it('is refused even for a Global Administrator', async () => {
    mockIsAdmin.mockResolvedValue(true);
    expect((await handler(req('GET', { type: 'mail', targetUserId: SUBJECT }), ctx)).status).toBe(400);
    expect((await handler(req('DELETE', {}, { type: 'mail', path: 'Board', targetUserId: SUBJECT }), ctx)).status).toBe(400);
    expect(mockList).not.toHaveBeenCalled();
    expect(mockRemove).not.toHaveBeenCalled();
  });
});
