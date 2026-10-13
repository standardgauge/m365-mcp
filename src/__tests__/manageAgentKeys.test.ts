/**
 * /api/manage/agent-keys and /api/manage/agent-key-policy.
 *
 * Covers:
 *   - both routes are console-only (401 without a console session)
 *   - POST mints a key: 201, the key value once, a `client` session bound to
 *     it that inherits the caller's absolute-lifetime anchor and names the key
 *   - POST refuses a service the caller cannot reach (tenant-disabled, opted
 *     out, or removed from keys by the tenant policy) and when keys are off
 *   - GET lists the caller's own keys; ?all=true needs Global Administrator
 *   - DELETE revokes an own key (row marked, session deleted, audited); a
 *     non-admin cannot revoke someone else's; an admin can
 *   - policy GET is open to any console user; POST needs Global Administrator,
 *     merges fields, validates service names, and audits before/after
 */
import { jest } from '@jest/globals';
import { randomBytes } from 'crypto';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';
import type { UserSession } from '../services/tokenCache.js';
import type { AgentKeyRecord, CreateAgentKeyInput, TenantAgentKeyPolicy } from '../services/agentKeys.js';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

const mockConsoleAuth = jest.fn<() => Promise<AuthResult | null>>();
const mockIsAdmin = jest.fn<() => Promise<boolean>>();
const mockAuditAdminRefusal = jest.fn<(op: string, res?: string) => void>();
const mockLogAccess = jest.fn<(e: Record<string, unknown>) => void>();
const mockGetEnabled = jest.fn<() => Promise<string[]>>();
const mockUserDisabled = jest.fn<() => Promise<string[]>>();
const mockStoreSession = jest.fn<(s: UserSession) => Promise<void>>();
const mockDeleteSessionByKey = jest.fn<(k: string, u: string) => Promise<void>>();
const mockCreateKey = jest.fn<(t: string, id: string, i: CreateAgentKeyInput) => Promise<AgentKeyRecord>>();
const mockFindKey = jest.fn<(t: string, id: string) => Promise<AgentKeyRecord | null>>();
const mockListKeys = jest.fn<(t: string, u?: string) => Promise<AgentKeyRecord[]>>();
const mockMarkRevoked = jest.fn<(r: AgentKeyRecord, by: string) => Promise<void>>();
const mockGetPolicy = jest.fn<() => Promise<TenantAgentKeyPolicy>>();
const mockSetPolicy = jest.fn<(t: string, p: TenantAgentKeyPolicy) => Promise<void>>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateConsoleRequest: () => mockConsoleAuth(),
  checkGlobalAdmin: () => mockIsAdmin(),
  auditAdminRefusal: (_a: unknown, op: unknown, res?: unknown) => mockAuditAdminRefusal(op as string, res as string | undefined),
  authorizeAdmin: async (_a: unknown, op: unknown, res?: unknown) => {
    const ok = await mockIsAdmin();
    if (!ok) mockAuditAdminRefusal(op as string, res as string | undefined);
    return ok;
  },
}));
jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (e: unknown) => mockLogAccess(e as Record<string, unknown>),
}));
jest.mock('../services/serviceSettings.js', () => ({
  ALL_SERVICE_KEYS: ['mail', 'sharepoint', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'],
  getEnabledServices: () => mockGetEnabled(),
  getAllowedSites: async () => [],
}));
jest.mock('../services/userServiceOverrides.js', () => ({
  getUserServiceOverrides: () => mockUserDisabled(),
}));
jest.mock('../services/tokenCache.js', () => ({
  storeSession: async (s: UserSession) => { s._storageKey = 'minted-row'; return mockStoreSession(s); },
  deleteSessionByKey: (k: unknown, u: unknown) => mockDeleteSessionByKey(k as string, u as string),
  getTenantIdFromSession: (s: { tenantId?: string }) => s.tenantId ?? '',
}));
jest.mock('../services/agentKeys.js', () => ({
  ...jest.requireActual<object>('../services/agentKeys.js'),
  createAgentKey: (t: unknown, id: unknown, i: unknown) => mockCreateKey(t as string, id as string, i as CreateAgentKeyInput),
  findAgentKey: (t: unknown, id: unknown) => mockFindKey(t as string, id as string),
  listAgentKeys: (t: unknown, u?: unknown) => mockListKeys(t as string, u as string | undefined),
  markAgentKeyRevoked: (r: unknown, by: unknown) => mockMarkRevoked(r as AgentKeyRecord, by as string),
  getTenantAgentKeyPolicy: () => mockGetPolicy(),
  setTenantAgentKeyPolicy: (t: unknown, p: unknown) => mockSetPolicy(t as string, p as TenantAgentKeyPolicy),
}));
jest.mock('../services/securityHeaders.js', () => ({ withSecurity: (h: unknown) => h }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/admin/manageAgentKeys.js';
import '../functions/admin/manageAgentKeyPolicy.js';
import { AGENT_KEY_PREFIX } from '../services/agentKeys.js';

type Resp = { status: number; jsonBody?: Record<string, unknown> & { keys?: unknown[]; key?: string; error?: string } };
interface Reg { handler: (req: HttpRequest, ctx: InvocationContext) => Promise<Resp> }
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: Reg) => void>;
const keys = httpMock.mock.calls.find((c) => c[0] === 'manageAgentKeys')![1].handler;
const policy = httpMock.mock.calls.find((c) => c[0] === 'manageAgentKeyPolicy')![1].handler;

const TENANT = 'tenant-a';
const USER = 'user-1';
const ANCHOR = Date.now() - 5 * 24 * 60 * 60 * 1000;
const SESSION = {
  userId: USER, homeAccountId: `${USER}.${TENANT}`, displayName: 'Adele Vance', email: 'adele@fabrikam.com',
  tenantId: TENANT, accessToken: 'at', expiresAt: Date.now() + 3_600_000, sessionToken: 'browser-token',
  sessionCreatedAt: Date.now(), sessionAbsoluteCreatedAt: ANCHOR, kind: 'browser' as const,
};
const AUTH = { userId: USER, session: SESSION } as unknown as AuthResult;
const ctx = { error: jest.fn() } as unknown as InvocationContext;

function req(method: string, body?: unknown, query: Record<string, string> = {}): HttpRequest {
  return {
    method,
    json: async () => body,
    query: new Map(Object.entries(query)),
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}
function rec(over: Partial<AgentKeyRecord> = {}): AgentKeyRecord {
  return {
    keyId: 'k1', tenantId: TENANT, userId: USER, userEmail: 'adele@fabrikam.com', label: 'nightly',
    sessionStorageKey: 'minted-row', policy: { services: ['mail'], readOnlyServices: [], allowedSites: [] },
    createdAt: 1, expiresAt: 0, lastUsedAt: 0, revokedAt: 0, _storageKey: 'kr', ...over,
  };
}

beforeEach(() => {
  jest.clearAllMocks();
  mockConsoleAuth.mockResolvedValue(AUTH);
  mockIsAdmin.mockResolvedValue(false);
  mockGetEnabled.mockResolvedValue(['mail', 'sharepoint', 'calendar']);
  mockUserDisabled.mockResolvedValue([]);
  mockStoreSession.mockResolvedValue();
  mockDeleteSessionByKey.mockResolvedValue();
  mockCreateKey.mockImplementation(async (_t, id, i) => rec({ keyId: id, label: i.label, policy: i.policy, expiresAt: i.expiresAt }));
  mockFindKey.mockResolvedValue(rec());
  mockListKeys.mockResolvedValue([rec()]);
  mockMarkRevoked.mockResolvedValue();
  mockGetPolicy.mockResolvedValue({ allowKeys: true, disabledServices: [], readOnlyServices: [] });
  mockSetPolicy.mockResolvedValue();
});

describe('console-only', () => {
  it.each([['GET'], ['POST'], ['DELETE']])('agent-keys %s is 401 without a console session', async (m) => {
    mockConsoleAuth.mockResolvedValue(null);
    expect((await keys(req(m), ctx)).status).toBe(401);
  });
  it.each([['GET'], ['POST']])('agent-key-policy %s is 401 without a console session', async (m) => {
    mockConsoleAuth.mockResolvedValue(null);
    expect((await policy(req(m), ctx)).status).toBe(401);
  });
});

describe('POST /api/manage/agent-keys', () => {
  it('mints a key bound to a fresh client session of the caller', async () => {
    const r = await keys(req('POST', { label: 'nightly', services: ['mail', 'sharepoint'], readOnlyServices: ['sharepoint'], expiresInDays: 30 }), ctx);
    expect(r.status).toBe(201);
    expect(r.jsonBody!.key!.startsWith(AGENT_KEY_PREFIX)).toBe(true);
    expect(r.jsonBody).toMatchObject({ label: 'nightly', policy: { services: ['mail', 'sharepoint'], readOnlyServices: ['sharepoint'], allowedSites: [] } });
    expect(r.jsonBody!.expiresAt).toBeGreaterThan(Date.now());

    const minted = mockStoreSession.mock.calls[0][0];
    expect(minted.kind).toBe('client');
    expect(minted.userId).toBe(USER);
    expect(minted.sessionAbsoluteCreatedAt).toBe(ANCHOR);
    expect(minted.sessionToken).not.toBe('browser-token');
    expect(minted.deviceLabel).toBe(`agent-key:nightly#${r.jsonBody!.keyId as string}`);

    const [token, keyId, input] = mockCreateKey.mock.calls[0];
    expect(token).toBe(r.jsonBody!.key);
    expect(keyId).toBe(r.jsonBody!.keyId);
    expect(input).toMatchObject({ tenantId: TENANT, userId: USER, sessionStorageKey: 'minted-row' });
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'policy.agent_key.create', result: 'allowed', resource: `key:${keyId}` }));
  });

  it('refuses a service outside what the caller may grant', async () => {
    mockUserDisabled.mockResolvedValue(['sharepoint']);
    let r = await keys(req('POST', { label: 'x', services: ['sharepoint'] }), ctx);
    expect(r.status).toBe(400);
    expect(r.jsonBody!.error).toContain('sharepoint');

    mockUserDisabled.mockResolvedValue([]);
    mockGetPolicy.mockResolvedValue({ allowKeys: true, disabledServices: ['calendar'], readOnlyServices: [] });
    r = await keys(req('POST', { label: 'x', services: ['calendar'] }), ctx);
    expect(r.status).toBe(400);

    r = await keys(req('POST', { label: 'x', services: ['teams'] }), ctx);
    expect(r.status).toBe(400);
    expect(mockStoreSession).not.toHaveBeenCalled();
  });

  it('refuses when the tenant has disabled keys, and audits it', async () => {
    mockGetPolicy.mockResolvedValue({ allowKeys: false, disabledServices: [], readOnlyServices: [] });
    const r = await keys(req('POST', { label: 'x', services: ['mail'] }), ctx);
    expect(r.status).toBe(403);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'policy.agent_key.create', result: 'denied' }));
  });

  it('validates the label and lifetime', async () => {
    expect((await keys(req('POST', { services: ['mail'] }), ctx)).status).toBe(400);
    expect((await keys(req('POST', { label: 'a#b', services: ['mail'] }), ctx)).status).toBe(400);
    expect((await keys(req('POST', { label: 'x', services: ['mail'], expiresInDays: 0 }), ctx)).status).toBe(400);
    expect((await keys(req('POST', { label: 'x', services: ['mail'], expiresInDays: 9999 }), ctx)).status).toBe(400);
  });
});

describe('GET /api/manage/agent-keys', () => {
  it('lists the caller’s own keys without the hash', async () => {
    const r = await keys(req('GET'), ctx);
    expect(r.status).toBe(200);
    expect(mockListKeys).toHaveBeenCalledWith(TENANT, USER);
    expect(r.jsonBody!.keys).toHaveLength(1);
    expect(JSON.stringify(r.jsonBody)).not.toContain('sessionStorageKey');
  });
  it('needs Global Administrator for ?all=true or another user', async () => {
    let r = await keys(req('GET', undefined, { all: 'true' }), ctx);
    expect(r.status).toBe(403);
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('admin.agent_keys.read', 'tenant');
    r = await keys(req('GET', undefined, { userId: 'other' }), ctx);
    expect(r.status).toBe(403);
    mockIsAdmin.mockResolvedValue(true);
    r = await keys(req('GET', undefined, { all: 'true' }), ctx);
    expect(r.status).toBe(200);
    expect(mockListKeys).toHaveBeenCalledWith(TENANT, undefined);
  });
});

describe('DELETE /api/manage/agent-keys', () => {
  it('revokes an own key: marks the row, deletes its session, audits', async () => {
    const r = await keys(req('DELETE', undefined, { keyId: 'k1' }), ctx);
    expect(r.status).toBe(200);
    expect(mockMarkRevoked).toHaveBeenCalledWith(expect.objectContaining({ keyId: 'k1' }), USER);
    expect(mockDeleteSessionByKey).toHaveBeenCalledWith('minted-row', USER);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'policy.agent_key.revoke', result: 'allowed', resource: 'key:k1' }));
  });
  it('is idempotent on an already revoked key', async () => {
    mockFindKey.mockResolvedValue(rec({ revokedAt: 5 }));
    expect((await keys(req('DELETE', undefined, { keyId: 'k1' }), ctx)).status).toBe(200);
    expect(mockMarkRevoked).not.toHaveBeenCalled();
  });
  it('refuses another user’s key unless Global Administrator', async () => {
    mockFindKey.mockResolvedValue(rec({ userId: 'other' }));
    let r = await keys(req('DELETE', undefined, { keyId: 'k1' }), ctx);
    expect(r.status).toBe(403);
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('policy.agent_key.revoke', 'key:k1');
    expect(mockDeleteSessionByKey).not.toHaveBeenCalled();
    mockIsAdmin.mockResolvedValue(true);
    r = await keys(req('DELETE', undefined, { keyId: 'k1' }), ctx);
    expect(r.status).toBe(200);
    expect(mockDeleteSessionByKey).toHaveBeenCalledWith('minted-row', 'other');
  });
  it('404s an unknown key and 400s a missing id', async () => {
    mockFindKey.mockResolvedValue(null);
    expect((await keys(req('DELETE', undefined, { keyId: 'nope' }), ctx)).status).toBe(404);
    expect((await keys(req('DELETE'), ctx)).status).toBe(400);
  });
});

describe('/api/manage/agent-key-policy', () => {
  it('GET returns the policy to any console user', async () => {
    const r = await policy(req('GET'), ctx);
    expect(r.status).toBe(200);
    expect(r.jsonBody).toEqual({ allowKeys: true, disabledServices: [], readOnlyServices: [] });
  });
  it('POST needs Global Administrator', async () => {
    const r = await policy(req('POST', { allowKeys: false }), ctx);
    expect(r.status).toBe(403);
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('policy.agent_key_policy.set', 'tenant');
    expect(mockSetPolicy).not.toHaveBeenCalled();
  });
  it('POST merges the given fields over the stored policy and audits before/after', async () => {
    mockIsAdmin.mockResolvedValue(true);
    mockGetPolicy.mockResolvedValue({ allowKeys: true, disabledServices: ['teams'], readOnlyServices: [] });
    const r = await policy(req('POST', { readOnlyServices: ['sharepoint'] }), ctx);
    expect(r.status).toBe(200);
    expect(mockSetPolicy).toHaveBeenCalledWith(TENANT, { allowKeys: true, disabledServices: ['teams'], readOnlyServices: ['sharepoint'] });
    const entry = mockLogAccess.mock.calls.find((c) => c[0].operation === 'policy.agent_key_policy.set')![0];
    expect(entry.before).toContain('"teams"');
    expect(entry.after).toContain('"sharepoint"');
  });
  it('POST validates its fields', async () => {
    mockIsAdmin.mockResolvedValue(true);
    expect((await policy(req('POST', {}), ctx)).status).toBe(400);
    expect((await policy(req('POST', { disabledServices: ['bogus'] }), ctx)).status).toBe(400);
    expect((await policy(req('POST', { allowKeys: 'no' }), ctx)).status).toBe(400);
  });
});
