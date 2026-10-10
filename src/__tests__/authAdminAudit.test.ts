/**
 * Audit rows for events that are not tool calls: sign-in and its failures,
 * install handoff, logout, refresh, and admin policy changes with the value
 * before and after.
 *
 * Each row needs an actor (empty when no user was identified yet), a target,
 * and, for a policy change, before and after. The client address comes from the
 * per-request scope and is covered in clientAddress.test.ts.
 */

import { randomBytes } from 'crypto';
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';
import { hashSessionToken } from '../services/credentialCrypto.js';
import { installConfirmationCode } from '../services/installConfirm.js';

// The callback mints a console session, which is MAC'd under the session key.
process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

type Row = Record<string, unknown>;

const mockLogAccess = jest.fn<(entry: Row) => void>();
const mockAcquireTokenByCode = jest.fn<() => Promise<unknown>>();
const mockAcquireTokenSilent = jest.fn<() => Promise<unknown>>();
const mockAttach = jest.fn<(nonce: string, record: Row) => Promise<boolean>>();
const mockConsume = jest.fn<(nonce: string) => Promise<unknown>>();
const mockCreateHandoff = jest.fn<(id: string, record: Row) => Promise<void>>();
const mockGetHandoff = jest.fn<(id: string) => Promise<unknown>>();
const mockRecordHandoffFailure = jest.fn<() => Promise<number | null>>();
const mockDeleteHandoff = jest.fn<() => Promise<boolean>>();
const mockAuthenticate = jest.fn<() => Promise<AuthResult | null>>();
const mockAuthenticateAllowExpired = jest.fn<() => Promise<AuthResult | null>>();
const mockIsAdmin = jest.fn<() => Promise<boolean>>();
const mockAuditAdminRefusal = jest.fn<(operation: string, resource?: string) => void>();
const mockDeleteAllUserSessions = jest.fn<() => Promise<void>>();
const mockListGlobal = jest.fn<() => Promise<Array<{ path: string; description?: string }>>>();
const mockAddGlobal = jest.fn<() => Promise<void>>();
const mockRemoveGlobal = jest.fn<() => Promise<void>>();
const mockListUser = jest.fn<(userId: string, type: string) => Promise<Array<{ path: string }>>>();
const mockClearUser = jest.fn<() => Promise<void>>();
const mockGetAllowedSites = jest.fn<() => Promise<Array<{ id: string; name: string }>>>();
const mockSetAllowedSites = jest.fn<() => Promise<void>>();

jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (entry: unknown) => mockLogAccess(entry as Row),
}));

jest.mock('../services/graphClient.js', () => ({
  acquireTokenByCode: () => mockAcquireTokenByCode(),
  acquireTokenSilent: () => mockAcquireTokenSilent(),
  createGraphClient: () => ({
    api: () => ({
      select: () => ({
        get: async () => ({ id: USER_ID, displayName: 'Adele Vance', mail: 'adele@fabrikam.com' }),
      }),
    }),
  }),
}));

jest.mock('../services/tokenCache.js', () => ({
  storeSession: async () => undefined,
  deleteAllUserSessions: () => mockDeleteAllUserSessions(),
  getTenantId: async () => TENANT,
  SESSION_TTL_MS: 7 * 24 * 60 * 60 * 1000,
}));

jest.mock('../services/tenantUtils.js', () => ({
  extractTenantId: (homeAccountId: string) => homeAccountId.split('.')[1],
}));

jest.mock('../services/tableStorage.js', () => ({
  attachSessionToInstallNonce: (nonce: string, record: unknown) => mockAttach(nonce, record as Row),
  consumeInstallNonce: (nonce: string) => mockConsume(nonce),
  createInstallHandoff: (id: string, record: unknown) => mockCreateHandoff(id, record as Row),
  getInstallHandoff: (id: string) => mockGetHandoff(id),
  recordInstallHandoffFailure: () => mockRecordHandoffFailure(),
  deleteInstallHandoff: () => mockDeleteHandoff(),
}));

jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => mockAuthenticate(),
  authenticateConsoleRequest: () => mockAuthenticate(),
  authenticateRequestAllowExpired: () => mockAuthenticateAllowExpired(),
  checkGlobalAdmin: () => mockIsAdmin(),
  auditAdminRefusal: (_auth: unknown, operation: unknown, resource?: unknown) =>
    mockAuditAdminRefusal(operation as string, resource as string | undefined),
  authorizeAdmin: async (_auth: unknown, operation: unknown, resource?: unknown) => {
    const isAdmin = await mockIsAdmin();
    if (!isAdmin) mockAuditAdminRefusal(operation as string, resource as string | undefined);
    return isAdmin;
  },
}));

jest.mock('../services/denyList.js', () => ({
  listGlobalDenyEntries: () => mockListGlobal(),
  addGlobalDenyEntry: () => mockAddGlobal(),
  removeGlobalDenyEntry: () => mockRemoveGlobal(),
  listUserDenyEntries: (userId: string, type: string) => mockListUser(userId, type),
  addUserDenyEntry: async () => undefined,
  removeUserDenyEntry: async () => undefined,
  clearUserDenyList: () => mockClearUser(),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getAllowedSites: () => mockGetAllowedSites(),
  setAllowedSites: () => mockSetAllowedSites(),
}));

jest.mock('../services/securityHeaders.js', () => ({ withSecurity: (h: unknown) => h }));
jest.mock('../services/rateLimit.js', () => ({ withRateLimit: (_r: unknown, h: unknown) => h }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/auth/callback.js';
import '../functions/auth/installConfirm.js';
import '../functions/auth/installPoll.js';
import '../functions/auth/logout.js';
import '../functions/auth/refresh.js';
import '../functions/admin/getDenyList.js';
import '../functions/admin/manageAllowedSites.js';

type Handler = (req: HttpRequest, ctx: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: { handler: Handler }) => void>;
function handlerFor(name: string): Handler {
  const reg = httpMock.mock.calls.find((c) => c[0] === name);
  if (!reg) throw new Error(`${name} was not registered`);
  return reg[1].handler;
}

const HOST = 'mcp.example.test';
const TENANT = 'tenant-abc';
const FOREIGN = 'tenant-xyz';
const USER_ID = 'user-1';
const STATE = 'state-123';
const NONCE = 'a'.repeat(64);
// The cookies /api/auth/login sets beside oauth_state (PKCE verifier and nonce).
const FLOW = `oauth_state=${STATE}; oauth_pkce=verifier; oauth_nonce=n`;

const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;

const SESSION = {
  userId: USER_ID,
  homeAccountId: `${USER_ID}.${TENANT}`,
  displayName: 'Adele Vance',
  email: 'adele@fabrikam.com',
  tenantId: TENANT,
  deviceLabel: 'laptop',
};
const AUTH = { userId: USER_ID, session: SESSION } as unknown as AuthResult;

function request(opts: {
  method?: string;
  query?: Record<string, string>;
  cookie?: string;
  body?: unknown;
  form?: Record<string, string>;
}): HttpRequest {
  // A same-origin browser request, so logout's Origin check passes.
  const headers = new Map<string, string>([['host', HOST], ['origin', `https://${HOST}`]]);
  if (opts.cookie) headers.set('cookie', opts.cookie);
  return {
    method: opts.method ?? 'GET',
    query: new Map(Object.entries(opts.query ?? {})),
    headers,
    json: async () => opts.body ?? {},
    text: async () => new URLSearchParams(opts.form ?? {}).toString(),
  } as unknown as HttpRequest;
}

function rows(operation: string): Row[] {
  return mockLogAccess.mock.calls.map((c) => c[0]).filter((r) => r.operation === operation);
}

const ORIGINAL_TENANT = process.env.AZURE_TENANT_ID;

beforeEach(() => {
  jest.clearAllMocks();
  process.env.AZURE_TENANT_ID = TENANT;
  mockAcquireTokenByCode.mockResolvedValue({
    accessToken: 'graph-token',
    homeAccountId: `${USER_ID}.${TENANT}`,
    expiresOn: new Date(Date.now() + 3_600_000),
  });
  mockAttach.mockResolvedValue(true);
  mockCreateHandoff.mockResolvedValue(undefined);
  mockDeleteHandoff.mockResolvedValue(true);
  mockIsAdmin.mockResolvedValue(true);
  mockListGlobal.mockResolvedValue([]);
  mockListUser.mockResolvedValue([]);
});

afterAll(() => {
  if (ORIGINAL_TENANT === undefined) delete process.env.AZURE_TENANT_ID;
  else process.env.AZURE_TENANT_ID = ORIGINAL_TENANT;
});

describe('sign-in (callback)', () => {
  const callback = handlerFor('callback');
  const ok = { code: 'c', state: STATE };

  it('records a browser sign-in against the signed-in user', async () => {
    const res = await callback(request({ query: ok, cookie: FLOW }), ctx);
    expect(res.status).toBe(302);
    expect(rows('auth.login')).toEqual([expect.objectContaining({
      tenantId: TENANT, userId: USER_ID, userEmail: 'adele@fabrikam.com', resource: 'browser', result: 'allowed',
    })]);
    expect(rows('auth.install_handoff')).toEqual([]);
  });

  it('records an installer sign-in and parks the handoff with the user\'s tenant', async () => {
    await callback(request({ query: ok, cookie: `${FLOW}; install_nonce=${NONCE}` }), ctx);
    expect(rows('auth.login')[0]).toMatchObject({ resource: 'install', result: 'allowed' });
    // Nothing is attached until the browser enters the installer's code.
    expect(rows('auth.install_handoff')).toEqual([]);
    expect(mockCreateHandoff.mock.calls[0][1]).toMatchObject({ tenantId: TENANT });
  });

  it('records a handoff that could not be recorded as denied', async () => {
    mockCreateHandoff.mockRejectedValue(new Error('storage down'));
    await callback(request({ query: ok, cookie: `${FLOW}; install_nonce=${NONCE}` }), ctx);
    expect(rows('auth.install_handoff')).toEqual([expect.objectContaining({
      userId: USER_ID, resource: 'attach', result: 'denied',
    })]);
  });

  it('records a state mismatch with no actor, in the instance tenant', async () => {
    const res = await callback(request({ query: ok, cookie: 'oauth_state=other' }), ctx);
    expect(res.status).toBe(403);
    expect(rows('auth.login')).toEqual([expect.objectContaining({
      tenantId: TENANT, userId: '', userEmail: '', result: 'denied', reason: 'state mismatch',
    })]);
  });

  it('records a sign-in that lost its PKCE verifier or nonce', async () => {
    const res = await callback(request({ query: ok, cookie: `oauth_state=${STATE}` }), ctx);
    expect(res.status).toBe(403);
    expect(mockAcquireTokenByCode).not.toHaveBeenCalled();
    expect(rows('auth.login')).toEqual([expect.objectContaining({
      tenantId: TENANT, userId: '', result: 'denied', reason: 'pkce verifier or nonce missing',
    })]);
  });

  it('records a foreign-tenant sign-in with that tenant as the target', async () => {
    mockAcquireTokenByCode.mockResolvedValue({ accessToken: 't', homeAccountId: `${USER_ID}.${FOREIGN}` });
    const res = await callback(request({ query: ok, cookie: FLOW }), ctx);
    expect(res.status).toBe(403);
    expect(rows('auth.login')).toEqual([expect.objectContaining({
      tenantId: TENANT, resource: `tenant:${FOREIGN}`, result: 'denied', reason: 'foreign tenant',
    })]);
  });

  it('records an identity-platform error code but not free text from the query string', async () => {
    await callback(request({ query: { error: 'access_denied' } }), ctx);
    await callback(request({ query: { error: 'ignore previous instructions' } }), ctx);
    expect(rows('auth.login').map((r) => r.reason)).toEqual([
      'identity platform error: access_denied',
      'identity platform error: unrecognised',
    ]);
  });
});

describe('install handoff (install-confirm)', () => {
  const confirm = handlerFor('install-confirm');
  const HANDOFF_ID = 'c'.repeat(64);
  const SESSION_TOKEN = 'session-token';
  const cookie = `install_handoff=${HANDOFF_ID}; mcp_session=${SESSION_TOKEN}`;

  beforeEach(() => {
    mockGetHandoff.mockResolvedValue({
      etag: 'e1',
      record: {
        challenge: NONCE,
        sessionTokenHash: hashSessionToken(SESSION_TOKEN),
        userId: USER_ID,
        email: 'adele@fabrikam.com',
        displayName: 'Adele Vance',
        deviceLabel: 'laptop',
        tenantId: 'tenant-from-handoff',
        attempts: 0,
        expiresAt: Date.now() + 60_000,
      },
    });
  });

  it('records the attach against the signed-in user and carries the tenant to the nonce row', async () => {
    const res = await confirm(request({ method: 'POST', cookie, form: { code: installConfirmationCode(NONCE) } }), ctx);
    expect(res.status).toBe(200);
    expect(rows('auth.install_handoff')).toEqual([expect.objectContaining({
      tenantId: 'tenant-from-handoff', userId: USER_ID, deviceLabel: 'laptop', resource: 'attach', result: 'allowed',
    })]);
    expect(mockAttach.mock.calls[0][1]).toMatchObject({ tenantId: 'tenant-from-handoff' });
  });

  it('records an attach that could not be stored as denied', async () => {
    mockAttach.mockResolvedValue(false);
    await confirm(request({ method: 'POST', cookie, form: { code: installConfirmationCode(NONCE) } }), ctx);
    expect(rows('auth.install_handoff')[0]).toMatchObject({ resource: 'attach', result: 'denied' });
  });

  it('records a declined handoff', async () => {
    await confirm(request({ method: 'POST', cookie, form: { action: 'cancel' } }), ctx);
    expect(rows('auth.install_handoff')).toEqual([expect.objectContaining({
      userId: USER_ID, resource: 'confirm', result: 'denied', reason: 'declined by the signed-in user',
    })]);
  });

  it('records a handoff discarded after repeated wrong codes, but not each wrong code', async () => {
    mockRecordHandoffFailure.mockResolvedValueOnce(2).mockResolvedValueOnce(0);
    await confirm(request({ method: 'POST', cookie, form: { code: 'WRONG-CODE' } }), ctx);
    expect(rows('auth.install_handoff')).toEqual([]);
    await confirm(request({ method: 'POST', cookie, form: { code: 'WRONG-CODE' } }), ctx);
    expect(rows('auth.install_handoff')).toEqual([expect.objectContaining({
      resource: 'confirm', result: 'denied', reason: 'discarded after repeated wrong confirmation codes',
    })]);
  });

  it('records nothing for a page view', async () => {
    expect((await confirm(request({ cookie }), ctx)).status).toBe(200);
    expect(mockLogAccess).not.toHaveBeenCalled();
  });
});

describe('install handoff (install-poll)', () => {
  const poll = handlerFor('install-poll');
  const verifier = { nonce_verifier: 'b'.repeat(32) };

  it('records the session handed to the installer against its user and tenant', async () => {
    mockConsume.mockResolvedValue({
      sessionToken: 's', userId: USER_ID, email: 'adele@fabrikam.com', displayName: 'Adele',
      deviceLabel: 'laptop', tenantId: 'tenant-from-record', expiresAt: Date.now() + 1000,
    });
    const res = await poll(request({ query: verifier }), ctx);
    expect(res.status).toBe(200);
    expect(rows('auth.install_handoff')).toEqual([expect.objectContaining({
      tenantId: 'tenant-from-record', userId: USER_ID, deviceLabel: 'laptop', resource: 'poll', result: 'allowed',
    })]);
    expect(JSON.stringify(mockLogAccess.mock.calls)).not.toContain('"s"');
  });

  it('records an expired nonce as denied', async () => {
    mockConsume.mockResolvedValue(null);
    expect((await poll(request({ query: verifier }), ctx)).status).toBe(410);
    expect(rows('auth.install_handoff')[0]).toMatchObject({ tenantId: TENANT, result: 'denied' });
  });

  it('records nothing while the sign-in is pending', async () => {
    mockConsume.mockResolvedValue('pending');
    expect((await poll(request({ query: verifier }), ctx)).status).toBe(202);
    expect(mockLogAccess).not.toHaveBeenCalled();
  });
});

describe('logout', () => {
  const logout = handlerFor('authLogout');

  it('records a logout that ended a session', async () => {
    mockAuthenticate.mockResolvedValue(AUTH);
    await logout(request({ method: 'POST' }), ctx);
    expect(rows('auth.logout')).toEqual([expect.objectContaining({ userId: USER_ID, result: 'allowed' })]);
  });

  it('records a failed cleanup as denied', async () => {
    mockAuthenticate.mockResolvedValue(AUTH);
    mockDeleteAllUserSessions.mockRejectedValue(new Error('storage down'));
    await logout(request({ method: 'POST' }), ctx);
    expect(rows('auth.logout')[0]).toMatchObject({ result: 'denied' });
  });

  it('records nothing for an anonymous logout', async () => {
    mockAuthenticate.mockResolvedValue(null);
    await logout(request({ method: 'POST' }), ctx);
    expect(mockLogAccess).not.toHaveBeenCalled();
  });
});

describe('refresh', () => {
  const refresh = handlerFor('refresh');

  it('records a successful refresh', async () => {
    mockAuthenticateAllowExpired.mockResolvedValue(AUTH);
    mockAcquireTokenSilent.mockResolvedValue({ accessToken: 'new', expiresOn: new Date() });
    expect((await refresh(request({ method: 'POST' }), ctx)).status).toBe(200);
    expect(rows('auth.refresh')).toEqual([expect.objectContaining({ userId: USER_ID, result: 'allowed' })]);
  });

  it('records a failed refresh as denied', async () => {
    mockAuthenticateAllowExpired.mockResolvedValue(AUTH);
    mockAcquireTokenSilent.mockRejectedValue(new Error('interaction_required'));
    expect((await refresh(request({ method: 'POST' }), ctx)).status).toBe(401);
    expect(rows('auth.refresh')[0]).toMatchObject({ result: 'denied' });
  });
});

describe('deny-list changes', () => {
  const global = handlerFor('denyListGlobal');
  const user = handlerFor('denyListUser');
  const clear = handlerFor('denyListUserClear');

  beforeEach(() => mockAuthenticate.mockResolvedValue(AUTH));

  it('records a global add with no entry before and the entry after', async () => {
    await global(request({ method: 'POST', body: { type: 'mail', path: 'Finance', description: 'board' } }), ctx);
    expect(rows('policy.deny_list.global.add')).toEqual([expect.objectContaining({
      tenantId: TENANT, userId: USER_ID, resource: 'mail:Finance', result: 'allowed',
      before: 'null', after: '{"path":"Finance","description":"board"}',
    })]);
  });

  it('records a global remove with the entry before and nothing after', async () => {
    mockListGlobal.mockResolvedValue([{ path: 'Finance', description: 'board' }]);
    await global(request({ method: 'DELETE', body: { type: 'mail', path: 'Finance' } }), ctx);
    expect(rows('policy.deny_list.global.remove')[0]).toMatchObject({
      before: '{"path":"Finance","description":"board"}', after: 'null',
    });
  });

  it('records a non-admin global change as an admin-check failure and changes nothing', async () => {
    mockIsAdmin.mockResolvedValue(false);
    const res = await global(request({ method: 'POST', body: { type: 'mail', path: 'Finance' } }), ctx);
    expect(res.status).toBe(403);
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('policy.deny_list.global.add', undefined);
    expect(mockAddGlobal).not.toHaveBeenCalled();
  });

  it("records a non-admin's attempt on another user's list", async () => {
    mockIsAdmin.mockResolvedValue(false);
    await user(request({ method: 'POST', body: { type: 'mail', path: 'HR', targetUserId: 'user-2' } }), ctx);
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('policy.deny_list.user.add', 'user:user-2/mail:HR');
  });

  it('records a user-list add against the target user', async () => {
    await user(request({ method: 'POST', body: { type: 'mail', path: 'HR', targetUserId: 'user-2' } }), ctx);
    expect(rows('policy.deny_list.user.add')[0]).toMatchObject({
      userId: USER_ID, resource: 'user:user-2/mail:HR', before: 'null', after: '{"path":"HR","description":""}',
    });
  });

  it('records a clear with both lists before', async () => {
    mockListUser.mockImplementation(async (_u, type) => (type === 'mail' ? [{ path: 'HR' }] : [{ path: '/sites/x' }]));
    await clear(request({ method: 'POST', body: {} }), ctx);
    expect(rows('policy.deny_list.user.clear')[0]).toMatchObject({
      resource: `user:${USER_ID}`,
      before: '{"sharepoint":["/sites/x"],"mail":["HR"]}',
      after: '{"sharepoint":[],"mail":[]}',
    });
  });
});

describe('allowed-sites changes', () => {
  const sites = handlerFor('manageAllowedSites');

  beforeEach(() => mockAuthenticate.mockResolvedValue(AUTH));

  it('records the list before and after', async () => {
    mockGetAllowedSites.mockResolvedValue([{ id: 's1', name: 'Finance' }]);
    const after = [{ id: 's1', name: 'Finance' }, { id: 's2', name: 'HR' }];
    await sites(request({ method: 'POST', body: { allowedSites: after } }), ctx);
    expect(rows('policy.allowed_sites.set')).toEqual([expect.objectContaining({
      tenantId: TENANT, userId: USER_ID, resource: 'tenant', result: 'allowed',
      before: JSON.stringify([{ id: 's1', name: 'Finance' }]), after: JSON.stringify(after),
    })]);
  });

  it('records a non-admin attempt and changes nothing', async () => {
    mockIsAdmin.mockResolvedValue(false);
    expect((await sites(request({ method: 'POST', body: { allowedSites: [] } }), ctx)).status).toBe(403);
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('policy.allowed_sites.set', 'tenant');
    expect(mockSetAllowedSites).not.toHaveBeenCalled();
  });
});
