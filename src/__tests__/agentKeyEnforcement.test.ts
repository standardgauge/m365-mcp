/**
 * Enforcement of an agent key's policy on both surfaces.
 *
 * REST (services/policyEnforcement.ts):
 *   - a service outside the key is refused even though the tenant and user allow it
 *   - a service the key is read-only on refuses writes but not reads
 *   - the SharePoint allow-list is narrowed by the key's picks
 *
 * MCP (functions/mcp/mcpEndpoint.ts):
 *   - tools/list hides tools outside the key and the write tools of a key-read-only service
 *   - tools/call refuses both, and logs the refusal as denied
 *   - set_email_output_mode is hidden and refused for a key that holds mail
 *
 * REST routes (route-level, through the real wrapper):
 *   - SharePoint write/createFolder/delete refuse a key-read-only or tenant-read-only
 *     service, and honour the key's site picks on the body-level siteId
 *   - any non-GET route is gated as mutating even without the `mutating` flag
 *   - GET /api/sharepoint/sites?admin=true never returns the admin inventory to a key
 *   - GET /api/sharepoint/resolve-site refuses a key
 *
 * Routes that must never honour a key:
 *   - POST /api/mail/settings (loosening draft mode)
 *   - POST /api/auth/logout (ending the owner's sessions)
 */
import { jest } from '@jest/globals';
import { randomBytes } from 'crypto';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';
import type { UserSession } from '../services/tokenCache.js';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

const mockAuthenticateRequest = jest.fn<() => Promise<AuthResult | null>>();
const mockGetAllowedSites = jest.fn<() => Promise<Array<{ id: string; name: string }>>>();
const mockLogAccess = jest.fn<(e: Record<string, unknown>) => void>();
const mockDeleteAllUserSessions = jest.fn<() => Promise<void>>();
const mockSetUserEmailSettings = jest.fn<() => Promise<void>>();
const mockCheckGlobalAdmin = jest.fn<() => Promise<boolean>>();
const mockGetReadOnlyServices = jest.fn<() => Promise<string[]>>();
const mockGraphCall = jest.fn<(path: string, verb: string) => Promise<unknown>>();

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => mockAuthenticateRequest(),
  checkGlobalAdmin: () => mockCheckGlobalAdmin(),
  auditAdminRefusal: () => undefined,
}));
jest.mock('../services/serviceSettings.js', () => ({
  ALL_SERVICE_KEYS: ['mail', 'sharepoint', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'],
  getEnabledServices: async () => ['mail', 'sharepoint', 'calendar', 'onedrive'],
  getReadOnlyServices: () => mockGetReadOnlyServices(),
  getAllowedSites: () => mockGetAllowedSites(),
}));
jest.mock('../services/userServiceOverrides.js', () => ({
  isServiceDisabledForUser: async () => false,
  getUserServiceOverrides: async () => [],
}));
jest.mock('../services/userMailConfig.js', () => ({ isMailIndexingDisabled: async () => false }));
jest.mock('../services/denyList.js', () => ({
  isPathDenied: async () => false,
  filterDeniedPaths: async (_t: unknown, _u: unknown, _s: unknown, items: unknown) => items,
}));
jest.mock('../services/tokenCache.js', () => ({
  getTenantId: async () => 'tenant-a',
  getTenantIdFromSession: (s: { tenantId?: string }) => s.tenantId ?? 'tenant-a',
  getValidAccessTokenForSession: async () => 'at',
  deleteAllUserSessions: () => mockDeleteAllUserSessions(),
}));
jest.mock('../services/userEmailSettings.js', () => ({
  getUserEmailSettings: async () => ({ emailOutputMode: 'draft', preferredEmailOutputMode: 'draft', enforced: false, enforcedBy: null }),
  setUserEmailSettings: () => mockSetUserEmailSettings(),
}));
jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (e: unknown) => mockLogAccess(e as Record<string, unknown>),
}));
jest.mock('../services/frontendUrl.js', () => ({
  resolveAuthUrlBase: () => 'https://mcp.example.test',
  resolveFrontendUrl: () => 'https://mcp.example.test',
}));
jest.mock('../services/graphClient.js', () => ({
  createGraphClient: () => {
    const req = (path: string) => {
      const r = {
        select: () => r,
        header: () => r,
        get: () => mockGraphCall(path, 'get'),
        put: () => mockGraphCall(path, 'put'),
        post: () => mockGraphCall(path, 'post'),
        patch: () => mockGraphCall(path, 'patch'),
        delete: () => mockGraphCall(path, 'delete'),
      };
      return r;
    };
    return { api: req };
  },
}));
jest.mock('../services/securityHeaders.js', () => ({ withSecurity: (h: unknown) => h }));
jest.mock('../services/rateLimit.js', () => ({ withRateLimit: (_r: unknown, h: unknown) => h }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import { checkServiceEnabled, checkReadOnly, checkAllowedSite } from '../services/policyEnforcement.js';
import '../functions/mcp/mcpEndpoint.js';
import '../functions/mail/manageEmailSettings.js';
import '../functions/auth/logout.js';
import '../functions/sharepoint/writeFile.js';
import '../functions/sharepoint/createFolder.js';
import '../functions/sharepoint/deleteFile.js';
import '../functions/sharepoint/listSites.js';
import '../functions/sharepoint/resolveSite.js';
import '../functions/onedrive/writeOneDriveFile.js';

type Resp = { status: number; jsonBody?: Record<string, unknown> };
interface Reg { handler: (req: HttpRequest, ctx: InvocationContext) => Promise<Resp> }
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: Reg) => void>;
const mcp = httpMock.mock.calls.find((c) => c[0] === 'mcpEndpoint')![1].handler;
const mailSettings = httpMock.mock.calls.find((c) => c[0] === 'manageEmailSettings')![1].handler;
const route = (name: string) => httpMock.mock.calls.find((c) => c[0] === name)![1].handler;
const spWrite = route('writeSharepointFile');
const spCreateFolder = route('createSharepointFolder');
const spDelete = route('deleteSharepointFile');
const spListSites = route('listSites');
const spResolveSite = route('resolveSite');
const odWrite = route('writeOneDriveFile');
const logout = httpMock.mock.calls.find((c) => c[0] === 'authLogout' || c[0] === 'logout')![1].handler;

const TENANT = 'tenant-a';
const USER = 'user-1';
function session(agentKey?: UserSession['agentKey']): UserSession {
  return {
    userId: USER, homeAccountId: `${USER}.${TENANT}`, displayName: 'Adele Vance', email: 'adele@fabrikam.com',
    tenantId: TENANT, accessToken: 'at', expiresAt: Date.now() + 3_600_000, sessionToken: '',
    sessionCreatedAt: Date.now(), sessionAbsoluteCreatedAt: Date.now(), kind: 'client',
    deviceLabel: agentKey ? `agent-key:${agentKey.label}#${agentKey.keyId}` : 'laptop', agentKey,
  };
}
const KEY = { keyId: 'abc123def456', label: 'nightly', policy: { services: ['sharepoint'], readOnlyServices: ['sharepoint'], allowedSites: [] as Array<{ id: string; name: string }> } };
const ctx = { error: jest.fn(), warn: jest.fn() } as unknown as InvocationContext;

beforeEach(() => {
  jest.clearAllMocks();
  mockGetAllowedSites.mockResolvedValue([]);
  mockGetReadOnlyServices.mockResolvedValue([]);
  mockCheckGlobalAdmin.mockResolvedValue(false);
  mockGraphCall.mockImplementation(async (path: string) => {
    if (path.startsWith('/sites?search')) {
      return { value: ['site-a', 'site-b', 'site-c'].map((id) => ({ id, name: id, displayName: id, webUrl: `https://x/${id}`, description: null, createdDateTime: '', lastModifiedDateTime: '' })) };
    }
    return { id: 'item-1', name: 'a.txt', parentReference: { path: '/drive/root:' } };
  });
});

function restReq(method: string, opts: { body?: unknown; query?: Record<string, string>; params?: Record<string, string> } = {}): HttpRequest {
  const q = new Map(Object.entries(opts.query ?? {}));
  return {
    method,
    json: async () => opts.body,
    query: { get: (k: string) => q.get(k) ?? null },
    params: opts.params ?? {},
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}
const graphWrites = () => mockGraphCall.mock.calls.filter(([, verb]) => verb !== 'get');
const WRITABLE = { ...KEY, policy: { ...KEY.policy, readOnlyServices: [] as string[] } };
const SITE_A_ONLY = { ...WRITABLE, policy: { ...WRITABLE.policy, allowedSites: [{ id: 'site-a', name: 'A' }] } };

describe('REST policy checks under a key', () => {
  it('refuses a service outside the key even though tenant and user allow it', async () => {
    expect(await checkServiceEnabled(USER, 'mail', session())).toBeNull();
    expect(await checkServiceEnabled(USER, 'mail', session(KEY))).toMatchObject({ status: 403, error: expect.stringContaining('agent key') });
    expect(await checkServiceEnabled(USER, 'sharepoint', session(KEY))).toBeNull();
  });

  it('refuses writes, not reads, on a service the key is read-only on', async () => {
    expect(await checkReadOnly(USER, 'sharepoint', session())).toBeNull();
    expect(await checkReadOnly(USER, 'sharepoint', session(KEY))).toMatchObject({ status: 403, error: expect.stringContaining('read-only for this agent key') });
    // Read routes never call checkReadOnly; checkServiceEnabled still passes.
    expect(await checkServiceEnabled(USER, 'sharepoint', session(KEY))).toBeNull();
  });

  it("narrows the site allow-list to the key's picks", async () => {
    const picky = { ...KEY, policy: { ...KEY.policy, allowedSites: [{ id: 'site-a', name: 'A' }] } };
    // Tenant allows every site: the key still sees only its pick.
    expect(await checkAllowedSite(USER, 'site-b', session())).toBeNull();
    expect(await checkAllowedSite(USER, 'site-a', session(picky))).toBeNull();
    expect(await checkAllowedSite(USER, 'site-b', session(picky))).toMatchObject({ status: 403 });
    // Tenant restricts: intersection.
    mockGetAllowedSites.mockResolvedValue([{ id: 'site-a', name: 'A' }, { id: 'site-b', name: 'B' }]);
    expect(await checkAllowedSite(USER, 'site-b', session(picky))).toMatchObject({ status: 403 });
    expect(await checkAllowedSite(USER, 'site-a', session(picky))).toBeNull();
    // A key whose picks are all outside the tenant list sees nothing, not everything.
    mockGetAllowedSites.mockResolvedValue([{ id: 'site-b', name: 'B' }]);
    expect(await checkAllowedSite(USER, 'site-a', session(picky))).toMatchObject({ status: 403 });
    expect(await checkAllowedSite(USER, 'site-b', session(picky))).toMatchObject({ status: 403 });
  });
});

describe('SharePoint REST writes under a key (route level)', () => {
  const calls: Array<[string, () => HttpRequest]> = [
    ['PUT /api/sharepoint/files', () => restReq('PUT', { body: { siteId: 'site-b', path: 'a.txt', content: 'x' } })],
    ['POST /api/sharepoint/folders', () => restReq('POST', { body: { siteId: 'site-b', name: 'f' } })],
    ['DELETE /api/sharepoint/files/{itemId}/delete', () => restReq('DELETE', { params: { itemId: 'item-1' }, query: { siteId: 'site-b' } })],
  ];
  const handlers = [spWrite, spCreateFolder, spDelete];

  it.each(calls.map((c, i) => [c[0], i] as const))('%s refuses a key that is read-only on SharePoint, before Graph', async (_name, i) => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(KEY) });
    const r = await handlers[i](calls[i][1](), ctx);
    expect(r.status).toBe(403);
    expect(String(r.jsonBody?.error)).toContain('read-only for this agent key');
    expect(mockGraphCall).not.toHaveBeenCalled();
  });

  it.each(calls.map((c, i) => [c[0], i] as const))('%s refuses when the tenant has SharePoint read-only, with no key', async (_name, i) => {
    mockGetReadOnlyServices.mockResolvedValue(['sharepoint']);
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session() });
    const r = await handlers[i](calls[i][1](), ctx);
    expect(r.status).toBe(403);
    expect(mockGraphCall).not.toHaveBeenCalled();
  });

  it.each(calls.map((c, i) => [c[0], i] as const))("%s refuses a site outside the key's picks though the tenant allows every site", async (_name, i) => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(SITE_A_ONLY) });
    const r = await handlers[i](calls[i][1](), ctx);
    expect(r.status).toBe(403);
    expect(String(r.jsonBody?.error)).toContain('not in the allowed sites list');
    expect(graphWrites()).toEqual([]);
  });

  it("writes to a site inside the key's picks", async () => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(SITE_A_ONLY) });
    const r = await spWrite(restReq('PUT', { body: { siteId: 'site-a', path: 'a.txt', content: 'x' } }), ctx);
    expect(r.status).toBe(200);
    expect(graphWrites()).toHaveLength(1);
  });

  it('gates a non-GET route registered without `mutating` (OneDrive PUT) on a key-read-only service', async () => {
    const odReadOnly = { ...KEY, policy: { services: ['onedrive'], readOnlyServices: ['onedrive'], allowedSites: [] } };
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(odReadOnly) });
    const r = await odWrite(restReq('PUT', { body: { path: 'a.txt', content: 'x' } }), ctx);
    expect(r.status).toBe(403);
    expect(String(r.jsonBody?.error)).toContain('read-only for this agent key');
    expect(mockGraphCall).not.toHaveBeenCalled();
  });
});

describe('SharePoint admin-only REST under a key', () => {
  it("GET /api/sharepoint/sites?admin=true returns only the key's sites when the owner is a Global Admin", async () => {
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(SITE_A_ONLY) });
    const r = await spListSites(restReq('GET', { query: { admin: 'true' } }), ctx);
    expect(r.status).toBe(200);
    expect((r.jsonBody as { sites: Array<{ id: string }> }).sites.map((s) => s.id)).toEqual(['site-a']);
    expect(mockCheckGlobalAdmin).not.toHaveBeenCalled();
  });

  it('GET /api/sharepoint/sites?admin=true still gives a Global Admin browser session the inventory', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session() });
    const r = await spListSites(restReq('GET', { query: { admin: 'true' } }), ctx);
    expect((r.jsonBody as { sites: Array<{ id: string }> }).sites.map((s) => s.id)).toEqual(['site-a', 'site-b', 'site-c']);
  });

  it('GET /api/sharepoint/resolve-site refuses a key even when the owner is a Global Admin', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(true);
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(WRITABLE) });
    const r = await spResolveSite(restReq('GET', { query: { url: 'https://x.sharepoint.com/sites/b' } }), ctx);
    expect(r.status).toBe(403);
    expect(mockGraphCall).not.toHaveBeenCalled();
  });
});

function rpc(body: unknown): HttpRequest {
  return { method: 'POST', json: async () => body, headers: new Map<string, string>() } as unknown as HttpRequest;
}
async function toolNames(auth: AuthResult): Promise<string[]> {
  mockAuthenticateRequest.mockResolvedValue(auth);
  const r = await mcp(rpc({ jsonrpc: '2.0', id: 1, method: 'tools/list' }), ctx);
  const result = (r.jsonBody as { result: { tools: Array<{ name: string }> } }).result;
  return result.tools.map((t) => t.name);
}

describe('MCP dispatch under a key', () => {
  it('tools/list advertises only what the key may call', async () => {
    const full = await toolNames({ userId: USER, session: session() });
    const scoped = await toolNames({ userId: USER, session: session(KEY) });
    expect(full).toContain('search_mail');
    expect(full).toContain('write_sharepoint_file');
    expect(scoped).toContain('read_file');
    expect(scoped).toContain('search_sharepoint');
    expect(scoped).not.toContain('search_mail');
    expect(scoped).not.toContain('write_sharepoint_file');
    expect(scoped).not.toContain('delete_sharepoint_file');
    expect(scoped.length).toBeLessThan(full.length);
  });

  it('tools/call refuses a tool outside the key and logs it as denied', async () => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(KEY) });
    const r = await mcp(rpc({ jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'search_mail', arguments: { q: 'x' } } }), ctx);
    const result = (r.jsonBody as { result: { isError?: boolean; content: Array<{ text: string }> } }).result;
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('not permitted for this agent key');
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      operation: 'search_mail', result: 'denied', source: 'mcp', deviceLabel: 'agent-key:nightly#abc123def456',
      reason: expect.stringContaining('agent key'),
    }));
  });

  it('tools/call refuses a write tool on a key-read-only service', async () => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(KEY) });
    const r = await mcp(rpc({ jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'write_sharepoint_file', arguments: { siteId: 's', path: 'a.txt', content: 'x' } } }), ctx);
    const result = (r.jsonBody as { result: { isError?: boolean; content: Array<{ text: string }> } }).result;
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('read-only for this agent key');
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'write_sharepoint_file', result: 'denied', reason: expect.stringContaining('read-only for this agent key') }));
  });
});

describe('MCP owner-settings tools under a key', () => {
  const MAIL_KEY = { ...KEY, policy: { services: ['mail'], readOnlyServices: [] as string[], allowedSites: [] as Array<{ id: string; name: string }> } };

  it('tools/list hides set_email_output_mode from a mail-scoped key but keeps the read tool', async () => {
    const full = await toolNames({ userId: USER, session: session() });
    const scoped = await toolNames({ userId: USER, session: session(MAIL_KEY) });
    expect(full).toContain('set_email_output_mode');
    expect(scoped).toContain('send_mail');
    expect(scoped).toContain('get_email_output_mode');
    expect(scoped).not.toContain('set_email_output_mode');
  });

  it('tools/call refuses set_email_output_mode for a mail-scoped key, writes nothing, and logs it as denied', async () => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(MAIL_KEY) });
    const r = await mcp(rpc({ jsonrpc: '2.0', id: 4, method: 'tools/call', params: { name: 'set_email_output_mode', arguments: { emailOutputMode: 'draft' } } }), ctx);
    const result = (r.jsonBody as { result: { isError?: boolean; content: Array<{ text: string }> } }).result;
    expect(result.isError).toBe(true);
    expect(result.content[0].text).toContain('cannot be called with an agent key');
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      operation: 'set_email_output_mode', result: 'denied', source: 'mcp', reason: expect.stringContaining('agent key'),
    }));
  });

  it('tools/call still lets a browser session tighten to draft', async () => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session() });
    const r = await mcp(rpc({ jsonrpc: '2.0', id: 5, method: 'tools/call', params: { name: 'set_email_output_mode', arguments: { emailOutputMode: 'draft' } } }), ctx);
    const result = (r.jsonBody as { result: { isError?: boolean } }).result;
    expect(result.isError).toBeFalsy();
    expect(mockSetUserEmailSettings).toHaveBeenCalledTimes(1);
  });
});

describe('routes a key must never reach', () => {
  it('POST /api/mail/settings refuses a key and audits the attempt', async () => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(KEY) });
    const r = await mailSettings({ method: 'POST', json: async () => ({ emailOutputMode: 'send' }), query: new Map(), headers: new Map() } as unknown as HttpRequest, ctx);
    expect(r.status).toBe(403);
    expect(mockSetUserEmailSettings).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'policy.email_output_mode.set', result: 'denied' }));
  });

  it('GET /api/mail/settings still answers for a key', async () => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(KEY) });
    const r = await mailSettings({ method: 'GET', query: new Map(), headers: new Map() } as unknown as HttpRequest, ctx);
    expect(r.status).toBe(200);
  });

  it('POST /api/auth/logout refuses a key and leaves the owner signed in', async () => {
    mockAuthenticateRequest.mockResolvedValue({ userId: USER, session: session(KEY) });
    const headers = new Map([['origin', 'https://mcp.example.test'], ['host', 'mcp.example.test'], ['sec-fetch-site', 'same-origin']]);
    const r = await logout({ method: 'POST', headers, query: new Map() } as unknown as HttpRequest, ctx);
    expect(r.status).toBe(403);
    expect(mockDeleteAllUserSessions).not.toHaveBeenCalled();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'auth.logout', result: 'denied' }));
  });
});
