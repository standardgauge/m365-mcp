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

jest.mock('../services/telemetry.js', () => ({}));
jest.mock('../services/authMiddleware.js', () => ({
  authenticateRequest: () => mockAuthenticateRequest(),
  checkGlobalAdmin: async () => false,
  auditAdminRefusal: () => undefined,
}));
jest.mock('../services/serviceSettings.js', () => ({
  ALL_SERVICE_KEYS: ['mail', 'sharepoint', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'],
  getEnabledServices: async () => ['mail', 'sharepoint', 'calendar'],
  getReadOnlyServices: async () => [],
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
jest.mock('../services/securityHeaders.js', () => ({ withSecurity: (h: unknown) => h }));
jest.mock('../services/rateLimit.js', () => ({ withRateLimit: (_r: unknown, h: unknown) => h }));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import { checkServiceEnabled, checkReadOnly, checkAllowedSite } from '../services/policyEnforcement.js';
import '../functions/mcp/mcpEndpoint.js';
import '../functions/mail/manageEmailSettings.js';
import '../functions/auth/logout.js';

type Resp = { status: number; jsonBody?: Record<string, unknown> };
interface Reg { handler: (req: HttpRequest, ctx: InvocationContext) => Promise<Resp> }
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: Reg) => void>;
const mcp = httpMock.mock.calls.find((c) => c[0] === 'mcpEndpoint')![1].handler;
const mailSettings = httpMock.mock.calls.find((c) => c[0] === 'manageEmailSettings')![1].handler;
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
});

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
