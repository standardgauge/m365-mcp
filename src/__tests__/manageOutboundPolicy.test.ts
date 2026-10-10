/**
 * Tests for the /api/manage/outbound-policy endpoint.
 *
 * Covers:
 *   - 401 when unauthenticated; 403 for non-admins on GET and POST, nothing read or written, refusal audited
 *   - GET without userId: the tenant row; with userId: both rows and the effective modes
 *   - POST scope=tenant / scope=user: writes only the channels given, stamps the admin, logs it with before/after
 *   - POST validation: bad scope, missing userId, unknown mode, unknown field, no channel given
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

type Scope = { scope: 'tenant' } | { scope: 'user'; userId: string };

const mockAuthenticateRequest = jest.fn<() => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockGetOutboundPolicy = jest.fn<(tenantId: string, target: Scope) => Promise<unknown>>();
const mockGetOutboundEnforcement = jest.fn<(tenantId: string, userId: string) => Promise<unknown>>();
const mockSetOutboundPolicy = jest.fn<(tenantId: string, target: Scope, modes: Record<string, string>, setBy: string) => Promise<unknown>>();
const mockAuditAdminRefusal = jest.fn<(operation: string) => void>();
const mockLogAccess = jest.fn<(entry: Record<string, unknown>) => void>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateConsoleRequest: () => mockAuthenticateRequest(),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
  authorizeAdmin: async (auth: unknown, operation: unknown) => {
    const isAdmin = await mockCheckGlobalAdmin((auth as AuthResult).userId);
    if (!isAdmin) mockAuditAdminRefusal(operation as string);
    return isAdmin;
  },
}));
jest.mock('../services/tokenCache.js', () => ({ getTenantId: async () => TENANT }));
jest.mock('../services/outboundPolicy.js', () => ({
  ...jest.requireActual<typeof import('../services/outboundPolicy.js')>('../services/outboundPolicy.js'),
  getOutboundPolicy: (t: unknown, s: unknown) => mockGetOutboundPolicy(t as string, s as Scope),
  getOutboundEnforcement: (t: unknown, u: unknown) => mockGetOutboundEnforcement(t as string, u as string),
  setOutboundPolicy: (t: unknown, s: unknown, m: unknown, b: unknown) =>
    mockSetOutboundPolicy(t as string, s as Scope, m as Record<string, string>, b as string),
}));
jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (e: unknown) => mockLogAccess(e as Record<string, unknown>),
}));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));
jest.mock('../services/securityHeaders.js', () => ({ withSecurity: (h: unknown) => h }));

import { app } from '@azure/functions';
import '../functions/admin/manageOutboundPolicy.js';

const TENANT = 'test-tenant';
const ADMIN = 'admin-user';
const TARGET = 'target-user';

interface HttpRegistration {
  methods: string[];
  route: string;
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const registration = httpMock.mock.calls.find((call) => call[0] === 'manageOutboundPolicy');
if (!registration) throw new Error('manageOutboundPolicy handler was not registered');
const handler = registration[1].handler;

const AUTH = { userId: ADMIN, session: { userId: ADMIN, tenantId: TENANT, email: 'admin@example.com' } } as unknown as AuthResult;
const ALLOW = { calendarInvites: 'allow', eventResponses: 'allow', teamsMessages: 'allow' };
const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;

function get(query: Record<string, string> = {}) {
  const q = new Map(Object.entries(query));
  return handler({ method: 'GET', query: { get: (k: string) => q.get(k) ?? null }, headers: new Map() } as unknown as HttpRequest, ctx);
}
function post(body: unknown) {
  return handler({ method: 'POST', query: { get: () => null }, headers: new Map(), json: async () => body } as unknown as HttpRequest, ctx);
}

beforeEach(() => {
  jest.clearAllMocks();
  jest.spyOn(console, 'log').mockImplementation(() => undefined);
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockCheckGlobalAdmin.mockResolvedValue(true);
  mockGetOutboundPolicy.mockResolvedValue(ALLOW);
  mockSetOutboundPolicy.mockImplementation(async (_t, _s, modes, setBy) => ({ ...ALLOW, ...modes, updatedBy: setBy }));
});

afterEach(() => jest.restoreAllMocks());

it('is mounted at api/manage/outbound-policy for GET and POST', () => {
  expect(registration[1].route).toBe('api/manage/outbound-policy');
  expect(registration[1].methods).toEqual(['GET', 'POST']);
});

it('401 without a session', async () => {
  mockAuthenticateRequest.mockResolvedValue(null);
  expect((await get()).status).toBe(401);
});

it('403 for a non-admin on GET and POST, touching nothing', async () => {
  mockCheckGlobalAdmin.mockResolvedValue(false);
  expect((await get()).status).toBe(403);
  expect((await post({ scope: 'tenant', teamsMessages: 'block' })).status).toBe(403);
  expect(mockGetOutboundPolicy).not.toHaveBeenCalled();
  expect(mockSetOutboundPolicy).not.toHaveBeenCalled();
  expect(mockAuditAdminRefusal.mock.calls).toEqual([['admin.outbound_policy.read'], ['set_outbound_policy']]);
});

it('GET returns the tenant row', async () => {
  const res = await get();
  expect(res.status).toBe(200);
  expect(res.jsonBody).toEqual({ tenant: ALLOW });
  expect(mockGetOutboundPolicy).toHaveBeenCalledWith(TENANT, { scope: 'tenant' });
});

it('GET with userId returns both rows and the effective modes', async () => {
  const enforcement = { tenant: ALLOW, user: { ...ALLOW, teamsMessages: 'block' }, effective: {} };
  mockGetOutboundEnforcement.mockResolvedValue(enforcement);
  const res = await get({ userId: TARGET });
  expect(res.jsonBody).toEqual({ userId: TARGET, ...enforcement });
});

it('POST scope=tenant writes the channels given and logs the change', async () => {
  const res = await post({ scope: 'tenant', calendarInvites: 'internal', teamsMessages: 'block' });
  expect(res.status).toBe(200);
  expect(mockSetOutboundPolicy).toHaveBeenCalledWith(TENANT, { scope: 'tenant' }, { calendarInvites: 'internal', teamsMessages: 'block' }, ADMIN);
  expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
    operation: 'set_outbound_policy', resource: 'tenant', reason: 'calendarInvites=internal teamsMessages=block',
    before: JSON.stringify(ALLOW),
    after: JSON.stringify({ calendarInvites: 'internal', eventResponses: 'allow', teamsMessages: 'block' }),
  }));
});

it('POST scope=user writes that user row', async () => {
  const res = await post({ scope: 'user', userId: TARGET, eventResponses: 'block' });
  expect(res.status).toBe(200);
  expect(mockSetOutboundPolicy).toHaveBeenCalledWith(TENANT, { scope: 'user', userId: TARGET }, { eventResponses: 'block' }, ADMIN);
  expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ resource: `user:${TARGET}` }));
});

it.each([
  [{ scope: 'org', teamsMessages: 'block' }, 'scope'],
  [{ scope: 'user', teamsMessages: 'block' }, 'userId'],
  [{ scope: 'tenant', teamsMessages: 'hold' }, 'teamsMessages must be one of'],
  [{ scope: 'tenant', mail: 'block' }, 'Unknown field'],
  [{ scope: 'tenant' }, 'at least one'],
])('POST rejects %j', async (body, message) => {
  const res = await post(body);
  expect(res.status).toBe(400);
  expect((res.jsonBody as { error: string }).error).toContain(message);
  expect(mockSetOutboundPolicy).not.toHaveBeenCalled();
});
