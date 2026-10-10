/**
 * Tests for the /api/manage/services admin endpoint, focused on the
 * readOnlyServices control added for calendar read-only mode.
 *
 * Covers:
 *   - GET returns both enabledServices and readOnlyServices
 *   - POST (admin) sets readOnlyServices
 *   - POST (admin) can update enabledServices and readOnlyServices independently
 *   - Non-admins are refused (403)
 *   - Empty body is a 400
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockAuditAdminRefusal = jest.fn<(operation: string) => void>();
const mockLogAccess = jest.fn<(entry: Record<string, unknown>) => void>();
const mockGetTenantId = jest.fn<(userId: string) => Promise<string>>();
const mockGetEnabledServices = jest.fn<(t: string) => Promise<string[]>>();
const mockSetEnabledServices = jest.fn<(t: string, s: string[]) => Promise<void>>();
const mockGetReadOnlyServices = jest.fn<(t: string) => Promise<string[]>>();
const mockSetReadOnlyServices = jest.fn<(t: string, s: string[]) => Promise<void>>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateConsoleRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  checkGlobalAdmin: (userId: unknown) => mockCheckGlobalAdmin(userId as string),
  authorizeAdmin: async (auth: unknown, operation: unknown) => {
    const isAdmin = await mockCheckGlobalAdmin((auth as AuthResult).userId);
    if (!isAdmin) mockAuditAdminRefusal(operation as string);
    return isAdmin;
  },
}));

jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (entry: unknown) => mockLogAccess(entry as Record<string, unknown>),
}));

jest.mock('../services/tokenCache.js', () => ({
  getTenantId: (userId: unknown) => mockGetTenantId(userId as string),
}));

jest.mock('../services/serviceSettings.js', () => ({
  getEnabledServices: (t: unknown) => mockGetEnabledServices(t as string),
  setEnabledServices: (t: unknown, s: unknown) => mockSetEnabledServices(t as string, s as string[]),
  getReadOnlyServices: (t: unknown) => mockGetReadOnlyServices(t as string),
  setReadOnlyServices: (t: unknown, s: unknown) => mockSetReadOnlyServices(t as string, s as string[]),
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/admin/manageServices.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{ status: number; jsonBody?: unknown }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const reg = httpMock.mock.calls.find((c) => c[0] === 'manageServices');
if (!reg) throw new Error('manageServices handler was not registered');
const handler = reg[1].handler;

const TENANT = 'tenant-a';
const ADMIN = 'admin-user';
const AUTH = { userId: ADMIN, session: { userId: ADMIN, tenantId: TENANT, email: 'a@example.com' } } as unknown as AuthResult;
const ctx = { error: jest.fn() } as unknown as InvocationContext;

function req(method: string, body?: unknown): HttpRequest {
  return { method, json: async () => body ?? {}, headers: new Map<string, string>() } as unknown as HttpRequest;
}

beforeEach(() => {
  jest.clearAllMocks();
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockCheckGlobalAdmin.mockResolvedValue(true);
  mockGetTenantId.mockResolvedValue(TENANT);
  mockGetEnabledServices.mockResolvedValue(['mail', 'calendar']);
  mockGetReadOnlyServices.mockResolvedValue(['calendar']);
});

describe('manageServices — readOnlyServices', () => {
  it('GET returns both enabledServices and readOnlyServices', async () => {
    const res = await handler(req('GET'), ctx);
    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({ enabledServices: ['mail', 'calendar'], readOnlyServices: ['calendar'] });
  });

  it('POST sets readOnlyServices for the tenant', async () => {
    const res = await handler(req('POST', { readOnlyServices: ['calendar'] }), ctx);
    expect(res.status).toBe(200);
    expect(mockSetReadOnlyServices).toHaveBeenCalledWith(TENANT, ['calendar']);
    expect(mockSetEnabledServices).not.toHaveBeenCalled();
  });

  it('POST audits each list it changes with the value before and after', async () => {
    await handler(req('POST', { enabledServices: ['mail'], readOnlyServices: [] }), ctx);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      tenantId: TENANT,
      operation: 'policy.services.set',
      resource: 'tenant',
      result: 'allowed',
      before: '["mail","calendar"]',
      after: '["mail"]',
    }));
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      operation: 'policy.read_only.set',
      before: '["calendar"]',
      after: '[]',
    }));
  });

  it('POST can update both lists in one call', async () => {
    await handler(req('POST', { enabledServices: ['mail', 'calendar'], readOnlyServices: [] }), ctx);
    expect(mockSetEnabledServices).toHaveBeenCalledWith(TENANT, ['mail', 'calendar']);
    expect(mockSetReadOnlyServices).toHaveBeenCalledWith(TENANT, []);
  });

  it('POST is refused for non-admins with 403', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(false);
    const res = await handler(req('POST', { readOnlyServices: ['calendar'] }), ctx);
    expect(res.status).toBe(403);
    expect(mockSetReadOnlyServices).not.toHaveBeenCalled();
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('policy.services.set');
    expect(mockLogAccess).not.toHaveBeenCalled();
  });

  it('POST with neither field returns 400', async () => {
    const res = await handler(req('POST', {}), ctx);
    expect(res.status).toBe(400);
  });

  it('POST rejects a non-array readOnlyServices with 400', async () => {
    const res = await handler(req('POST', { readOnlyServices: 'calendar' }), ctx);
    expect(res.status).toBe(400);
    expect(mockSetReadOnlyServices).not.toHaveBeenCalled();
  });
});
