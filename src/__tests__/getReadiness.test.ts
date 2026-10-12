/**
 * Tests for the GET /api/manage/readiness admin endpoint.
 *
 * Covers:
 *   - registered as a GET on api/manage/readiness
 *   - no console session is a 401
 *   - a signed-in non-admin is a 403, audited as a refusal
 *   - an admin gets the readiness report, uncached
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { AuthResult } from '../services/authMiddleware.js';

const mockAuthenticateRequest = jest.fn<(req: HttpRequest) => Promise<AuthResult | null>>();
const mockCheckGlobalAdmin = jest.fn<(userId: string) => Promise<boolean>>();
const mockAuditAdminRefusal = jest.fn<(operation: string) => void>();

jest.mock('../services/authMiddleware.js', () => ({
  authenticateConsoleRequest: (req: unknown) => mockAuthenticateRequest(req as HttpRequest),
  authorizeAdmin: async (auth: unknown, operation: unknown) => {
    const isAdmin = await mockCheckGlobalAdmin((auth as AuthResult).userId);
    if (!isAdmin) mockAuditAdminRefusal(operation as string);
    return isAdmin;
  },
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/admin/getReadiness.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{
    status: number;
    jsonBody?: unknown;
    headers?: Record<string, string>;
  }>;
  route?: string;
  methods?: string[];
  authLevel?: string;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const reg = httpMock.mock.calls.find((c) => c[0] === 'getReadiness');
if (!reg) throw new Error('getReadiness handler was not registered');
const registration = reg[1];
const handler = registration.handler;

const AUTH = { userId: 'admin-user', session: { userId: 'admin-user', tenantId: 'tenant-a' } } as unknown as AuthResult;
const ctx = { error: jest.fn() } as unknown as InvocationContext;
const req = { method: 'GET', headers: new Map<string, string>() } as unknown as HttpRequest;

const FLAGS = ['MCP_ENVELOPE_REQUIRE_AAD', 'MCP_SESSION_REQUIRE_IDENTITY_BINDING'] as const;
const saved: Record<string, string | undefined> = {};

beforeEach(() => {
  jest.clearAllMocks();
  mockAuthenticateRequest.mockResolvedValue(AUTH);
  mockCheckGlobalAdmin.mockResolvedValue(true);
  for (const flag of FLAGS) {
    saved[flag] = process.env[flag];
    delete process.env[flag];
  }
});

afterEach(() => {
  for (const flag of FLAGS) {
    const value = saved[flag];
    if (value === undefined) delete process.env[flag];
    else process.env[flag] = value;
  }
});

describe('GET /api/manage/readiness', () => {
  it('is registered as an anonymous GET on api/manage/readiness', () => {
    expect(registration.route).toBe('api/manage/readiness');
    expect(registration.methods).toEqual(['GET']);
    expect(registration.authLevel).toBe('anonymous');
  });

  it('refuses a request with no console session', async () => {
    mockAuthenticateRequest.mockResolvedValue(null);
    const res = await handler(req, ctx);
    expect(res.status).toBe(401);
  });

  it('refuses a non-admin and audits the refusal', async () => {
    mockCheckGlobalAdmin.mockResolvedValue(false);
    const res = await handler(req, ctx);
    expect(res.status).toBe(403);
    expect(mockAuditAdminRefusal).toHaveBeenCalledWith('admin.readiness.read');
  });

  it('returns the readiness report to an admin, uncached', async () => {
    process.env.MCP_ENVELOPE_REQUIRE_AAD = 'true';
    const res = await handler(req, ctx);
    expect(res.status).toBe(200);
    expect(res.headers?.['Cache-Control']).toBe('no-store');
    const body = res.jsonBody as { ready: boolean; checks: Array<{ id: string; enforced: boolean; value: string | null }> };
    expect(body.ready).toBe(false);
    expect(body.checks).toEqual([
      expect.objectContaining({ id: 'envelope-row-binding', enforced: true, value: 'true' }),
      expect.objectContaining({ id: 'session-identity-binding', enforced: false, value: null }),
    ]);
  });
});
