/**
 * Audit rows written by authMiddleware itself: a refused admin check, and the
 * silent session renewal authenticateRequest performs when a session has been
 * idle past its TTL.
 */

import { jest } from '@jest/globals';
import type { HttpRequest } from '@azure/functions';
import type { UserSession } from '../services/tokenCache.js';

const mockLogAccess = jest.fn<(entry: Record<string, unknown>) => void>();
const mockGetSessionByToken = jest.fn<(token: string) => Promise<UserSession | undefined>>();
const mockAcquireTokenSilent = jest.fn<() => Promise<unknown>>();
const mockGraphGet = jest.fn<() => Promise<unknown>>();

const TTL_MS = 7 * 24 * 60 * 60 * 1000;

jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (entry: unknown) => mockLogAccess(entry as Record<string, unknown>),
}));

jest.mock('../services/tokenCache.js', () => ({
  getSessionByToken: (token: string) => mockGetSessionByToken(token),
  SESSION_TTL_MS: 7 * 24 * 60 * 60 * 1000,
  storeSession: async () => undefined,
  isAbsoluteLifetimeExceeded: () => false,
  getValidAccessToken: async () => 'graph-token',
}));

jest.mock('../services/graphClient.js', () => ({
  acquireTokenSilent: () => mockAcquireTokenSilent(),
  createGraphClient: () => ({ api: () => ({ select: () => ({ get: () => mockGraphGet() }) }) }),
}));

import { authenticateRequest, authorizeAdmin } from '../services/authMiddleware.js';

const SESSION: UserSession = {
  userId: 'user-1',
  homeAccountId: 'user-1.tenant-abc',
  displayName: 'Adele Vance',
  email: 'adele@fabrikam.com',
  tenantId: 'tenant-abc',
  accessToken: 'a',
  expiresAt: Date.now(),
  sessionToken: 'tok',
  sessionCreatedAt: Date.now() - TTL_MS - 60_000,
  sessionAbsoluteCreatedAt: Date.now() - TTL_MS - 60_000,
};

const req = {
  headers: new Map<string, string>([['authorization', 'Bearer tok']]),
} as unknown as HttpRequest;

beforeEach(() => {
  jest.clearAllMocks();
  mockGetSessionByToken.mockResolvedValue({ ...SESSION });
});

describe('session renewal inside authenticateRequest', () => {
  it('records a renewal', async () => {
    mockAcquireTokenSilent.mockResolvedValue({ accessToken: 'new', expiresOn: new Date() });
    expect(await authenticateRequest(req)).not.toBeNull();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      tenantId: 'tenant-abc', userId: 'user-1', operation: 'auth.session_renew', result: 'allowed',
    }));
  });

  it('records a failed renewal as denied', async () => {
    mockAcquireTokenSilent.mockRejectedValue(new Error('interaction_required'));
    expect(await authenticateRequest(req)).toBeNull();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      operation: 'auth.session_renew', result: 'denied',
    }));
  });

  it('records nothing for an active session', async () => {
    mockGetSessionByToken.mockResolvedValue({ ...SESSION, sessionCreatedAt: Date.now() });
    expect(await authenticateRequest(req)).not.toBeNull();
    expect(mockLogAccess).not.toHaveBeenCalled();
  });
});

describe('authorizeAdmin', () => {
  const auth = { userId: 'user-1', session: SESSION };

  it('records a refusal under the operation that was refused', async () => {
    mockGraphGet.mockResolvedValue({ value: [] });
    expect(await authorizeAdmin(auth, 'policy.services.set', 'tenant')).toBe(false);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      tenantId: 'tenant-abc',
      userId: 'user-1',
      userEmail: 'adele@fabrikam.com',
      operation: 'policy.services.set',
      resource: 'tenant',
      result: 'denied',
      reason: 'Global Administrator role required',
    }));
  });

  it('records nothing for a Global Administrator', async () => {
    mockGraphGet.mockResolvedValue({ value: [{ roleTemplateId: '62e90394-69f5-4237-9190-012177145e10' }] });
    expect(await authorizeAdmin(auth, 'policy.services.set')).toBe(true);
    expect(mockLogAccess).not.toHaveBeenCalled();
  });

  it('records a refusal when the check itself fails (fail closed)', async () => {
    mockGraphGet.mockRejectedValue(new Error('graph down'));
    expect(await authorizeAdmin(auth, 'admin.audit_log.read')).toBe(false);
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ operation: 'admin.audit_log.read', result: 'denied' }));
  });
});
