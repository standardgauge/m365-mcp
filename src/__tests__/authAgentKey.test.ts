/**
 * authenticateRequest with an agent API key as the bearer.
 *
 * Covers:
 *   - a key no row matches is 'unknown' to the candidate loop, so the request is unauthenticated
 *   - revoked, expired and tenant-disabled keys are refused and audited
 *   - a key whose session row is gone is refused
 *   - a live key resolves to its bound session with the key context attached,
 *     the policy intersected with the tenant's, and the cached session untouched
 *   - last-use is recorded at most every AGENT_KEY_TOUCH_INTERVAL_MS
 *   - the key never passes authenticateRequestAllowExpired (the refresh path)
 *   - a storage failure on the key lookup is a 503-class error, not a 401
 *   - a live key marks the request scope with its id, so the deny list reads
 *     the key's own entries; outside a request scope the key is refused
 */
import { jest } from '@jest/globals';
import { randomBytes } from 'crypto';
import type { HttpRequest } from '@azure/functions';
import type { UserSession } from '../services/tokenCache.js';
import type { AgentKeyRecord, TenantAgentKeyPolicy } from '../services/agentKeys.js';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

const mockGetSessionByToken = jest.fn<(t: string) => Promise<UserSession | undefined>>();
const mockGetSessionByStorageKey = jest.fn<(k: string) => Promise<UserSession | undefined>>();
const mockStoreSession = jest.fn<(s: UserSession) => Promise<void>>();
const mockLoadKey = jest.fn<(t: string) => Promise<AgentKeyRecord | null>>();
const mockTenantPolicy = jest.fn<(t: string) => Promise<TenantAgentKeyPolicy>>();
const mockTouch = jest.fn<(r: AgentKeyRecord) => Promise<void>>();
const mockLogAccess = jest.fn<(e: Record<string, unknown>) => void>();

jest.mock('../services/tokenCache.js', () => ({
  getSessionByToken: (t: unknown) => mockGetSessionByToken(t as string),
  getSessionByStorageKey: (k: unknown) => mockGetSessionByStorageKey(k as string),
  storeSession: (s: unknown) => mockStoreSession(s as UserSession),
  getValidAccessToken: async () => 'at',
  isAbsoluteLifetimeExceeded: (s: UserSession, now = Date.now(), maxLifetimeMs = 30 * 24 * 60 * 60 * 1000) => {
    const anchor = s.sessionAbsoluteCreatedAt || s.sessionCreatedAt;
    return !anchor || now - anchor > maxLifetimeMs;
  },
  SESSION_TTL_MS: 7 * 24 * 60 * 60 * 1000,
}));
jest.mock('../services/agentKeys.js', () => ({
  ...jest.requireActual<object>('../services/agentKeys.js'),
  loadAgentKeyByToken: (t: unknown) => mockLoadKey(t as string),
  getTenantAgentKeyPolicy: (t: unknown) => mockTenantPolicy(t as string),
  touchAgentKey: (r: unknown) => mockTouch(r as AgentKeyRecord),
}));
jest.mock('../services/graphClient.js', () => ({
  acquireTokenSilent: async () => ({ accessToken: 'new', expiresOn: new Date(Date.now() + 3_600_000) }),
  createGraphClient: () => ({ api: () => ({ select: () => ({ get: async () => ({ value: [] }) }) }) }),
}));
jest.mock('../services/auditLog.js', () => ({
  ...jest.requireActual<object>('../services/auditLog.js'),
  logAccess: (e: unknown) => mockLogAccess(e as Record<string, unknown>),
}));

import { authenticateRequest, authenticateRequestAllowExpired } from '../services/authMiddleware.js';
import { SessionStoreUnavailableError } from '../services/sessionStoreError.js';
import { AGENT_KEY_TOUCH_INTERVAL_MS, generateAgentKeyToken } from '../services/agentKeys.js';
import { currentAgentKeyId, runWithClientAddress } from '../services/clientAddress.js';

const TENANT = 'tenant-a';
const USER = 'user-1';
const KEY = generateAgentKeyToken();

function session(over: Partial<UserSession> = {}): UserSession {
  return {
    userId: USER, homeAccountId: `${USER}.${TENANT}`, displayName: 'Adele Vance', email: 'adele@fabrikam.com',
    tenantId: TENANT, accessToken: 'at', expiresAt: Date.now() + 3_600_000, sessionToken: '',
    sessionCreatedAt: Date.now(), sessionAbsoluteCreatedAt: Date.now(), kind: 'client',
    deviceLabel: 'agent-key:nightly#abc123def456', _storageKey: 'row-key-1', ...over,
  };
}
function record(over: Partial<AgentKeyRecord> = {}): AgentKeyRecord {
  return {
    keyId: 'abc123def456', tenantId: TENANT, userId: USER, userEmail: 'adele@fabrikam.com', label: 'nightly',
    sessionStorageKey: 'row-key-1',
    policy: { services: ['mail', 'sharepoint'], readOnlyServices: ['mail'], allowedSites: [] },
    createdAt: Date.now() - 1000, expiresAt: Date.now() + 90 * 24 * 60 * 60 * 1000, lastUsedAt: 0, revokedAt: 0, _storageKey: 'k-row', ...over,
  };
}
function req(token: string): HttpRequest {
  return { method: 'POST', headers: new Map([['authorization', `Bearer ${token}`]]) } as unknown as HttpRequest;
}

/** authenticateRequest inside a request scope, as withSecurity runs every route. */
function authenticate(request: HttpRequest) {
  return runWithClientAddress(request, () => authenticateRequest(request));
}

beforeEach(() => {
  jest.clearAllMocks();
  mockGetSessionByToken.mockResolvedValue(undefined);
  mockGetSessionByStorageKey.mockResolvedValue(session());
  mockStoreSession.mockResolvedValue();
  mockLoadKey.mockResolvedValue(record());
  mockTenantPolicy.mockResolvedValue({ allowKeys: true, disabledServices: [], readOnlyServices: [], maxKeyLifetimeDays: 365, ownersMayWiden: true });
  mockTouch.mockResolvedValue();
});

describe('authenticateRequest with an agent key', () => {
  it('resolves a live key to its session with the key context attached', async () => {
    const auth = await authenticate(req(KEY));
    expect(auth).not.toBeNull();
    expect(auth!.userId).toBe(USER);
    expect(auth!.session.agentKey).toEqual({
      keyId: 'abc123def456', label: 'nightly',
      policy: { services: ['mail', 'sharepoint'], readOnlyServices: ['mail'], allowedSites: [] },
    });
    expect(mockLoadKey).toHaveBeenCalledWith(KEY);
    expect(mockGetSessionByStorageKey).toHaveBeenCalledWith('row-key-1');
    // The session-token table is never consulted for a key.
    expect(mockGetSessionByToken).not.toHaveBeenCalled();
  });

  it('does not attach the context to the cached session object', async () => {
    const cached = session();
    mockGetSessionByStorageKey.mockResolvedValue(cached);
    const auth = await authenticate(req(KEY));
    expect(auth!.session).not.toBe(cached);
    expect(cached.agentKey).toBeUndefined();
  });

  it("applies the tenant's key policy on top of the key's own", async () => {
    mockTenantPolicy.mockResolvedValue({ allowKeys: true, disabledServices: ['sharepoint'], readOnlyServices: ['mail', 'calendar'], maxKeyLifetimeDays: 365, ownersMayWiden: true });
    const auth = await authenticate(req(KEY));
    expect(auth!.session.agentKey!.policy).toEqual({ services: ['mail'], readOnlyServices: ['mail'], allowedSites: [] });
  });

  it('is unauthenticated when no key row matches', async () => {
    mockLoadKey.mockResolvedValue(null);
    expect(await authenticate(req(KEY))).toBeNull();
    expect(mockLogAccess).not.toHaveBeenCalled();
  });

  it.each([
    ['revoked', { revokedAt: Date.now() - 5 }, 'key revoked'],
    ['expired', { expiresAt: Date.now() - 5 }, 'key expired'],
  ])('refuses a %s key and audits the refusal', async (_name, over, reason) => {
    mockLoadKey.mockResolvedValue(record(over));
    expect(await authenticate(req(KEY))).toBeNull();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      operation: 'auth.agent_key', result: 'denied', reason, resource: 'abc123def456', userId: USER, tenantId: TENANT,
    }));
    expect(mockGetSessionByStorageKey).not.toHaveBeenCalled();
  });

  it('refuses every key when the tenant has disabled them', async () => {
    mockTenantPolicy.mockResolvedValue({ allowKeys: false, disabledServices: [], readOnlyServices: [], maxKeyLifetimeDays: 365, ownersMayWiden: true });
    expect(await authenticate(req(KEY))).toBeNull();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ result: 'denied', reason: expect.stringContaining('disabled') }));
  });

  it('refuses a key whose session row is gone', async () => {
    mockGetSessionByStorageKey.mockResolvedValue(undefined);
    expect(await authenticate(req(KEY))).toBeNull();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({ result: 'denied', reason: expect.stringContaining('no longer exists') }));
  });

  it('refuses a key whose session belongs to someone else', async () => {
    mockGetSessionByStorageKey.mockResolvedValue(session({ userId: 'someone-else' }));
    expect(await authenticate(req(KEY))).toBeNull();
  });

  it("caps the bound session by the key's own lifetime, not the 30-day sign-in cap", async () => {
    const day = 24 * 60 * 60 * 1000;
    // A 90-day key, 45 days in: past the 30-day cap a sign-in would get, still inside its own.
    const created = Date.now() - 45 * day;
    mockLoadKey.mockResolvedValue(record({ createdAt: created, expiresAt: created + 90 * day }));
    mockGetSessionByStorageKey.mockResolvedValue(session({ sessionCreatedAt: Date.now() - day, sessionAbsoluteCreatedAt: created }));
    expect(await authenticate(req(KEY))).not.toBeNull();

    // The same key with its session's anchor somehow older than the key's lifetime is refused.
    mockGetSessionByStorageKey.mockResolvedValue(session({ sessionCreatedAt: Date.now() - day, sessionAbsoluteCreatedAt: Date.now() - 100 * day }));
    expect(await authenticate(req(KEY))).toBeNull();
  });

  it('records last use, but not on every call', async () => {
    await authenticate(req(KEY));
    expect(mockTouch).toHaveBeenCalledTimes(1);
    mockLoadKey.mockResolvedValue(record({ lastUsedAt: Date.now() - AGENT_KEY_TOUCH_INTERVAL_MS / 2 }));
    await authenticate(req(KEY));
    expect(mockTouch).toHaveBeenCalledTimes(1);
  });

  it('surfaces a key-store outage as SessionStoreUnavailableError', async () => {
    mockLoadKey.mockRejectedValue(new Error('503 from storage'));
    await expect(authenticate(req(KEY))).rejects.toBeInstanceOf(SessionStoreUnavailableError);
  });

  it('marks the request scope with the key id', async () => {
    const r = req(KEY);
    const seen = await runWithClientAddress(r, async () => {
      expect(currentAgentKeyId()).toBeUndefined();
      await authenticateRequest(r);
      return currentAgentKeyId();
    });
    expect(seen).toBe('abc123def456');
  });

  it('does not mark the scope for a key it refuses', async () => {
    mockLoadKey.mockResolvedValue(record({ revokedAt: Date.now() - 5 }));
    const r = req(KEY);
    const seen = await runWithClientAddress(r, async () => {
      await authenticateRequest(r);
      return currentAgentKeyId();
    });
    expect(seen).toBeUndefined();
  });

  it("refuses a key outside a request scope, where its deny entries could not apply", async () => {
    expect(await authenticateRequest(req(KEY))).toBeNull();
    expect(mockLogAccess).toHaveBeenCalledWith(expect.objectContaining({
      operation: 'auth.agent_key', result: 'denied', reason: expect.stringContaining('request scope'),
    }));
  });

  it('is not a credential for the refresh path', async () => {
    expect(await authenticateRequestAllowExpired(req(KEY))).toBeNull();
    expect(mockLoadKey).not.toHaveBeenCalled();
  });
});
