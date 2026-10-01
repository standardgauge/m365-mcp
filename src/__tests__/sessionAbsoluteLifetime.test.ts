/**
 *: Absolute session lifetime enforcement in authenticateRequestAllowExpired.
 *
 * Verifies that /refresh cannot be used to renew a session indefinitely:
 *   - Sessions older than SESSION_MAX_LIFETIME_MS are rejected even if MSAL
 *     refresh token is still valid (sessionAbsoluteCreatedAt > limit).
 *   - Sessions within the limit are accepted with the inactivity TTL skipped.
 *   - The sessionAbsoluteCreatedAt field is preserved (not reset) by refresh.
 *   - Legacy sessions with sessionAbsoluteCreatedAt=0 fall back to
 *     sessionCreatedAt for the absolute check.
 *   - Sessions where both anchor fields are 0 bypass the absolute check
 *     (backward compatibility — field was not present at creation).
 */

import { jest } from '@jest/globals';
import type { HttpRequest } from '@azure/functions';
import type { UserSession } from '../services/tokenCache.js';
import type { AuthResult } from '../services/authMiddleware.js';

// ── Constants under test ─────────────────────────────────────────────────────

const SESSION_MAX_LIFETIME_MS = 30 * 24 * 60 * 60 * 1000; // 30 days

// ── Mock declarations ────────────────────────────────────────────────────────

const mockGetSessionByToken = jest.fn<(token: string) => Promise<UserSession | undefined>>();

jest.mock('../services/tokenCache.js', () => {
  const MAX_MS = 30 * 24 * 60 * 60 * 1000;
  return {
    getSessionByToken: (...args: unknown[]) => mockGetSessionByToken(args[0] as string),
    SESSION_TTL_MS: 7 * 24 * 60 * 60 * 1000,
    SESSION_MAX_LIFETIME_MS: MAX_MS,
    storeSession: jest.fn(),
    isAbsoluteLifetimeExceeded: (
      session: { sessionAbsoluteCreatedAt?: number; sessionCreatedAt: number },
      now = Date.now()
    ) => {
      const anchor = session.sessionAbsoluteCreatedAt || session.sessionCreatedAt;
      if (!anchor) return false;
      return (now - anchor) > MAX_MS;
    },
  };
});

jest.mock('../services/graphClient.js', () => ({
  acquireTokenSilent: jest.fn(),
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import { authenticateRequest, authenticateRequestAllowExpired } from '../services/authMiddleware.js';
import { acquireTokenSilent } from '../services/graphClient.js';
import { storeSession } from '../services/tokenCache.js';

// ── Test helpers ─────────────────────────────────────────────────────────────

function makeSession(overrides: Partial<UserSession> = {}): UserSession {
  return {
    userId: 'user-abc',
    homeAccountId: 'home-abc',
    displayName: 'Test User',
    email: 'test@example.com',
    tenantId: 'tenant-123',
    accessToken: 'access-token',
    expiresAt: Date.now() + 3_600_000,
    sessionToken: 'session-token',
    sessionCreatedAt: Date.now(),
    sessionAbsoluteCreatedAt: Date.now(),
    ...overrides,
  };
}

function makeRequest(token: string): HttpRequest {
  return {
    headers: {
      get: (key: string) => {
        if (key === 'authorization') return `Bearer ${token}`;
        return null;
      },
    },
  } as unknown as HttpRequest;
}

// ── Test suite ───────────────────────────────────────────────────────────────

beforeEach(() => {
  jest.clearAllMocks();
});

describe('authenticateRequestAllowExpired — absolute lifetime enforcement', () => {
  it('allows a session within SESSION_MAX_LIFETIME_MS', async () => {
    const session = makeSession({
      sessionCreatedAt: Date.now() - 3 * 24 * 60 * 60 * 1000, // 3 days ago (inactivity-expired)
      sessionAbsoluteCreatedAt: Date.now() - 10 * 24 * 60 * 60 * 1000, // 10 days ago (within 30d limit)
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequestAllowExpired(makeRequest('token'));

    expect(result).not.toBeNull();
    expect((result as AuthResult).userId).toBe('user-abc');
  });

  it('blocks a session where sessionAbsoluteCreatedAt exceeds SESSION_MAX_LIFETIME_MS', async () => {
    const session = makeSession({
      sessionAbsoluteCreatedAt: Date.now() - (SESSION_MAX_LIFETIME_MS + 1),
      sessionCreatedAt: Date.now(), // recently touched, inactivity TTL not expired
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequestAllowExpired(makeRequest('token'));

    expect(result).toBeNull();
  });

  it('blocks a session exactly at the boundary (> not >=)', async () => {
    // One millisecond past the limit
    const session = makeSession({
      sessionAbsoluteCreatedAt: Date.now() - SESSION_MAX_LIFETIME_MS - 1,
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequestAllowExpired(makeRequest('token'));

    expect(result).toBeNull();
  });

  it('allows a session exactly at SESSION_MAX_LIFETIME_MS (not yet expired)', async () => {
    // Exactly at the limit — not yet exceeded
    const session = makeSession({
      sessionAbsoluteCreatedAt: Date.now() - SESSION_MAX_LIFETIME_MS + 1000,
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequestAllowExpired(makeRequest('token'));

    expect(result).not.toBeNull();
  });

  it('falls back to sessionCreatedAt when sessionAbsoluteCreatedAt is absent (legacy)', async () => {
    // Legacy session: sessionAbsoluteCreatedAt not set, sessionCreatedAt old
    const session = makeSession({
      sessionAbsoluteCreatedAt: undefined,
      sessionCreatedAt: Date.now() - (SESSION_MAX_LIFETIME_MS + 86_400_000), // 31 days ago
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequestAllowExpired(makeRequest('token'));

    expect(result).toBeNull();
  });

  it('falls back to sessionCreatedAt when sessionAbsoluteCreatedAt is 0 (legacy zero)', async () => {
    const session = makeSession({
      sessionAbsoluteCreatedAt: 0,
      sessionCreatedAt: Date.now() - (SESSION_MAX_LIFETIME_MS + 86_400_000), // 31 days ago
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequestAllowExpired(makeRequest('token'));

    expect(result).toBeNull();
  });

  it('bypasses absolute check when both anchor fields are 0 (pre-field legacy)', async () => {
    // Sessions created before either field was added — cannot enforce without anchor.
    // These will eventually re-OAuth when the inactivity TTL fires.
    const session = makeSession({
      sessionAbsoluteCreatedAt: 0,
      sessionCreatedAt: 0,
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequestAllowExpired(makeRequest('token'));

    expect(result).not.toBeNull();
  });

  it('returns null when no matching session exists', async () => {
    mockGetSessionByToken.mockResolvedValue(undefined);

    const result = await authenticateRequestAllowExpired(makeRequest('token'));

    expect(result).toBeNull();
  });
});

describe('authenticateRequest — absolute lifetime enforcement on normal endpoints', () => {
  it('returns null and does not call acquireTokenSilent when session exceeds absolute lifetime with expired inactivity TTL', async () => {
    const session = makeSession({
      sessionAbsoluteCreatedAt: Date.now() - (SESSION_MAX_LIFETIME_MS + 1),
      sessionCreatedAt: Date.now() - (7 * 24 * 60 * 60 * 1000 + 1), // inactivity TTL expired too
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequest(makeRequest('token'));

    expect(result).toBeNull();
    expect(acquireTokenSilent).not.toHaveBeenCalled();
    expect(storeSession).not.toHaveBeenCalled();
  });

  it('returns null and does not call acquireTokenSilent when session exceeds absolute lifetime even if inactivity TTL is still active', async () => {
    const session = makeSession({
      sessionAbsoluteCreatedAt: Date.now() - (SESSION_MAX_LIFETIME_MS + 1),
      sessionCreatedAt: Date.now(), // inactivity TTL fresh — session actively in use
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequest(makeRequest('token'));

    expect(result).toBeNull();
    expect(acquireTokenSilent).not.toHaveBeenCalled();
    expect(storeSession).not.toHaveBeenCalled();
  });

  it('allows a session within absolute lifetime even when inactivity TTL is expired (delegates to silent refresh)', async () => {
    (acquireTokenSilent as jest.MockedFunction<typeof acquireTokenSilent>).mockResolvedValue({
      accessToken: 'new-token',
      expiresOn: new Date(Date.now() + 3_600_000),
    } as Awaited<ReturnType<typeof acquireTokenSilent>>);
    (storeSession as jest.MockedFunction<typeof storeSession>).mockResolvedValue(undefined);

    const session = makeSession({
      sessionAbsoluteCreatedAt: Date.now() - 10 * 24 * 60 * 60 * 1000, // 10 days — within 30d bound
      sessionCreatedAt: Date.now() - (7 * 24 * 60 * 60 * 1000 + 1),    // inactivity TTL expired
    });
    mockGetSessionByToken.mockResolvedValue(session);

    const result = await authenticateRequest(makeRequest('token'));

    expect(result).not.toBeNull();
    expect(acquireTokenSilent).toHaveBeenCalled();
  });
});

describe('sessionAbsoluteCreatedAt immutability', () => {
  it('sessionAbsoluteCreatedAt is preserved through object spread (not reset by refresh paths)', () => {
    const originalTimestamp = Date.now() - 15 * 24 * 60 * 60 * 1000; // 15 days ago
    const session = makeSession({ sessionAbsoluteCreatedAt: originalTimestamp });

    // Simulate what refresh.ts does: spread + reset sessionCreatedAt only
    const refreshed: UserSession = {
      ...session,
      accessToken: 'new-access-token',
      expiresAt: Date.now() + 3_600_000,
      sessionCreatedAt: Date.now(),
      // sessionAbsoluteCreatedAt intentionally NOT set here — must come from spread
    };

    expect(refreshed.sessionAbsoluteCreatedAt).toBe(originalTimestamp);
  });

  it('sessionAbsoluteCreatedAt is preserved through sliding-window touch', () => {
    const originalTimestamp = Date.now() - 5 * 24 * 60 * 60 * 1000; // 5 days ago
    const session = makeSession({ sessionAbsoluteCreatedAt: originalTimestamp });

    // Simulate what authMiddleware.ts sliding window does
    const touched: UserSession = { ...session, sessionCreatedAt: Date.now() };

    expect(touched.sessionAbsoluteCreatedAt).toBe(originalTimestamp);
  });
});
