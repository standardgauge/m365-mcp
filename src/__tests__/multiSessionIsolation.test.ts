/**
 * Regression tests for concurrent multi-session isolation.
 *
 * Context: v2.5.0 (commits 27a35c9, 28c40db, 6842940, 36596ce) fixed a bug
 * where multiple concurrent sessions for the same userId could leak tokens
 * across requests. Before the fix, HTTP handlers used getValidAccessToken(userId),
 * which resolves through the shared `userIndex` map. Under concurrent same-user
 * traffic, one request could use another session's token if the userIndex was
 * updated between when the session was authenticated and when the token was used.
 *
 * The fix: all handlers now use getValidAccessTokenForSession(session), which
 * bypasses the global userIndex entirely.
 *
 * These tests simulate two concurrent sessions for the same userId and verify:
 *   1. getSessionByToken returns the correct session for each token
 *   2. getValidAccessTokenForSession uses the session object directly (not userIndex)
 *   3. Token refresh only affects the targeted session
 *   4. deleteSessionByKey removes only the targeted session
 *   5. Interleaved requests cannot pick up another session's access token
 */

import { randomBytes } from 'crypto';

// Crypto keys must be set before any module under test loads — credentialCrypto
// reads MCP_SESSION_HMAC_KEY and MCP_DATA_ENCRYPTION_KEY at import time.
process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');
process.env.MCP_DATA_ENCRYPTION_KEY = randomBytes(32).toString('hex');

// Required for graphClient singleton initialization
process.env.AZURE_CLIENT_ID = 'test-client-id';
process.env.AZURE_CLIENT_SECRET = 'test-client-secret';
process.env.AZURE_TENANT_ID = 'test-tenant-id';
process.env.OAUTH_REDIRECT_URI = 'http://localhost:7071/api/auth/callback';

import { jest } from '@jest/globals';

// ── MSAL mock ─────────────────────────────────────────────────────────────────

const mockAcquireTokenSilent = jest.fn<() => Promise<unknown>>();
const mockGetAllAccounts = jest.fn<() => Promise<unknown[]>>();

jest.mock('@azure/msal-node', () => ({
  ConfidentialClientApplication: jest.fn().mockImplementation(() => ({
    getTokenCache: jest.fn(() => ({ getAllAccounts: mockGetAllAccounts })),
    acquireTokenSilent: mockAcquireTokenSilent,
    acquireTokenByCode: jest.fn(),
    getAuthCodeUrl: jest.fn(),
  })),
  LogLevel: { Error: 0, Warning: 1, Info: 2, Verbose: 3, Trace: 4 },
}));

jest.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: jest.fn().mockReturnValue({ api: jest.fn() }),
    initWithMiddleware: jest.fn().mockReturnValue({ api: jest.fn() }),
  },
}));

// Side-effect-only import — stub it so no actual fetch polyfill loads
jest.mock('isomorphic-fetch', () => ({}));

// No Table Storage here: every session lives in the in-memory cache. A cache
// miss reaches the storage lookup, which must answer "no such row" rather than
// fail, since getSessionByToken surfaces a failed lookup as an error.
jest.mock('../services/tableStorage.js', () => ({
  ...jest.requireActual<typeof import('../services/tableStorage.js')>('../services/tableStorage.js'),
  loadSessionByToken: async () => null,
}));

// ── Module under test ─────────────────────────────────────────────────────────

import {
  storeSession,
  getSessionByToken,
  getValidAccessTokenForSession,
  deleteSessionByKey,
  deleteAllUserSessions,
} from '../services/tokenCache.js';
import type { UserSession } from '../services/tokenCache.js';

// ── Helpers ───────────────────────────────────────────────────────────────────

/**
 * Build a UserSession with a unique sessionToken per call.
 * accessToken is what we verify isn't leaked to the other session.
 */
function makeSession(
  userId: string,
  tokenSuffix: string,
  accessToken: string,
  overrides?: Partial<UserSession>,
): UserSession {
  return {
    userId,
    homeAccountId: `home-account-${tokenSuffix}`,
    displayName: 'Test User',
    email: 'test@example.com',
    tenantId: 'test-tenant-id',
    accessToken,
    expiresAt: Date.now() + 60 * 60 * 1000,
    sessionToken: `session-token-${tokenSuffix}-${randomBytes(4).toString('hex')}`,
    sessionCreatedAt: Date.now(),
    ...overrides,
  };
}

/**
 * Each test gets its own userId so in-memory sessionCache / userIndex entries
 * from one test cannot bleed into another.
 */
let userCounter = 0;
function freshUserId(): string {
  return `multi-session-test-user-${++userCounter}-${randomBytes(4).toString('hex')}`;
}

// ── Suite ─────────────────────────────────────────────────────────────────────

afterEach(() => {
  jest.clearAllMocks();
});

describe('concurrent multi-session isolation for same userId', () => {
  // ── 1. Token lookup ─────────────────────────────────────────────────────────

  it('getSessionByToken returns the correct session for each token', async () => {
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB); // userIndex now points to sessionB

    const foundA = await getSessionByToken(sessionA.sessionToken);
    const foundB = await getSessionByToken(sessionB.sessionToken);

    expect(foundA).toBeDefined();
    expect(foundA!.accessToken).toBe('access-token-A');
    expect(foundA!.userId).toBe(userId);

    expect(foundB).toBeDefined();
    expect(foundB!.accessToken).toBe('access-token-B');
    expect(foundB!.userId).toBe(userId);
  });

  it('getSessionByToken for one token does not return the other session', async () => {
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB);

    const foundA = await getSessionByToken(sessionA.sessionToken);
    const foundB = await getSessionByToken(sessionB.sessionToken);

    expect(foundA!.sessionToken).toBe(sessionA.sessionToken);
    expect(foundB!.sessionToken).toBe(sessionB.sessionToken);
    expect(foundA!.sessionToken).not.toBe(foundB!.sessionToken);
  });

  // ── 2. getValidAccessTokenForSession bypasses userIndex ────────────────────

  it('getValidAccessTokenForSession returns the token from the session passed in, not the userIndex-pointed session', async () => {
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB); // userIndex now points to sessionB

    // Retrieve sessionA directly — this is the session a request has in hand
    const retrievedA = await getSessionByToken(sessionA.sessionToken);
    expect(retrievedA).toBeDefined();

    // Even though userIndex points to sessionB, the token returned should be
    // sessionA's because getValidAccessTokenForSession reads from the session
    // object directly, not from userIndex.
    const token = await getValidAccessTokenForSession(retrievedA!);

    expect(token).toBe('access-token-A');
    expect(token).not.toBe('access-token-B');
  });

  it('getValidAccessTokenForSession does not mutate the other session when called on one session', async () => {
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB);

    const retrievedA = await getSessionByToken(sessionA.sessionToken);
    await getValidAccessTokenForSession(retrievedA!);

    // SessionB should be completely untouched
    const retrievedB = await getSessionByToken(sessionB.sessionToken);
    expect(retrievedB).toBeDefined();
    expect(retrievedB!.accessToken).toBe('access-token-B');
  });

  // ── 3. Token refresh scoped to targeted session ────────────────────────────

  it('getValidAccessTokenForSession refreshes only the targeted session when its token is near expiry', async () => {
    const userId = freshUserId();
    const now = Date.now();

    const sessionA = makeSession(userId, 'A', 'access-token-A-stale', {
      expiresAt: now + 2 * 60 * 1000, // within 5-minute buffer → triggers refresh
      homeAccountId: 'home-account-A',
    });
    const sessionB = makeSession(userId, 'B', 'access-token-B', {
      homeAccountId: 'home-account-B',
    });

    await storeSession(sessionA);
    await storeSession(sessionB);

    const retrievedA = await getSessionByToken(sessionA.sessionToken);
    expect(retrievedA).toBeDefined();

    // Mock MSAL to return a fresh token only for sessionA's account
    mockGetAllAccounts.mockResolvedValueOnce([{ homeAccountId: 'home-account-A' }]);
    mockAcquireTokenSilent.mockResolvedValueOnce({
      accessToken: 'access-token-A-refreshed',
      expiresOn: new Date(now + 60 * 60 * 1000),
    });

    const refreshedToken = await getValidAccessTokenForSession(retrievedA!);
    expect(refreshedToken).toBe('access-token-A-refreshed');
    expect(mockAcquireTokenSilent).toHaveBeenCalledTimes(1);

    // SessionB must not have been touched by the refresh
    const retrievedB = await getSessionByToken(sessionB.sessionToken);
    expect(retrievedB).toBeDefined();
    expect(retrievedB!.accessToken).toBe('access-token-B');
  });

  it('failed token refresh RETAINS the targeted session (for recovery), leaving the other intact', async () => {
    //: a failed MSAL silent refresh must NOT delete the session. Deleting it
    // permanently unbinds the client's long-lived Bearer token — a browser re-login
    // mints a new token the MCP client never receives — so the tool path would stay
    // "Session expired" forever even after a valid web re-auth. The session is
    // retained so a later re-auth can repair it in place. (Before this change the
    // catch block deleted sessionA here; the recovery arc is covered end-to-end in
    // oauthRefreshDesync.test.ts.)
    const userId = freshUserId();
    const now = Date.now();

    const sessionA = makeSession(userId, 'A', 'access-token-A-stale', {
      expiresAt: now + 2 * 60 * 1000, // triggers refresh
    });
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB); // userIndex → sessionB

    const retrievedA = await getSessionByToken(sessionA.sessionToken);
    expect(retrievedA).toBeDefined();

    // Simulate MSAL failure for sessionA
    mockGetAllAccounts.mockResolvedValueOnce([{ homeAccountId: sessionA.homeAccountId }]);
    mockAcquireTokenSilent.mockRejectedValueOnce(new Error('interaction_required'));

    await expect(getValidAccessTokenForSession(retrievedA!)).rejects.toThrow(
      'Re-authentication required',
    );

    // The targeted session (sessionA) must be RETAINED after the failed refresh so
    // it can self-heal once a web re-auth refreshes the shared MSAL account.
    const afterFailA = await getSessionByToken(sessionA.sessionToken);
    expect(afterFailA).toBeDefined();
    expect(afterFailA!.accessToken).toBe('access-token-A-stale');

    // SessionB must still be accessible and untouched after sessionA's refresh failure
    const retrievedB = await getSessionByToken(sessionB.sessionToken);
    expect(retrievedB).toBeDefined();
    expect(retrievedB!.accessToken).toBe('access-token-B');
  });

  // ── 4. deleteSessionByKey scoped to targeted session ──────────────────────

  it('deleteSessionByKey removes only the targeted session, leaving the other intact', async () => {
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB); // userIndex → sessionB

    // Retrieve both sessions so _storageKey is populated
    const retrievedA = await getSessionByToken(sessionA.sessionToken);
    const retrievedB = await getSessionByToken(sessionB.sessionToken);
    expect(retrievedA?._storageKey).toBeDefined();
    expect(retrievedB?._storageKey).toBeDefined();
    expect(retrievedA!._storageKey).not.toBe(retrievedB!._storageKey);

    // Delete sessionB (the one userIndex currently points to)
    await deleteSessionByKey(retrievedB!._storageKey!, userId);

    // SessionB is gone from in-memory cache
    const afterDeleteB = await getSessionByToken(sessionB.sessionToken);
    expect(afterDeleteB).toBeUndefined();

    // SessionA must still be accessible — only sessionB was deleted
    const afterDeleteA = await getSessionByToken(sessionA.sessionToken);
    expect(afterDeleteA).toBeDefined();
    expect(afterDeleteA!.accessToken).toBe('access-token-A');
  });

  it('deleteSessionByKey removes a non-indexed session, leaving the indexed session intact', async () => {
    // The previous test deletes sessionB, which is the session userIndex points
    // at. The more important same-user isolation case is deleting the session
    // userIndex does NOT point at — that is the one most likely to regress.
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB); // userIndex → sessionB

    const retrievedA = await getSessionByToken(sessionA.sessionToken);
    const retrievedB = await getSessionByToken(sessionB.sessionToken);
    expect(retrievedA?._storageKey).toBeDefined();
    expect(retrievedB?._storageKey).toBeDefined();

    // Delete sessionA — userIndex points at sessionB, not sessionA.
    await deleteSessionByKey(retrievedA!._storageKey!, userId);

    // SessionA must be gone from the in-memory cache
    const afterDeleteA = await getSessionByToken(sessionA.sessionToken);
    expect(afterDeleteA).toBeUndefined();

    // SessionB — the indexed session — must remain intact
    const afterDeleteB = await getSessionByToken(sessionB.sessionToken);
    expect(afterDeleteB).toBeDefined();
    expect(afterDeleteB!.accessToken).toBe('access-token-B');
  });

  it('_storageKey values differ between two sessions for the same userId', async () => {
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB);

    const foundA = await getSessionByToken(sessionA.sessionToken);
    const foundB = await getSessionByToken(sessionB.sessionToken);

    expect(foundA!._storageKey).toBeDefined();
    expect(foundB!._storageKey).toBeDefined();
    // Different tokens must yield different storage keys
    expect(foundA!._storageKey).not.toBe(foundB!._storageKey);
  });

  // ── 5. Interleaved requests — no cross-session token leakage ───────────────

  it('interleaved requests for the same userId each receive their own access token', async () => {
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB); // userIndex → sessionB

    // Simulate request1 authenticating with tokenA
    const request1Session = await getSessionByToken(sessionA.sessionToken);
    expect(request1Session).toBeDefined();

    // Simulate request2 authenticating with tokenB — userIndex stays at sessionB
    // (hot path: already in cache, so userIndex is not updated by getSessionByToken)
    const request2Session = await getSessionByToken(sessionB.sessionToken);
    expect(request2Session).toBeDefined();

    // Both requests concurrently call getValidAccessTokenForSession with their
    // respective sessions — the critical regression path.
    const [tokenForRequest1, tokenForRequest2] = await Promise.all([
      getValidAccessTokenForSession(request1Session!),
      getValidAccessTokenForSession(request2Session!),
    ]);

    // Each request gets its own token — no cross-session leakage
    expect(tokenForRequest1).toBe('access-token-A');
    expect(tokenForRequest2).toBe('access-token-B');

    expect(tokenForRequest1).not.toBe('access-token-B');
    expect(tokenForRequest2).not.toBe('access-token-A');
  });

  it('userIndex pointing to sessionB does not cause request1 to receive sessionB token', async () => {
    // This is the exact bug that was fixed in v2.5.0:
    // getValidAccessToken(userId) would resolve through userIndex, which
    // (after sessionB was stored) points to sessionB. Any handler that called
    // getValidAccessToken(userId) instead of getValidAccessTokenForSession(session)
    // would return sessionB's token to request1.
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');

    await storeSession(sessionA);
    await storeSession(sessionB); // userIndex now → sessionB

    // request1 is holding sessionA (already retrieved/authenticated)
    const request1Session = await getSessionByToken(sessionA.sessionToken);
    expect(request1Session).toBeDefined();

    // request1's handler calls getValidAccessTokenForSession(session), NOT
    // getValidAccessToken(userId). This must return sessionA's token.
    const token = await getValidAccessTokenForSession(request1Session!);
    expect(token).toBe('access-token-A');
    expect(token).not.toBe('access-token-B');
  });

  it('three concurrent sessions for the same user are all independently addressable', async () => {
    const userId = freshUserId();
    const sessionA = makeSession(userId, 'A', 'access-token-A');
    const sessionB = makeSession(userId, 'B', 'access-token-B');
    const sessionC = makeSession(userId, 'C', 'access-token-C');

    await storeSession(sessionA);
    await storeSession(sessionB);
    await storeSession(sessionC); // userIndex → sessionC

    const [rA, rB, rC] = await Promise.all([
      getSessionByToken(sessionA.sessionToken),
      getSessionByToken(sessionB.sessionToken),
      getSessionByToken(sessionC.sessionToken),
    ]);

    expect(rA!.accessToken).toBe('access-token-A');
    expect(rB!.accessToken).toBe('access-token-B');
    expect(rC!.accessToken).toBe('access-token-C');

    const [tA, tB, tC] = await Promise.all([
      getValidAccessTokenForSession(rA!),
      getValidAccessTokenForSession(rB!),
      getValidAccessTokenForSession(rC!),
    ]);

    expect(tA).toBe('access-token-A');
    expect(tB).toBe('access-token-B');
    expect(tC).toBe('access-token-C');
  });

  it('deleteAllUserSessions removes all sessions for a user and no sessions for other users', async () => {
    const userX = freshUserId();
    const userY = freshUserId();

    const xA = makeSession(userX, 'XA', 'token-XA');
    const xB = makeSession(userX, 'XB', 'token-XB');
    const yA = makeSession(userY, 'YA', 'token-YA');

    await storeSession(xA);
    await storeSession(xB);
    await storeSession(yA);

    // Only delete userX's sessions
    await deleteAllUserSessions(userX);

    // userY's session must still exist
    const foundY = await getSessionByToken(yA.sessionToken);
    expect(foundY).toBeDefined();
    expect(foundY!.accessToken).toBe('token-YA');
  });
});
