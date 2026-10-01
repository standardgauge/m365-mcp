/**
 * Regression test for — OAuth refresh desync.
 *
 * Reported symptom (your-mcp-host.example.com, pat@example.com): every Mail tool call
 * returned "Session expired. Please re-authenticate at: .../api/auth/login".
 * The user DID re-authenticate on the web (valid web session, signed in on the
 * admin page), yet the MCP tool path stayed expired for a week.
 *
 * Root cause: the tool-path token resolver (getValidAccessTokenForSession) used
 * to DELETE the session whenever MSAL silent refresh failed. That permanently
 * unbinds the MCP client's long-lived Bearer token — a browser re-login mints a
 * BRAND-NEW session token the client never receives — so no successful web
 * re-auth can ever re-bind the client's existing token. The web session store
 * and the token the tool path reads were out of sync, with no recovery path.
 *
 * The fix: on MSAL failure the tool path RETAINS the session and logs the cause.
 * A later web re-auth refreshes the shared MSAL account (same homeAccountId),
 * after which the next tool call silently repairs the SAME session in place and
 * returns a fresh token — no client reconfiguration required.
 *
 * These tests model the desync directly:
 *   1. valid web session + stale tool-path token: MSAL fails → the tool path
 *      throws re-auth, but the session is RETAINED (not deleted) and the failure
 *      is logged (no silent failure).
 *   2. recovery: after a web re-auth has refreshed the shared MSAL cache, the
 *      SAME stale session (same Bearer token) resolves to a fresh access token
 *      in place — the desync self-heals.
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

// ── Module under test ─────────────────────────────────────────────────────────

import {
  storeSession,
  getSessionByToken,
  getValidAccessTokenForSession,
} from '../services/tokenCache.js';
import type { UserSession } from '../services/tokenCache.js';

// ── Helpers ───────────────────────────────────────────────────────────────────

let userCounter = 0;
function freshUserId(): string {
  return `desync-test-user-${++userCounter}-${randomBytes(4).toString('hex')}`;
}

/**
 * A session whose Graph access token is already stale (expiry in the past), so
 * getValidAccessTokenForSession is forced down the MSAL silent-refresh path.
 * This models the token an MCP client holds after its access token expired.
 */
function makeStaleSession(userId: string): UserSession {
  return {
    userId,
    homeAccountId: `home-account-${userId}`,
    displayName: 'Pat Example',
    email: 'pat@example.com',
    tenantId: 'test-tenant-id',
    accessToken: 'stale-access-token',
    expiresAt: Date.now() - 60 * 1000, // already expired → must refresh
    sessionToken: `session-token-${randomBytes(8).toString('hex')}`,
    sessionCreatedAt: Date.now(),
    sessionAbsoluteCreatedAt: Date.now(),
  };
}

afterEach(() => {
  jest.clearAllMocks();
});

describe(' — OAuth refresh desync (valid web session + stale tool-path token)', () => {
  it('retains the session and logs the cause when MSAL silent refresh fails (no self-destruct, no silent failure)', async () => {
    const userId = freshUserId();
    const session = makeStaleSession(userId);
    await storeSession(session);

    const held = await getSessionByToken(session.sessionToken);
    expect(held).toBeDefined();

    // MSAL refresh token is momentarily unusable (revoked / not yet re-consented)
    mockGetAllAccounts.mockResolvedValueOnce([{ homeAccountId: session.homeAccountId }]);
    mockAcquireTokenSilent.mockRejectedValueOnce(new Error('invalid_grant: token revoked'));

    const errorSpy = jest.spyOn(console, 'error').mockImplementation(() => {});

    await expect(getValidAccessTokenForSession(held!)).rejects.toThrow(
      'Re-authentication required',
    );

    // The failure must be logged with its underlying cause — a refresh that fails
    // silently and destroys recoverable state is a write path with no consumer.
    const loggedRefreshFailure = errorSpy.mock.calls.some((args) =>
      args.some(
        (a) =>
          typeof a === 'string' &&
          a.includes('Silent token refresh FAILED') &&
          a.includes('invalid_grant'),
      ),
    );
    expect(loggedRefreshFailure).toBe(true);
    errorSpy.mockRestore();

    // The session MUST still be resolvable by the client's original Bearer token.
    // If it were deleted, the client's token would map to nothing and no web
    // re-auth could ever repair it (the permanent-desync failure mode).
    const afterFailure = await getSessionByToken(session.sessionToken);
    expect(afterFailure).toBeDefined();
    expect(afterFailure!.sessionToken).toBe(session.sessionToken);
  });

  it('self-heals: after a web re-auth refreshes the shared MSAL account, the SAME stale token resolves to a fresh access token', async () => {
    const userId = freshUserId();
    const session = makeStaleSession(userId);
    await storeSession(session);

    const held = await getSessionByToken(session.sessionToken);
    expect(held).toBeDefined();

    // ── Before web re-auth: MSAL cannot refresh → tool path errors, session kept ──
    mockGetAllAccounts.mockResolvedValueOnce([{ homeAccountId: session.homeAccountId }]);
    mockAcquireTokenSilent.mockRejectedValueOnce(new Error('interaction_required'));

    await expect(getValidAccessTokenForSession(held!)).rejects.toThrow(
      'Re-authentication required',
    );

    // ── After web re-auth: the shared MSAL account (same homeAccountId) now has a
    //    fresh refresh token, so silent acquisition succeeds. The client keeps
    //    sending its ORIGINAL Bearer token; the same session must recover. ──
    const freshExpiry = new Date(Date.now() + 60 * 60 * 1000);
    mockGetAllAccounts.mockResolvedValueOnce([{ homeAccountId: session.homeAccountId }]);
    mockAcquireTokenSilent.mockResolvedValueOnce({
      accessToken: 'recovered-access-token',
      expiresOn: freshExpiry,
    });

    const recovered = await getSessionByToken(session.sessionToken);
    expect(recovered).toBeDefined();

    const token = await getValidAccessTokenForSession(recovered!);
    expect(token).toBe('recovered-access-token');

    // The repaired token/expiry are persisted on the same session, so subsequent
    // calls return the fresh token from cache without another MSAL round-trip.
    const afterRecovery = await getSessionByToken(session.sessionToken);
    expect(afterRecovery).toBeDefined();
    expect(afterRecovery!.accessToken).toBe('recovered-access-token');
    expect(afterRecovery!.expiresAt).toBe(freshExpiry.getTime());
  });
});
