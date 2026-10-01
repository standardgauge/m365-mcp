/**
 * Unit tests for graphClient.ts and tokenCache.ts.
 *
 * MSAL is mocked — no real network calls are made.
 * Tests cover:
 *   - Env var validation
 *   - Graph client factory
 *   - Token cache: returns cached token when valid
 *   - Token cache: triggers silent refresh when token is near expiry
 *   - Token cache: throws "Re-authentication required" when refresh fails
 */

import { randomBytes } from 'crypto';

//: tableStorage.saveSession now hashes the session token using a key
// from MCP_SESSION_HMAC_KEY and encrypts the access token using a key from
// MCP_DATA_ENCRYPTION_KEY. Both must be set before tokenCache.storeSession is
// called by any test, even though Table Storage itself is mocked away in
// these tests (the saveSession path silently swallows storage errors).
process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');
process.env.MCP_DATA_ENCRYPTION_KEY = randomBytes(32).toString('hex');

import { jest } from '@jest/globals';

// ── MSAL mock ──────────────────────────────────────────────────────────────────

const mockGetAllAccounts = jest.fn<() => Promise<unknown[]>>();
const mockAcquireTokenSilent = jest.fn<() => Promise<unknown>>();
const mockGetTokenCache = jest.fn(() => ({ getAllAccounts: mockGetAllAccounts }));
const mockAcquireTokenByCode = jest.fn<() => Promise<unknown>>();
const mockGetAuthCodeUrl = jest.fn<() => Promise<string>>();

jest.mock('@azure/msal-node', () => ({
  ConfidentialClientApplication: jest.fn().mockImplementation(() => ({
    getTokenCache: mockGetTokenCache,
    acquireTokenSilent: mockAcquireTokenSilent,
    acquireTokenByCode: mockAcquireTokenByCode,
    getAuthCodeUrl: mockGetAuthCodeUrl,
  })),
  LogLevel: { Error: 0, Warning: 1, Info: 2, Verbose: 3, Trace: 4 },
}));

jest.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: jest.fn().mockReturnValue({ api: jest.fn() }),
    initWithMiddleware: jest.fn().mockReturnValue({ api: jest.fn() }),
  },
}));

// isomorphic-fetch is a side-effect import — stub it out
jest.mock('isomorphic-fetch', () => ({}));

// ── Set required env vars ──────────────────────────────────────────────────────

const MOCK_ENV = {
  AZURE_CLIENT_ID: 'test-client-id',
  AZURE_CLIENT_SECRET: 'test-client-secret',
  AZURE_TENANT_ID: 'test-tenant-id',
  OAUTH_REDIRECT_URI: 'http://localhost:7071/api/auth/callback',
};

beforeAll(() => {
  Object.assign(process.env, MOCK_ENV);
});

afterAll(() => {
  for (const key of Object.keys(MOCK_ENV)) {
    delete process.env[key];
  }
});

afterEach(() => {
  jest.clearAllMocks();
});

// ── Imports after mocks ────────────────────────────────────────────────────────

import { createGraphClient, getMsalApp } from '../services/graphClient.js';
import { storeSession, getSession, getValidAccessToken, deleteSession } from '../services/tokenCache.js';

// ── graphClient tests ──────────────────────────────────────────────────────────

describe('getMsalApp', () => {
  it('returns a ConfidentialClientApplication', () => {
    const app = getMsalApp();
    expect(app).toBeDefined();
    expect(app.getTokenCache).toBeDefined();
  });

  it('returns the same instance on repeated calls (singleton)', () => {
    const a = getMsalApp();
    const b = getMsalApp();
    expect(a).toBe(b);
  });
});

describe('createGraphClient', () => {
  it('creates a client that passes the token to the auth provider', () => {
    const client = createGraphClient('my-access-token');
    expect(client).toBeDefined();
  });
});

// ── tokenCache tests ───────────────────────────────────────────────────────────

const MOCK_SESSION = {
  userId: 'user-001',
  homeAccountId: 'home-001',
  displayName: 'Test User',
  email: 'test@example.com',
  tenantId: 'test-tenant-id',
  accessToken: 'valid-token',
  expiresAt: Date.now() + 60 * 60 * 1000, // 1 hour from now
  sessionToken: 'mock-session-token',
  sessionCreatedAt: Date.now(),
};

describe('storeSession / getSession / deleteSession', () => {
  afterEach(async () => deleteSession(MOCK_SESSION.userId));

  it('stores and retrieves a session', async () => {
    await storeSession(MOCK_SESSION);
    const retrieved = await getSession(MOCK_SESSION.userId);
    expect(retrieved).toBeDefined();
    expect(retrieved!.email).toBe(MOCK_SESSION.email);
  });

  it('returns undefined for an unknown userId', async () => {
    expect(await getSession('nobody')).toBeUndefined();
  });

  it('deletes a session', async () => {
    await storeSession(MOCK_SESSION);
    await deleteSession(MOCK_SESSION.userId);
    expect(await getSession(MOCK_SESSION.userId)).toBeUndefined();
  });
});

describe('getValidAccessToken', () => {
  afterEach(async () => deleteSession(MOCK_SESSION.userId));

  it('throws "Re-authentication required" when no session exists', async () => {
    await expect(getValidAccessToken('unknown-user')).rejects.toThrow(
      'Re-authentication required'
    );
  });

  it('returns the cached token when it has more than 5 minutes remaining', async () => {
    await storeSession({ ...MOCK_SESSION, expiresAt: Date.now() + 10 * 60 * 1000 });
    const token = await getValidAccessToken(MOCK_SESSION.userId);
    expect(token).toBe(MOCK_SESSION.accessToken);
    // Silent refresh should NOT have been called
    expect(mockAcquireTokenSilent).not.toHaveBeenCalled();
  });

  it('calls acquireTokenSilent when the token is near expiry', async () => {
    // Store a session that expires in 2 minutes (below the 5-minute buffer)
    await storeSession({ ...MOCK_SESSION, expiresAt: Date.now() + 2 * 60 * 1000 });

    const mockAccount = { homeAccountId: MOCK_SESSION.homeAccountId };
    mockGetAllAccounts.mockResolvedValueOnce([mockAccount]);
    mockAcquireTokenSilent.mockResolvedValueOnce({
      accessToken: 'refreshed-token',
      expiresOn: new Date(Date.now() + 60 * 60 * 1000),
    });

    const token = await getValidAccessToken(MOCK_SESSION.userId);
    expect(mockAcquireTokenSilent).toHaveBeenCalledTimes(1);
    expect(token).toBe('refreshed-token');
  });

  it('throws "Re-authentication required" but RETAINS the session when silent refresh fails', async () => {
    await storeSession({ ...MOCK_SESSION, expiresAt: Date.now() + 2 * 60 * 1000 });

    const mockAccount = { homeAccountId: MOCK_SESSION.homeAccountId };
    mockGetAllAccounts.mockResolvedValueOnce([mockAccount]);
    mockAcquireTokenSilent.mockRejectedValueOnce(new Error('interaction_required'));

    await expect(getValidAccessToken(MOCK_SESSION.userId)).rejects.toThrow(
      'Re-authentication required'
    );

    //: the session must be RETAINED, not cleared. Deleting it permanently
    // unbinds the MCP client's long-lived Bearer token so no web re-auth can
    // re-bind it — the desync would never self-heal. The session survives so a
    // later re-auth can repair it in place (recovery arc: oauthRefreshDesync.test.ts).
    expect(await getSession(MOCK_SESSION.userId)).toBeDefined();
  });

  it('throws "Re-authentication required" but RETAINS the session when MSAL cache has no matching account', async () => {
    await storeSession({ ...MOCK_SESSION, expiresAt: Date.now() + 1 * 60 * 1000 });

    // No accounts in MSAL cache
    mockGetAllAccounts.mockResolvedValueOnce([]);

    await expect(getValidAccessToken(MOCK_SESSION.userId)).rejects.toThrow(
      'Re-authentication required'
    );

    //: retained (see above) — a subsequent web re-auth repopulates the MSAL
    // account and the next call repairs this session rather than 401-ing forever.
    expect(await getSession(MOCK_SESSION.userId)).toBeDefined();
  });
});
