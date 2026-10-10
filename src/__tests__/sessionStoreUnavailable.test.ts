/**
 * A session storage failure is a server error, never an unknown token.
 *
 * getSessionByToken used to catch every error from the storage lookup, log it
 * and return undefined, so a Table Storage 429, 5xx, auth or network failure
 * ended as a 401 and sent clients back to sign in over an outage. It now
 * throws SessionStoreUnavailableError, and every auth entry point lets it
 * through rather than turning it into an auth result.
 *
 * These tests run the real tokenCache and authMiddleware over a mocked
 * tableStorage. The HTTP entry points that map the error to a 503 are covered
 * in their own suites.
 */

import { randomBytes } from 'crypto';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');
process.env.MCP_DATA_ENCRYPTION_KEY = randomBytes(32).toString('hex');

import { jest } from '@jest/globals';
import type { HttpRequest } from '@azure/functions';
import type { UserSession } from '../services/tokenCache.js';

// ── Storage ──────────────────────────────────────────────────────────────────

const mockRows = new Map<string, UserSession>();
const mockStorage: { failure: Error | null; failFor: Set<string>; lookups: string[] } = {
  failure: null,
  failFor: new Set(),
  lookups: [],
};

jest.mock('../services/tableStorage.js', () => {
  const { hashSessionToken } = require('../services/credentialCrypto.js');
  class SessionRowMissingError extends Error {}
  const rowKeyFor = (token: string) => hashSessionToken(token).slice(0, 32);
  return {
    SessionRowMissingError,
    saveSession: async (session: UserSession) => {
      const rowKey = session._storageKey ?? rowKeyFor(session.sessionToken);
      mockRows.set(rowKey, { ...session, sessionToken: '', _storageKey: rowKey });
    },
    loadSessionByToken: async (token: string) => {
      mockStorage.lookups.push(token);
      if (mockStorage.failure) throw mockStorage.failure;
      if (mockStorage.failFor.has(token)) throw Object.assign(new Error('ServerBusy'), { statusCode: 503 });
      const row = mockRows.get(rowKeyFor(token));
      return row ? { ...row, sessionToken: token } : null;
    },
    loadSession: async () => null,
    sessionRowExists: async (storageKey: string) => {
      if (mockStorage.failure) throw mockStorage.failure;
      return mockRows.has(storageKey);
    },
    removeSession: async () => undefined,
    removeSessionByKey: async () => undefined,
    listAllSessions: async () => [],
  };
});

jest.mock('../services/graphClient.js', () => ({
  acquireTokenSilent: jest.fn(),
  createGraphClient: jest.fn(),
}));

// Console checks that don't touch storage always pass, so the test reaches the lookup.
jest.mock('../services/consoleSession.js', () => ({
  CONSOLE_COOKIE: 'mcp_console',
  checkBrowserOrigin: () => ({ ok: true }),
  readCookie: (req: HttpRequest, name: string) =>
    (req.headers.get('cookie') ?? '').match(new RegExp(`(?:^|;\\s*)${name}=([^;]+)`))?.[1] ?? null,
  verifyConsoleToken: (value: string | null) => (value ? { issuedAt: Date.now() } : null),
}));

type TokenCache = typeof import('../services/tokenCache.js');
type AuthMiddleware = typeof import('../services/authMiddleware.js');

function load(): { cache: TokenCache; auth: AuthMiddleware } {
  let cache: TokenCache | undefined;
  let auth: AuthMiddleware | undefined;
  jest.isolateModules(() => {
    cache = require('../services/tokenCache.js');
    auth = require('../services/authMiddleware.js');
  });
  return { cache: cache!, auth: auth! };
}

function storageError(statusCode: number | undefined, message: string): Error {
  return Object.assign(new Error(message), statusCode === undefined ? {} : { statusCode });
}

function newSession(): UserSession {
  const now = Date.now();
  return {
    userId: `user-${randomBytes(4).toString('hex')}`,
    homeAccountId: 'home',
    displayName: 'Test User',
    email: 'user@example.com',
    tenantId: 'tenant-1',
    accessToken: 'access-token',
    expiresAt: now + 3_600_000,
    sessionToken: randomBytes(32).toString('hex'),
    sessionCreatedAt: now,
    sessionAbsoluteCreatedAt: now,
  };
}

function request(headers: Record<string, string>): HttpRequest {
  return { headers: new Map(Object.entries(headers)) } as unknown as HttpRequest;
}

let now = Date.now();
beforeEach(() => {
  mockRows.clear();
  mockStorage.failure = null;
  mockStorage.failFor.clear();
  mockStorage.lookups = [];
  now = Date.now();
  jest.spyOn(Date, 'now').mockImplementation(() => now);
  jest.spyOn(console, 'error').mockImplementation(() => {});
  jest.spyOn(console, 'warn').mockImplementation(() => {});
});
afterEach(() => jest.restoreAllMocks());

const FAILURES: Array<[string, Error]> = [
  ['throttled (429)', storageError(429, 'TooManyRequests')],
  ['server error (503)', storageError(503, 'ServerBusy')],
  ['auth failure (403)', storageError(403, 'AuthorizationFailure: secret-account-detail')],
  ['network failure', storageError(undefined, 'getaddrinfo ENOTFOUND storage.example.com')],
];

// ── getSessionByToken ────────────────────────────────────────────────────────

describe('getSessionByToken', () => {
  it.each(FAILURES)('throws SessionStoreUnavailableError when storage is %s', async (_label, failure) => {
    const { cache } = load();
    mockStorage.failure = failure;

    const err = await cache.getSessionByToken('some-token').catch((e: unknown) => e);
    expect(err).toBeInstanceOf(cache.SessionStoreUnavailableError);
    // Fixed message: handlers that echo err.message must not leak storage detail.
    expect((err as Error).message).toBe('Session store unavailable');
    expect((err as Error).cause).toBe(failure);
  });

  it('still returns undefined when storage answers that no session has the token', async () => {
    const { cache } = load();
    await expect(cache.getSessionByToken('unknown-token')).resolves.toBeUndefined();
  });

  it('caches nothing from a failed lookup, so the session works once storage is back', async () => {
    const { cache } = load();
    const session = newSession();
    await load().cache.storeSession(session); // written by another replica

    mockStorage.failure = storageError(503, 'ServerBusy');
    await expect(cache.getSessionByToken(session.sessionToken)).rejects.toBeInstanceOf(
      cache.SessionStoreUnavailableError,
    );

    mockStorage.failure = null;
    await expect(cache.getSessionByToken(session.sessionToken)).resolves.toMatchObject({
      userId: session.userId,
    });
  });

  it('throws, not undefined, for a cached session past the revalidation stale limit', async () => {
    const { cache } = load();
    const session = newSession();
    await cache.storeSession(session);
    mockStorage.failure = storageError(503, 'ServerBusy');

    now += cache.SESSION_REVALIDATE_MS + cache.SESSION_REVALIDATE_MAX_STALE_MS;
    await expect(cache.getSessionByToken(session.sessionToken)).rejects.toBeInstanceOf(
      cache.SessionStoreUnavailableError,
    );
  });

  it('getSession keeps returning undefined past the stale limit', async () => {
    const { cache } = load();
    const session = newSession();
    await cache.storeSession(session);
    mockStorage.failure = storageError(503, 'ServerBusy');

    now += cache.SESSION_REVALIDATE_MS + cache.SESSION_REVALIDATE_MAX_STALE_MS;
    await expect(cache.getSession(session.userId)).resolves.toBeUndefined();
  });
});

// ── Auth entry points ────────────────────────────────────────────────────────

describe('authenticateRequest', () => {
  it.each(FAILURES)('rejects with SessionStoreUnavailableError when storage is %s', async (_label, failure) => {
    const { cache, auth } = load();
    mockStorage.failure = failure;
    await expect(
      auth.authenticateRequest(request({ authorization: 'Bearer some-token' })),
    ).rejects.toBeInstanceOf(cache.SessionStoreUnavailableError);
  });

  it('does not fall through to a later candidate after an unanswered lookup', async () => {
    const { cache, auth } = load();
    const session = newSession();
    await load().cache.storeSession(session);

    // The bearer lookup fails; the cookie would authenticate if it were tried.
    // The bearer may have been a session that should be refused, so the
    // request must not be let in on the cookie instead.
    mockStorage.failFor.add('bearer-token');

    await expect(
      auth.authenticateRequest(
        request({ authorization: 'Bearer bearer-token', cookie: `mcp_session=${session.sessionToken}` }),
      ),
    ).rejects.toBeInstanceOf(cache.SessionStoreUnavailableError);
    expect(mockStorage.lookups).toEqual(['bearer-token']);
  });

  it('still returns null for a token storage does not know', async () => {
    const { auth } = load();
    await expect(auth.authenticateRequest(request({ authorization: 'Bearer unknown' }))).resolves.toBeNull();
  });
});

describe('authenticateRequestAllowExpired', () => {
  it.each(FAILURES)('rejects with SessionStoreUnavailableError when storage is %s', async (_label, failure) => {
    const { cache, auth } = load();
    mockStorage.failure = failure;
    await expect(
      auth.authenticateRequestAllowExpired(request({ authorization: 'Bearer some-token' })),
    ).rejects.toBeInstanceOf(cache.SessionStoreUnavailableError);
  });

  it('still returns null for a token storage does not know', async () => {
    const { auth } = load();
    await expect(
      auth.authenticateRequestAllowExpired(request({ authorization: 'Bearer unknown' })),
    ).resolves.toBeNull();
  });
});

describe('authenticateConsoleRequest', () => {
  it.each(FAILURES)('rejects with SessionStoreUnavailableError when storage is %s', async (_label, failure) => {
    const { cache, auth } = load();
    mockStorage.failure = failure;
    await expect(
      auth.authenticateConsoleRequest(request({ cookie: 'mcp_session=some-token; mcp_console=console-token' })),
    ).rejects.toBeInstanceOf(cache.SessionStoreUnavailableError);
  });
});
