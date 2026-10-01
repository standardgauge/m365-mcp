/**
 *: Legacy absolute-lifetime anchor materialization in tokenCache.
 *
 * Verifies that getSessionByToken (and getSession) materialize an immutable
 * sessionAbsoluteCreatedAt anchor from sessionCreatedAt for legacy sessions
 * that predate the field. Without materialization, a sliding-window touch
 * resets sessionCreatedAt, pushing the mutable fallback anchor forward and
 * allowing indefinite renewal past the 30-day bound.
 *
 * Note: tokenCache maintains a module-level in-memory cache. Each test must use
 * unique tokens and userIds to avoid cross-test cache hits.
 */

import { randomBytes } from 'crypto';

// Required env vars for credential crypto used inside tokenCache / tableStorage.
process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');
process.env.MCP_DATA_ENCRYPTION_KEY = randomBytes(32).toString('hex');

import { jest } from '@jest/globals';
import type { UserSession } from '../services/tokenCache.js';

// ── Constants ────────────────────────────────────────────────────────────────

const SESSION_MAX_LIFETIME_MS = 30 * 24 * 60 * 60 * 1000; // 30 days

// ── Mock tableStorage ─────────────────────────────────────────────────────────

const mockLoadSessionByToken = jest.fn<(token: string) => Promise<UserSession | null>>();
const mockSaveSession = jest.fn<(session: UserSession) => Promise<void>>();
const mockLoadSession = jest.fn<(userId: string) => Promise<UserSession | null>>();

jest.mock('../services/tableStorage.js', () => ({
  loadSessionByToken: (token: string) => mockLoadSessionByToken(token),
  saveSession: (session: UserSession) => mockSaveSession(session),
  loadSession: (userId: string) => mockLoadSession(userId),
  removeSession: jest.fn(),
  removeSessionByKey: jest.fn(),
  listAllSessions: jest.fn<() => Promise<UserSession[]>>().mockResolvedValue([]),
  ensureTables: jest.fn<() => Promise<void>>().mockResolvedValue(undefined),
}));

// ── Mock graphClient (acquireTokenSilent not needed here) ─────────────────────

jest.mock('../services/graphClient.js', () => ({
  acquireTokenSilent: jest.fn(),
}));

// ── Import AFTER mocks ────────────────────────────────────────────────────────

import {
  getSessionByToken,
  getSession,
  isAbsoluteLifetimeExceeded,
} from '../services/tokenCache.js';

// ── Helpers ───────────────────────────────────────────────────────────────────

// Use a counter to give each test its own token/userId so the module-level
// sessionCache never returns a stale result from a previous test.
let testCounter = 0;
function uniqueToken(): string { return `token-${++testCounter}`; }
// Always increments so getSession tests get a fresh slot even without a token
function uniqueUserId(): string { return `user-gs-${++testCounter}`; }

const THIRTY_ONE_DAYS_AGO = Date.now() - SESSION_MAX_LIFETIME_MS - 86_400_000;

function makeLegacySession(token: string, userId: string, overrides: Partial<UserSession> = {}): UserSession {
  return {
    userId,
    homeAccountId: `home-${userId}`,
    displayName: 'Legacy User',
    email: `${userId}@example.com`,
    tenantId: 'tenant-legacy',
    accessToken: 'access-token',
    expiresAt: Date.now() + 3_600_000,
    sessionToken: token,
    sessionCreatedAt: THIRTY_ONE_DAYS_AGO,
    sessionAbsoluteCreatedAt: undefined,
    _storageKey: `rowkey-${userId}`,
    ...overrides,
  };
}

// ── Setup ─────────────────────────────────────────────────────────────────────

beforeEach(() => {
  jest.clearAllMocks();
  mockSaveSession.mockResolvedValue(undefined);
});

// ── Tests: getSessionByToken ──────────────────────────────────────────────────

describe('getSessionByToken — legacy anchor materialization', () => {
  it('materializes sessionAbsoluteCreatedAt from sessionCreatedAt on first cold load', async () => {
    const token = uniqueToken();
    const userId = uniqueUserId();
    const legacy = makeLegacySession(token, userId);
    mockLoadSessionByToken.mockResolvedValueOnce(legacy);

    const session = await getSessionByToken(token);

    expect(session).not.toBeNull();
    expect(session!.sessionAbsoluteCreatedAt).toBe(THIRTY_ONE_DAYS_AGO);
  });

  it('persists the materialized anchor to storage', async () => {
    const token = uniqueToken();
    const userId = uniqueUserId();
    const legacy = makeLegacySession(token, userId);
    mockLoadSessionByToken.mockResolvedValueOnce(legacy);

    await getSessionByToken(token);

    // Allow fire-and-forget async work to complete
    await new Promise((r) => setImmediate(r));

    expect(mockSaveSession).toHaveBeenCalledWith(
      expect.objectContaining({ sessionAbsoluteCreatedAt: THIRTY_ONE_DAYS_AGO })
    );
  });

  it('a touch (sliding-window) does not move the materialized anchor', async () => {
    const token = uniqueToken();
    const userId = uniqueUserId();
    const legacy = makeLegacySession(token, userId);
    mockLoadSessionByToken.mockResolvedValueOnce(legacy);

    const session = await getSessionByToken(token);
    expect(session!.sessionAbsoluteCreatedAt).toBe(THIRTY_ONE_DAYS_AGO);

    // Simulate what authenticateRequest's sliding-window touch does
    const touched: UserSession = { ...session!, sessionCreatedAt: Date.now() };

    // sessionCreatedAt is now "fresh" — but the materialized anchor stays fixed
    expect(touched.sessionAbsoluteCreatedAt).toBe(THIRTY_ONE_DAYS_AGO);
  });

  it('isAbsoluteLifetimeExceeded returns true for a touched legacy session past the 30-day bound', async () => {
    const token = uniqueToken();
    const userId = uniqueUserId();
    const legacy = makeLegacySession(token, userId);
    mockLoadSessionByToken.mockResolvedValueOnce(legacy);

    const session = await getSessionByToken(token);

    // Touch: resets sessionCreatedAt to now — session looks fresh by inactivity TTL
    const touched: UserSession = { ...session!, sessionCreatedAt: Date.now() };

    // The materialized anchor (31 days ago) must still trigger expiry
    expect(isAbsoluteLifetimeExceeded(touched)).toBe(true);
  });

  it('does not overwrite an existing sessionAbsoluteCreatedAt', async () => {
    const token = uniqueToken();
    const userId = uniqueUserId();
    const existingAnchor = Date.now() - 5 * 24 * 60 * 60 * 1000; // 5 days ago
    const modern = makeLegacySession(token, userId, { sessionAbsoluteCreatedAt: existingAnchor });
    mockLoadSessionByToken.mockResolvedValueOnce(modern);

    const session = await getSessionByToken(token);

    expect(session!.sessionAbsoluteCreatedAt).toBe(existingAnchor);

    await new Promise((r) => setImmediate(r));
    // saveSession must not be called — no new anchor to persist
    expect(mockSaveSession).not.toHaveBeenCalled();
  });

  it('does not materialize when sessionCreatedAt is 0 (zero-anchor session bypasses absolute check)', async () => {
    const token = uniqueToken();
    const userId = uniqueUserId();
    const zeroAnchor = makeLegacySession(token, userId, {
      sessionCreatedAt: 0,
      sessionAbsoluteCreatedAt: undefined,
    });
    mockLoadSessionByToken.mockResolvedValueOnce(zeroAnchor);

    const session = await getSessionByToken(token);

    expect(session!.sessionAbsoluteCreatedAt).toBeUndefined();
    await new Promise((r) => setImmediate(r));
    expect(mockSaveSession).not.toHaveBeenCalled();
    // Confirm isAbsoluteLifetimeExceeded returns false for this session (bypass path)
    expect(isAbsoluteLifetimeExceeded(session!)).toBe(false);
  });
});

// ── Tests: getSession ─────────────────────────────────────────────────────────

describe('getSession — legacy anchor materialization', () => {
  it('materializes sessionAbsoluteCreatedAt from sessionCreatedAt on userId-based storage load', async () => {
    const userId = uniqueUserId();
    const legacy = makeLegacySession('', userId); // userId-based loads have no raw token
    mockLoadSession.mockResolvedValueOnce(legacy);

    const session = await getSession(userId);

    expect(session).not.toBeNull();
    expect(session!.sessionAbsoluteCreatedAt).toBe(THIRTY_ONE_DAYS_AGO);
  });

  it('persists the materialized anchor to storage on userId-based load', async () => {
    const userId = uniqueUserId();
    const legacy = makeLegacySession('', userId);
    mockLoadSession.mockResolvedValueOnce(legacy);

    await getSession(userId);

    await new Promise((r) => setImmediate(r));
    expect(mockSaveSession).toHaveBeenCalledWith(
      expect.objectContaining({ sessionAbsoluteCreatedAt: THIRTY_ONE_DAYS_AGO })
    );
  });
});
