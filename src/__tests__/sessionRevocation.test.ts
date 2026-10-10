/**
 * Session deletion reaches sessions cached in every replica's memory.
 *
 * Two copies of tokenCache are loaded in isolated module registries, standing
 * in for two replicas, over one shared in-memory sessions table. Deleting a
 * session's rows on replica A must stop the session authenticating on replica
 * B within SESSION_REVALIDATE_MS, and B must not write the deleted row back
 * when it refreshes the session in that window.
 */

import { randomBytes } from 'crypto';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');
process.env.MCP_DATA_ENCRYPTION_KEY = randomBytes(32).toString('hex');

import { jest } from '@jest/globals';
import type { UserSession } from '../services/tokenCache.js';

type TokenCache = typeof import('../services/tokenCache.js');

// ── Shared sessions table ────────────────────────────────────────────────────

const mockRows = new Map<string, UserSession>();
const mockStorage = { down: false, rowReads: 0 };

jest.mock('../services/tableStorage.js', () => {
  const { hashSessionToken } = require('../services/credentialCrypto.js');
  class SessionRowMissingError extends Error {}
  const rowKeyFor = (token: string) => hashSessionToken(token).slice(0, 32);
  return {
    SessionRowMissingError,
    saveSession: async (session: UserSession, mode: 'create' | 'update' = 'create') => {
      if (mockStorage.down) throw new Error('storage unreachable');
      if (mode === 'update' && session._storageKey) {
        if (!mockRows.has(session._storageKey)) throw new SessionRowMissingError(session._storageKey);
        mockRows.set(session._storageKey, { ...session, sessionToken: '' });
        return;
      }
      const rowKey = rowKeyFor(session.sessionToken);
      mockRows.set(rowKey, { ...session, sessionToken: '', _storageKey: rowKey });
    },
    loadSessionByToken: async (token: string) => {
      if (mockStorage.down) throw new Error('storage unreachable');
      const row = mockRows.get(rowKeyFor(token));
      return row ? { ...row, sessionToken: token } : null;
    },
    loadSession: async () => null,
    sessionRowExists: async (storageKey: string) => {
      mockStorage.rowReads++;
      if (mockStorage.down) throw new Error('storage unreachable');
      return mockRows.has(storageKey);
    },
    removeSession: async (userId: string) => {
      for (const [key, row] of mockRows) if (row.userId === userId) mockRows.delete(key);
    },
    removeSessionByKey: async (key: string) => { mockRows.delete(key); },
    listAllSessions: async () => [...mockRows.values()],
  };
});

jest.mock('../services/graphClient.js', () => ({
  acquireTokenSilent: jest.fn(),
}));

// ── Helpers ──────────────────────────────────────────────────────────────────

function loadReplica(): TokenCache {
  let mod: TokenCache | undefined;
  jest.isolateModules(() => {
    mod = require('../services/tokenCache.js');
  });
  return mod!;
}

let counter = 0;
function newSession(userId = `user-${++counter}`): UserSession {
  const now = Date.now();
  return {
    userId,
    homeAccountId: `home-${userId}`,
    displayName: 'Test User',
    email: `${userId}@example.com`,
    tenantId: 'tenant-1',
    accessToken: 'access-token',
    expiresAt: now + 3_600_000,
    sessionToken: `token-${++counter}-${randomBytes(8).toString('hex')}`,
    sessionCreatedAt: now,
    sessionAbsoluteCreatedAt: now,
  };
}

let now = Date.now();
beforeEach(() => {
  mockRows.clear();
  mockStorage.down = false;
  mockStorage.rowReads = 0;
  now = Date.now();
  jest.spyOn(Date, 'now').mockImplementation(() => now);
});
afterEach(() => jest.restoreAllMocks());

// ── Tests ────────────────────────────────────────────────────────────────────

describe('session deletion across replicas', () => {
  it('a session deleted on one replica stops authenticating on another within SESSION_REVALIDATE_MS', async () => {
    const a = loadReplica();
    const b = loadReplica();
    const session = newSession();
    const token = session.sessionToken;
    await a.storeSession(session);

    // Replica B serves the token and caches it.
    expect(await b.getSessionByToken(token)).toBeDefined();

    // Logout on replica A deletes every row for the user.
    await a.deleteAllUserSessions(session.userId);
    expect(await a.getSessionByToken(token)).toBeUndefined();

    // Replica B learns of it at the next revalidation, and no later.
    now += a.SESSION_REVALIDATE_MS;
    expect(await b.getSessionByToken(token)).toBeUndefined();
    expect(mockRows.size).toBe(0);
  });

  it('a refresh on another replica does not write the deleted row back', async () => {
    const a = loadReplica();
    const b = loadReplica();
    const session = newSession();
    await a.storeSession(session);
    const cached = (await b.getSessionByToken(session.sessionToken))!;

    await a.deleteAllUserSessions(session.userId);

    // Inside the revalidation window, B refreshes its cached copy (TTL
    // refresh, sliding-window touch, access-token refresh all do this).
    await expect(
      b.storeSession({ ...cached, sessionCreatedAt: now }),
    ).rejects.toBeInstanceOf(b.SessionRowMissingError);

    expect(mockRows.size).toBe(0);
    // Evicted at once, without waiting for revalidation.
    expect(await b.getSessionByToken(session.sessionToken)).toBeUndefined();
  });

  it('deleting the row by hand behaves the same as logout', async () => {
    const b = loadReplica();
    const session = newSession();
    await b.storeSession(session);
    expect(await b.getSessionByToken(session.sessionToken)).toBeDefined();

    mockRows.clear(); // purge-credentials.sh, or a row deleted in the portal

    now += b.SESSION_REVALIDATE_MS;
    expect(await b.getSessionByToken(session.sessionToken)).toBeUndefined();
  });

  it('logout evicts every cached session for the user on its own replica', async () => {
    const a = loadReplica();
    const laptop = newSession('user-two-devices');
    const phone = newSession('user-two-devices');
    await a.storeSession(laptop);
    await a.storeSession(phone);

    await a.deleteAllUserSessions('user-two-devices');

    // No time passes: both are gone immediately on the replica that logged out.
    expect(await a.getSessionByToken(laptop.sessionToken)).toBeUndefined();
    expect(await a.getSessionByToken(phone.sessionToken)).toBeUndefined();
  });
});

describe('cached-session revalidation', () => {
  it('checks storage at most once per SESSION_REVALIDATE_MS', async () => {
    const b = loadReplica();
    const session = newSession();
    await b.storeSession(session);

    for (let i = 0; i < 5; i++) await b.getSessionByToken(session.sessionToken);
    expect(mockStorage.rowReads).toBe(0);

    now += b.SESSION_REVALIDATE_MS;
    for (let i = 0; i < 5; i++) await b.getSessionByToken(session.sessionToken);
    expect(mockStorage.rowReads).toBe(1);
  });

  it('serves a cached session through a storage outage up to the stale limit, then refuses it', async () => {
    const b = loadReplica();
    const session = newSession();
    await b.storeSession(session);
    mockStorage.down = true;

    now += b.SESSION_REVALIDATE_MS;
    expect(await b.getSessionByToken(session.sessionToken)).toBeDefined();

    // Refused as a storage failure (a 503 upstream), not as an unknown token.
    now += b.SESSION_REVALIDATE_MAX_STALE_MS;
    await expect(b.getSessionByToken(session.sessionToken)).rejects.toBeInstanceOf(
      b.SessionStoreUnavailableError,
    );

    // Storage back and the row still there: the session works again.
    mockStorage.down = false;
    expect(await b.getSessionByToken(session.sessionToken)).toBeDefined();
  });

  it('a freshly minted session still creates its row', async () => {
    const a = loadReplica();
    const session = newSession();
    await a.storeSession(session);
    expect(mockRows.size).toBe(1);
    expect(session._storageKey).toBeDefined();
  });
});
