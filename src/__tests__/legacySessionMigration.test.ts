/**
 * An unknown Bearer token costs one point read, not a partition scan. Legacy
 * session rows (RowKey = userId) are re-keyed to RowKey = tokenHash[:32] once,
 * before the first lookup, so the token lookup never needs to scan for them.
 * Exercises tableStorage.ts against an in-memory table that counts scans.
 */

import { jest } from '@jest/globals';
import { randomBytes } from 'crypto';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');
process.env.MCP_DATA_ENCRYPTION_KEY = randomBytes(32).toString('hex');
process.env.AZURE_STORAGE_CONNECTION_STRING =
  'DefaultEndpointsProtocol=https;AccountName=test;AccountKey=dGVzdA==;EndpointSuffix=core.windows.net';

// ── In-memory Azure Table Storage ────────────────────────────────────────────

type Entity = Record<string, unknown> & { partitionKey: string; rowKey: string; etag?: string };
const tables = new Map<string, Map<string, Entity>>();
let etagCounter = 0;
const calls = { list: 0, get: 0 };
/** When set, the next getEntity throws this instead of reading. */
let getFault: Error | null = null;

function tableFor(name: string): Map<string, Entity> {
  if (!tables.has(name)) tables.set(name, new Map());
  return tables.get(name)!;
}

function makeTableClient(name: string) {
  const rows = () => tableFor(name);
  const key = (pk: string, rk: string) => `${pk}|${rk}`;
  return {
    async getEntity(pk: string, rk: string) {
      calls.get++;
      if (getFault) {
        const fault = getFault;
        getFault = null;
        throw fault;
      }
      const row = rows().get(key(pk, rk));
      if (!row) throw Object.assign(new Error('not found'), { statusCode: 404 });
      return { ...row };
    },
    async createEntity(entity: Entity) {
      if (rows().has(key(entity.partitionKey, entity.rowKey))) {
        throw Object.assign(new Error('conflict'), { statusCode: 409 });
      }
      rows().set(key(entity.partitionKey, entity.rowKey), { ...entity, etag: `e${++etagCounter}` });
    },
    async upsertEntity(entity: Entity, mode: string) {
      const existing = rows().get(key(entity.partitionKey, entity.rowKey));
      const merged = mode === 'Merge' && existing ? { ...existing, ...entity } : { ...entity };
      rows().set(key(entity.partitionKey, entity.rowKey), { ...merged, etag: `e${++etagCounter}` });
    },
    async updateEntity(entity: Entity, mode: string) {
      const existing = rows().get(key(entity.partitionKey, entity.rowKey));
      if (!existing) throw Object.assign(new Error('not found'), { statusCode: 404 });
      const merged = mode === 'Merge' ? { ...existing, ...entity } : { ...entity };
      rows().set(key(entity.partitionKey, entity.rowKey), { ...merged, etag: `e${++etagCounter}` });
    },
    async deleteEntity(pk: string, rk: string, opts?: { etag?: string }) {
      const existing = rows().get(key(pk, rk));
      if (!existing) throw Object.assign(new Error('not found'), { statusCode: 404 });
      if (opts?.etag && opts.etag !== existing.etag) {
        throw Object.assign(new Error('precondition failed'), { statusCode: 412 });
      }
      rows().delete(key(pk, rk));
    },
    listEntities() {
      calls.list++;
      const snapshot = [...rows().values()].map((r) => ({ ...r }));
      return (async function* () { yield* snapshot; })();
    },
  };
}

jest.mock('@azure/data-tables', () => ({
  TableClient: { fromConnectionString: (_cs: string, name: string) => makeTableClient(name) },
  TableServiceClient: { fromConnectionString: () => ({ createTable: async () => undefined }) },
}));

import {
  saveSession,
  loadSessionByToken,
  migrateLegacySessionRows,
  type StoredSession,
} from '../services/tableStorage.js';
import { encryptWithDek, envelopeAad, hashSessionToken } from '../services/credentialCrypto.js';

const sessions = () => tableFor('mcpSessions');

/** A pre-multi-session row: RowKey = userId, envelope bound to that RowKey. */
function putLegacyRow(userId: string, token: string | null, accessToken: string): void {
  const envelope = encryptWithDek(accessToken, envelopeAad('mcpSessions', 'session', userId, 'accessToken'));
  sessions().set(`session|${userId}`, {
    partitionKey: 'session',
    rowKey: userId,
    homeAccountId: `${userId}.home`,
    displayName: userId,
    email: `${userId}@example.com`,
    tenantId: 'tenant',
    expiresAt: Date.now() + 3600_000,
    sessionCreatedAt: 1_000,
    accessTokenCiphertext: envelope.ciphertext,
    accessTokenIv: envelope.iv,
    accessTokenAuthTag: envelope.authTag,
    ...(token ? { sessionTokenHash: hashSessionToken(token) } : {}),
    etag: `e${++etagCounter}`,
  });
}

function session(userId: string, sessionToken: string): StoredSession {
  return {
    userId,
    homeAccountId: `${userId}.home`,
    displayName: userId,
    email: `${userId}@example.com`,
    tenantId: 'tenant',
    accessToken: `graph-${userId}`,
    expiresAt: Date.now() + 3600_000,
    sessionToken,
    sessionCreatedAt: Date.now(),
  };
}

beforeEach(() => {
  tables.clear();
  getFault = null;
  jest.spyOn(console, 'log').mockImplementation(() => {});
  jest.spyOn(console, 'warn').mockImplementation(() => {});
});

afterEach(() => {
  jest.restoreAllMocks();
});

describe('legacy session rows', () => {
  // Rows present before the first storage call are migrated by the
  // first-call initialisation, the way a process start sees them.
  test('are re-keyed before the first token lookup and stay reachable', async () => {
    putLegacyRow('legacy-user', 'legacy-token', 'graph-legacy');
    putLegacyRow('orphan-user', null, 'graph-orphan');

    const loaded = await loadSessionByToken('legacy-token');
    expect(loaded?.userId).toBe('legacy-user');
    expect(loaded?.accessToken).toBe('graph-legacy');
    expect(loaded?.sessionCreatedAt).toBe(1_000);

    const newKey = hashSessionToken('legacy-token').slice(0, 32);
    expect(loaded?._storageKey).toBe(newKey);
    expect(sessions().has('session|legacy-user')).toBe(false);
    // No token can reach a row without a token hash, so it is dropped.
    expect(sessions().has('session|orphan-user')).toBe(false);
    expect([...sessions().keys()]).toEqual([`session|${newKey}`]);
  });

  test('re-keyed rows are bound to their identity columns', async () => {
    putLegacyRow('bound-user', 'bound-token', 'graph-bound');
    await expect(migrateLegacySessionRows()).resolves.toEqual({ migrated: 1, dropped: 0 });

    // With the row-only fallback switched off, only a fully bound envelope
    // decrypts.
    process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING = 'true';
    try {
      expect((await loadSessionByToken('bound-token'))?.accessToken).toBe('graph-bound');

      const newKey = hashSessionToken('bound-token').slice(0, 32);
      sessions().get(`session|${newKey}`)!.homeAccountId = 'someone-else.home';
      await expect(loadSessionByToken('bound-token')).rejects.toThrow();
    } finally {
      delete process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING;
    }
  });

  test('a row whose envelope no longer decrypts is dropped', async () => {
    putLegacyRow('bad-user', 'bad-token', 'graph-bad');
    const row = sessions().get('session|bad-user')!;
    row.accessTokenAuthTag = Buffer.alloc(16).toString('base64');

    await expect(migrateLegacySessionRows()).resolves.toEqual({ migrated: 0, dropped: 1 });
    expect(sessions().size).toBe(0);
  });

  test('a row already at the new key wins over the legacy copy', async () => {
    await saveSession(session('user-a', 'tok-a'));
    putLegacyRow('user-a', 'tok-a', 'stale-graph-token');

    await expect(migrateLegacySessionRows()).resolves.toEqual({ migrated: 1, dropped: 0 });
    expect(sessions().has('session|user-a')).toBe(false);
    expect((await loadSessionByToken('tok-a'))?.accessToken).toBe('graph-user-a');
  });

  test('current rows are left alone', async () => {
    await saveSession(session('user-b', 'tok-b'));
    const before = new Map(sessions());

    await expect(migrateLegacySessionRows()).resolves.toEqual({ migrated: 0, dropped: 0 });
    expect(sessions()).toEqual(before);
  });
});

describe('unknown bearer token', () => {
  test('costs one point read and no scan', async () => {
    await saveSession(session('user-c', 'tok-c'));
    calls.list = 0;
    calls.get = 0;

    await expect(loadSessionByToken('not-a-session-token')).resolves.toBeNull();
    expect(calls.list).toBe(0);
    expect(calls.get).toBe(1);
  });

  test('a row at the same key with a different hash is a miss', async () => {
    await saveSession(session('user-d', 'tok-d'));
    const rowKey = hashSessionToken('tok-d').slice(0, 32);
    sessions().get(`session|${rowKey}`)!.sessionTokenHash = 'f'.repeat(64);

    await expect(loadSessionByToken('tok-d')).resolves.toBeNull();
  });
});

describe('storage failure on the point read', () => {
  // Only a 404 is a token miss. Anything else must not turn a valid token
  // into an unknown one.
  test.each([
    ['throttled (429)', Object.assign(new Error('too many requests'), { statusCode: 429 })],
    ['unavailable (503)', Object.assign(new Error('server busy'), { statusCode: 503 })],
    ['forbidden (403)', Object.assign(new Error('auth failed'), { statusCode: 403 })],
    ['network error (no status)', Object.assign(new Error('socket hang up'), { code: 'ECONNRESET' })],
  ])('%s is rethrown, not reported as a miss', async (_label, fault) => {
    await saveSession(session('user-e', 'tok-e'));
    getFault = fault;

    await expect(loadSessionByToken('tok-e')).rejects.toBe(fault);
    // The row is still there and the next read finds it.
    expect((await loadSessionByToken('tok-e'))?.userId).toBe('user-e');
  });
});
