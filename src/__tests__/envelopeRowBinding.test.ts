/**
 * Credential envelopes in Table Storage are bound to their row with
 * AES-GCM additional authenticated data. Exercises tableStorage.ts against an
 * in-memory table to show (1) an access-token envelope copied from user A's
 * row into user B's row no longer decrypts, and (2) envelopes written before
 * the binding are still read and get rewritten bound on that read, and (3) the
 * identity columns beside the envelope (userId, homeAccountId, tenantId) are
 * covered by the same tag, so a row whose identity was edited fails
 * authentication instead of steering a silent refresh at another user. Also the
 * per-account MSAL cache rows: binding, ETag-conditional writes, and the split
 * of the old shared row.
 */

import { jest } from '@jest/globals';
import { randomBytes, createCipheriv } from 'crypto';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');
process.env.MCP_DATA_ENCRYPTION_KEY = randomBytes(32).toString('hex');
process.env.AZURE_STORAGE_CONNECTION_STRING =
  'DefaultEndpointsProtocol=https;AccountName=test;AccountKey=dGVzdA==;EndpointSuffix=core.windows.net';

// ── In-memory Azure Table Storage ────────────────────────────────────────────

type Entity = Record<string, unknown> & { partitionKey: string; rowKey: string; etag?: string };
const tables = new Map<string, Map<string, Entity>>();
let etagCounter = 0;

function tableFor(name: string): Map<string, Entity> {
  if (!tables.has(name)) tables.set(name, new Map());
  return tables.get(name)!;
}

function makeTableClient(name: string) {
  const rows = () => tableFor(name);
  const key = (pk: string, rk: string) => `${pk}|${rk}`;
  return {
    async getEntity(pk: string, rk: string) {
      const row = rows().get(key(pk, rk));
      if (!row) throw Object.assign(new Error('not found'), { statusCode: 404 });
      return { ...row };
    },
    async upsertEntity(entity: Entity, mode: string) {
      const existing = rows().get(key(entity.partitionKey, entity.rowKey));
      const merged = mode === 'Merge' && existing ? { ...existing, ...entity } : { ...entity };
      rows().set(key(entity.partitionKey, entity.rowKey), { ...merged, etag: `e${++etagCounter}` });
    },
    async updateEntity(entity: Entity, mode: string, opts?: { etag?: string }) {
      const existing = rows().get(key(entity.partitionKey, entity.rowKey));
      if (!existing) throw Object.assign(new Error('not found'), { statusCode: 404 });
      if (opts?.etag && opts.etag !== '*' && opts.etag !== existing.etag) {
        throw Object.assign(new Error('precondition failed'), { statusCode: 412 });
      }
      const merged = mode === 'Merge' ? { ...existing, ...entity } : { ...entity };
      const etag = `e${++etagCounter}`;
      rows().set(key(entity.partitionKey, entity.rowKey), { ...merged, etag });
      return { etag };
    },
    async createEntity(entity: Entity) {
      if (rows().has(key(entity.partitionKey, entity.rowKey))) {
        throw Object.assign(new Error('conflict'), { statusCode: 409 });
      }
      const etag = `e${++etagCounter}`;
      rows().set(key(entity.partitionKey, entity.rowKey), { ...entity, etag });
      return { etag };
    },
    async deleteEntity(pk: string, rk: string, opts?: { etag?: string }) {
      const existing = rows().get(key(pk, rk));
      if (existing && opts?.etag && opts.etag !== existing.etag) {
        throw Object.assign(new Error('precondition failed'), { statusCode: 412 });
      }
      rows().delete(key(pk, rk));
    },
    listEntities() {
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
  saveMsalCachePartition,
  loadMsalCachePartition,
  splitLegacyMsalCache,
  resetLegacyMsalSplitForTests,
  MsalCacheConflictError,
  type StoredSession,
} from '../services/tableStorage.js';
import { encryptWithDek, envelopeAad } from '../services/credentialCrypto.js';

function session(userId: string, sessionToken: string, accessToken: string): StoredSession {
  return {
    userId,
    homeAccountId: `${userId}.home`,
    displayName: userId,
    email: `${userId}@example.com`,
    tenantId: 'tenant',
    accessToken,
    expiresAt: Date.now() + 3600_000,
    sessionToken,
    sessionCreatedAt: Date.now(),
  };
}

function sessionRow(userId: string): Entity {
  return [...tableFor('mcpSessions').values()].find((r) => r.userId === userId)!;
}

/** The pre-binding envelope shape: AES-GCM under the DEK with no AAD. */
function encryptUnbound(plaintext: string) {
  const iv = randomBytes(12);
  const dek = Buffer.from(process.env.MCP_DATA_ENCRYPTION_KEY!, 'hex');
  const cipher = createCipheriv('aes-256-gcm', dek, iv);
  const enc = Buffer.concat([cipher.update(plaintext, 'utf8'), cipher.final()]);
  return {
    ciphertext: enc.toString('base64'),
    iv: iv.toString('base64'),
    authTag: cipher.getAuthTag().toString('base64'),
  };
}

const flush = () => new Promise((resolve) => setImmediate(resolve));

beforeEach(() => {
  tables.clear();
  resetLegacyMsalSplitForTests();
  delete process.env.MCP_ENVELOPE_REQUIRE_AAD;
  delete process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING;
});

describe('session access token envelope binding', () => {
  test('round-trips through its own row', async () => {
    await saveSession(session('userA', 'tok-A', 'graph-token-A'));
    const loaded = await loadSessionByToken('tok-A');
    expect(loaded?.accessToken).toBe('graph-token-A');
  });

  test("user A's envelope copied into user B's row does not decrypt", async () => {
    await saveSession(session('userA', 'tok-A', 'graph-token-A'));
    await saveSession(session('userB', 'tok-B', 'graph-token-B'));

    const rowA = sessionRow('userA');
    const rowB = sessionRow('userB');
    tableFor('mcpSessions').set(`session|${rowB.rowKey}`, {
      ...rowB,
      accessTokenCiphertext: rowA.accessTokenCiphertext,
      accessTokenIv: rowA.accessTokenIv,
      accessTokenAuthTag: rowA.accessTokenAuthTag,
    });

    await expect(loadSessionByToken('tok-B')).rejects.toThrow();
  });

  test('a pre-binding unbound envelope is read and rewritten bound', async () => {
    await saveSession(session('userA', 'tok-A', 'placeholder'));
    const row = sessionRow('userA');
    const legacy = encryptUnbound('legacy-graph-token');
    tableFor('mcpSessions').set(`session|${row.rowKey}`, {
      ...row,
      accessTokenCiphertext: legacy.ciphertext,
      accessTokenIv: legacy.iv,
      accessTokenAuthTag: legacy.authTag,
    });

    expect((await loadSessionByToken('tok-A'))?.accessToken).toBe('legacy-graph-token');
    await flush();

    const rebound = sessionRow('userA');
    expect(rebound.accessTokenCiphertext).not.toBe(legacy.ciphertext);
    // Now bound: readable with legacy reads switched off.
    process.env.MCP_ENVELOPE_REQUIRE_AAD = 'true';
    expect((await loadSessionByToken('tok-A'))?.accessToken).toBe('legacy-graph-token');
  });

  test('unbound envelopes are refused once MCP_ENVELOPE_REQUIRE_AAD=true', async () => {
    await saveSession(session('userA', 'tok-A', 'placeholder'));
    const row = sessionRow('userA');
    const legacy = encryptUnbound('legacy-graph-token');
    tableFor('mcpSessions').set(`session|${row.rowKey}`, {
      ...row,
      accessTokenCiphertext: legacy.ciphertext,
      accessTokenIv: legacy.iv,
      accessTokenAuthTag: legacy.authTag,
    });

    process.env.MCP_ENVELOPE_REQUIRE_AAD = 'true';
    await expect(loadSessionByToken('tok-A')).rejects.toThrow();
  });
});

describe('session identity column binding', () => {
  /** Edit plaintext columns on a stored row, leaving the envelope alone. */
  function tamper(userId: string, edits: Record<string, unknown>): void {
    const row = sessionRow(userId);
    tableFor('mcpSessions').set(`session|${row.rowKey}`, { ...row, ...edits });
  }

  /** The row-only envelope shape written before identity binding. */
  function rowBoundRow(userId: string, accessToken: string): Entity {
    const row = sessionRow(userId);
    const envelope = encryptWithDek(
      accessToken,
      envelopeAad('mcpSessions', 'session', row.rowKey, 'accessToken'),
    );
    const updated = {
      ...row,
      accessTokenCiphertext: envelope.ciphertext,
      accessTokenIv: envelope.iv,
      accessTokenAuthTag: envelope.authTag,
    };
    tableFor('mcpSessions').set(`session|${row.rowKey}`, updated);
    return updated;
  }

  test("repointing homeAccountId at another user's MSAL account fails authentication", async () => {
    await saveSession(session('victim', 'tok-V', 'graph-token-V'));
    await saveSession(session('attacker', 'tok-X', 'graph-token-X'));

    // The attack from threat model row 2.3: the attacker holds a valid session
    // and storage write, and points their row at the victim's MSAL account so
    // the next silent refresh would mint the victim's token into it.
    tamper('attacker', { homeAccountId: sessionRow('victim').homeAccountId });

    await expect(loadSessionByToken('tok-X')).rejects.toThrow();
  });

  test.each([
    ['userId', 'someone-else'],
    ['tenantId', 'other-tenant'],
    ['homeAccountId', 'someone-else.home'],
  ])('editing %s fails authentication', async (column, value) => {
    await saveSession(session('userA', 'tok-A', 'graph-token-A'));
    tamper('userA', { [column]: value });
    await expect(loadSessionByToken('tok-A')).rejects.toThrow();
  });

  test('profile columns outside the binding can still change', async () => {
    await saveSession(session('userA', 'tok-A', 'graph-token-A'));
    tamper('userA', { displayName: 'Renamed' });
    const loaded = await loadSessionByToken('tok-A');
    expect(loaded?.displayName).toBe('Renamed');
    expect(loaded?.accessToken).toBe('graph-token-A');
  });

  test('a refresh write keeps the row authenticating', async () => {
    await saveSession(session('userA', 'tok-A', 'graph-token-A'));
    const loaded = (await loadSessionByToken('tok-A'))!;
    await saveSession({ ...loaded, sessionToken: '', accessToken: 'graph-token-A2' }, 'update');

    process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING = 'true';
    expect((await loadSessionByToken('tok-A'))?.accessToken).toBe('graph-token-A2');
  });

  test('a row-bound envelope is read and rewritten identity-bound', async () => {
    await saveSession(session('userA', 'tok-A', 'placeholder'));
    const legacy = rowBoundRow('userA', 'row-bound-token');

    expect((await loadSessionByToken('tok-A'))?.accessToken).toBe('row-bound-token');
    await flush();
    expect(sessionRow('userA').accessTokenCiphertext).not.toBe(legacy.accessTokenCiphertext);

    process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING = 'true';
    expect((await loadSessionByToken('tok-A'))?.accessToken).toBe('row-bound-token');

    // Once rebound, editing the identity fails even with legacy reads on.
    delete process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING;
    tamper('userA', { homeAccountId: 'someone-else.home' });
    await expect(loadSessionByToken('tok-A')).rejects.toThrow();
  });

  test('MCP_SESSION_REQUIRE_IDENTITY_BINDING=true refuses a replayed row-bound envelope', async () => {
    await saveSession(session('victim', 'tok-V', 'graph-token-V'));
    await saveSession(session('attacker', 'tok-X', 'placeholder'));
    // An envelope the attacker kept from before the rebind, put back with a
    // repointed identity.
    rowBoundRow('attacker', 'graph-token-X');
    tamper('attacker', { homeAccountId: sessionRow('victim').homeAccountId });

    process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING = 'true';
    await expect(loadSessionByToken('tok-X')).rejects.toThrow();
  });

  test('MCP_SESSION_REQUIRE_IDENTITY_BINDING=true also refuses unbound envelopes', async () => {
    await saveSession(session('userA', 'tok-A', 'placeholder'));
    const legacy = encryptUnbound('legacy-graph-token');
    tamper('userA', {
      accessTokenCiphertext: legacy.ciphertext,
      accessTokenIv: legacy.iv,
      accessTokenAuthTag: legacy.authTag,
    });

    process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING = 'true';
    await expect(loadSessionByToken('tok-A')).rejects.toThrow();
  });

  test('MCP_SESSION_REQUIRE_IDENTITY_BINDING=true refuses a row with no envelope', async () => {
    await saveSession(session('userA', 'tok-A', 'graph-token-A'));
    tamper('userA', {
      accessTokenCiphertext: undefined,
      accessTokenIv: undefined,
      accessTokenAuthTag: undefined,
      homeAccountId: 'someone-else.home',
    });

    expect((await loadSessionByToken('tok-A'))?.accessToken).toBe('');
    process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING = 'true';
    await expect(loadSessionByToken('tok-A')).rejects.toThrow();
  });
});

const HOME_A = '00000000-0000-4000-8000-00000000000a.72f988bf-0000-4000-8000-00000000c0de';
const HOME_B = '00000000-0000-4000-8000-00000000000b.72f988bf-0000-4000-8000-00000000c0de';
const LEGACY_AAD = envelopeAad('mcpMsalCache', 'cache', 'msal-token-cache', 'msalCache');

function accountCache(homeAccountId: string, rt: string) {
  return {
    Account: { [`${homeAccountId}-login.windows.net-tenant`]: { home_account_id: homeAccountId } },
    IdToken: {},
    AccessToken: {},
    RefreshToken: { [`${homeAccountId}-rt`]: { home_account_id: homeAccountId, secret: rt } },
    AppMetadata: { 'appmetadata-login.windows.net-client': { client_id: 'client' } },
  };
}

function setLegacyRow(envelope: { ciphertext: string; iv: string; authTag: string }) {
  tableFor('mcpMsalCache').set('cache|msal-token-cache', {
    partitionKey: 'cache', rowKey: 'msal-token-cache', ...envelope, etag: 'legacy-e',
  });
}

describe('MSAL cache rows, one per account', () => {
  test('round-trip through the account row, with the ETag the next write needs', async () => {
    const etag = await saveMsalCachePartition(HOME_A, '{"a":1}', undefined);
    const row = await loadMsalCachePartition(HOME_A);
    expect(row).toEqual({ data: '{"a":1}', etag });
    expect(tableFor('mcpMsalCache').has(`account|${HOME_A}`)).toBe(true);
  });

  test('a missing row is empty with no ETag', async () => {
    expect(await loadMsalCachePartition(HOME_A)).toEqual({ data: null, etag: undefined });
  });

  test("account A's envelope copied into account B's row does not decrypt", async () => {
    await saveMsalCachePartition(HOME_A, '{"a":1}', undefined);
    await saveMsalCachePartition(HOME_B, '{"b":1}', undefined);
    const table = tableFor('mcpMsalCache');
    const a = table.get(`account|${HOME_A}`)!;
    const b = table.get(`account|${HOME_B}`)!;
    table.set(`account|${HOME_B}`, { ...b, ciphertext: a.ciphertext, iv: a.iv, authTag: a.authTag });
    jest.spyOn(console, 'error').mockImplementation(() => undefined);

    const row = await loadMsalCachePartition(HOME_B);
    expect(row.data).toBeNull();
    // The ETag comes back so a fresh sign-in can replace the bad row.
    expect(row.etag).toBe(b.etag);
  });

  test('the shared row envelope substituted into an account row does not decrypt', async () => {
    const shared = encryptWithDek('{"shared":1}', LEGACY_AAD);
    tableFor('mcpMsalCache').set(`account|${HOME_A}`, {
      partitionKey: 'account', rowKey: HOME_A, ...shared, etag: 'x',
    });
    jest.spyOn(console, 'error').mockImplementation(() => undefined);
    expect((await loadMsalCachePartition(HOME_A)).data).toBeNull();
  });

  test('an unbound envelope in an account row is not read, even while legacy reads are allowed', async () => {
    tableFor('mcpMsalCache').set(`account|${HOME_A}`, {
      partitionKey: 'account', rowKey: HOME_A, ...encryptUnbound('{"a":1}'), etag: 'x',
    });
    jest.spyOn(console, 'error').mockImplementation(() => undefined);
    expect((await loadMsalCachePartition(HOME_A)).data).toBeNull();
  });

  test('a write with a stale ETag loses and leaves the newer row', async () => {
    const first = await saveMsalCachePartition(HOME_A, '{"v":1}', undefined);
    const second = await saveMsalCachePartition(HOME_A, '{"v":2}', first);

    await expect(saveMsalCachePartition(HOME_A, '{"v":"stale"}', first)).rejects.toBeInstanceOf(
      MsalCacheConflictError,
    );
    expect(await loadMsalCachePartition(HOME_A)).toEqual({ data: '{"v":2}', etag: second });
  });

  test('creating a row that another writer already created loses', async () => {
    await saveMsalCachePartition(HOME_A, '{"winner":1}', undefined);
    await expect(saveMsalCachePartition(HOME_A, '{"loser":1}', undefined)).rejects.toBeInstanceOf(
      MsalCacheConflictError,
    );
    expect((await loadMsalCachePartition(HOME_A)).data).toBe('{"winner":1}');
  });

  test('writing to one account never touches another account', async () => {
    await saveMsalCachePartition(HOME_A, '{"a":1}', undefined);
    const b = await saveMsalCachePartition(HOME_B, '{"b":1}', undefined);
    await saveMsalCachePartition(HOME_B, '{"b":2}', b);
    expect((await loadMsalCachePartition(HOME_A)).data).toBe('{"a":1}');
  });

  test('refuses a home account id that is not safe as a row key', async () => {
    await expect(saveMsalCachePartition('a/b#c', '{}', undefined)).rejects.toThrow('malformed home account id');
  });
});

describe('splitting the old shared MSAL cache row', () => {
  const shared = JSON.stringify({
    Account: { ...accountCache(HOME_A, 'rt-a').Account, ...accountCache(HOME_B, 'rt-b').Account },
    IdToken: {},
    AccessToken: {},
    RefreshToken: { ...accountCache(HOME_A, 'rt-a').RefreshToken, ...accountCache(HOME_B, 'rt-b').RefreshToken },
    AppMetadata: accountCache(HOME_A, '').AppMetadata,
  });

  beforeEach(() => {
    jest.spyOn(console, 'log').mockImplementation(() => undefined);
  });

  test('gives each account its own row and deletes the shared row on first load', async () => {
    setLegacyRow(encryptWithDek(shared, LEGACY_AAD));

    const a = JSON.parse((await loadMsalCachePartition(HOME_A)).data!);
    expect(a).toEqual(accountCache(HOME_A, 'rt-a'));
    const b = JSON.parse((await loadMsalCachePartition(HOME_B)).data!);
    expect(b).toEqual(accountCache(HOME_B, 'rt-b'));
    expect(tableFor('mcpMsalCache').has('cache|msal-token-cache')).toBe(false);
  });

  test('reads a shared row written before AAD binding', async () => {
    setLegacyRow(encryptUnbound(shared));
    expect(await splitLegacyMsalCache()).toBe(2);
    expect(JSON.parse((await loadMsalCachePartition(HOME_B)).data!)).toEqual(accountCache(HOME_B, 'rt-b'));
  });

  test('keeps an account row that already exists rather than the older shared copy', async () => {
    await saveMsalCachePartition(HOME_A, JSON.stringify(accountCache(HOME_A, 'rt-a-newer')), undefined);
    setLegacyRow(encryptWithDek(shared, LEGACY_AAD));

    expect(await splitLegacyMsalCache()).toBe(1);
    expect(JSON.parse((await loadMsalCachePartition(HOME_A)).data!)).toEqual(accountCache(HOME_A, 'rt-a-newer'));
  });

  test('leaves the shared row if a replica on the previous release wrote it during the split', async () => {
    setLegacyRow(encryptWithDek(shared, LEGACY_AAD));
    const table = tableFor('mcpMsalCache');
    const realGet = table.get.bind(table);
    const spy = jest.spyOn(table, 'get').mockImplementation((k: string) => {
      const row = realGet(k);
      if (k === 'cache|msal-token-cache' && row) {
        spy.mockRestore();
        table.set(k, { ...row, etag: 'written-meanwhile' });
      }
      return row;
    });

    expect(await splitLegacyMsalCache()).toBe(2);
    expect(table.has('cache|msal-token-cache')).toBe(true);
    // The next process to start finishes the job.
    expect(await splitLegacyMsalCache()).toBe(0);
    expect(table.has('cache|msal-token-cache')).toBe(false);
  });

  test('leaves a shared row it cannot decrypt in place', async () => {
    setLegacyRow(encryptWithDek(shared, envelopeAad('mcpMsalCache', 'cache', 'other', 'msalCache')));
    process.env.MCP_ENVELOPE_REQUIRE_AAD = 'true';
    jest.spyOn(console, 'error').mockImplementation(() => undefined);

    expect(await splitLegacyMsalCache()).toBe(0);
    expect(tableFor('mcpMsalCache').has('cache|msal-token-cache')).toBe(true);
  });

  test('does nothing when there is no shared row', async () => {
    expect(await splitLegacyMsalCache()).toBe(0);
  });
});
