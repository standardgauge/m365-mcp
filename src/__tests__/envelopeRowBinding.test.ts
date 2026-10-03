/**
 * AC-374: credential envelopes in Table Storage are bound to their row with
 * AES-GCM additional authenticated data. Exercises tableStorage.ts against an
 * in-memory table to show (1) an access-token envelope copied from user A's
 * row into user B's row no longer decrypts, and (2) envelopes written before
 * the binding are still read and get rewritten bound on that read.
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
      if (opts?.etag && opts.etag !== existing.etag) {
        throw Object.assign(new Error('precondition failed'), { statusCode: 412 });
      }
      const merged = mode === 'Merge' ? { ...existing, ...entity } : { ...entity };
      rows().set(key(entity.partitionKey, entity.rowKey), { ...merged, etag: `e${++etagCounter}` });
    },
    async deleteEntity(pk: string, rk: string) {
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
  saveMsalCache,
  loadMsalCache,
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

/** The pre-AC-374 envelope shape: AES-GCM under the DEK with no AAD. */
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
  delete process.env.MCP_ENVELOPE_REQUIRE_AAD;
});

describe('session access token envelope binding (AC-374)', () => {
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

  test('a pre-AC-374 unbound envelope is read and rewritten bound', async () => {
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

describe('MSAL cache envelope binding (AC-374)', () => {
  test('round-trips', async () => {
    await saveMsalCache('{"cache":1}');
    expect(await loadMsalCache()).toBe('{"cache":1}');
  });

  test('a session access token envelope substituted into the cache row does not decrypt', async () => {
    await saveSession(session('userA', 'tok-A', 'graph-token-A'));
    const rowA = sessionRow('userA');
    tableFor('mcpMsalCache').set('cache|msal-token-cache', {
      partitionKey: 'cache',
      rowKey: 'msal-token-cache',
      ciphertext: rowA.accessTokenCiphertext,
      iv: rowA.accessTokenIv,
      authTag: rowA.accessTokenAuthTag,
      etag: 'x',
    });
    jest.spyOn(console, 'error').mockImplementation(() => undefined);
    expect(await loadMsalCache()).toBeNull();
  });

  test('a pre-AC-374 unbound cache is read and rewritten bound', async () => {
    const legacy = encryptUnbound('{"legacy":true}');
    tableFor('mcpMsalCache').set('cache|msal-token-cache', {
      partitionKey: 'cache', rowKey: 'msal-token-cache', ...legacy, etag: 'e0',
    });

    expect(await loadMsalCache()).toBe('{"legacy":true}');
    const rebound = tableFor('mcpMsalCache').get('cache|msal-token-cache')!;
    expect(rebound.ciphertext).not.toBe(legacy.ciphertext);

    process.env.MCP_ENVELOPE_REQUIRE_AAD = 'true';
    expect(await loadMsalCache()).toBe('{"legacy":true}');
  });

  test('the rebind does not overwrite a newer cache written concurrently', async () => {
    const legacy = encryptUnbound('{"old":true}');
    const table = tableFor('mcpMsalCache');
    table.set('cache|msal-token-cache', {
      partitionKey: 'cache', rowKey: 'msal-token-cache', ...legacy, etag: 'e0',
    });
    // Another replica saves a newer cache between our read and our rebind.
    const newer = encryptWithDek(
      '{"new":true}',
      envelopeAad('mcpMsalCache', 'cache', 'msal-token-cache', 'msalCache'),
    );
    const realGet = table.get.bind(table);
    let raced = false;
    const spy = jest.spyOn(table, 'get').mockImplementation((k: string) => {
      const row = realGet(k);
      if (!raced && row) {
        raced = true;
        table.set(k, { ...row, ...newer, etag: 'e-newer' });
      }
      return row;
    });
    jest.spyOn(console, 'warn').mockImplementation(() => undefined);

    expect(await loadMsalCache()).toBe('{"old":true}');
    spy.mockRestore();
    expect(await loadMsalCache()).toBe('{"new":true}');
  });
});
