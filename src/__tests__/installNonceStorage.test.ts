/**
 * mcpInstallNonces holds the install handoff: the session token waits there
 * between the OAuth callback and the installer's poll. Exercises
 * tableStorage.ts against an in-memory table to show (1) the token is stored
 * as a row-bound envelope, never in plaintext, (2) a pre-encryption plaintext
 * row or a moved envelope is never handed out, and (3) rows nobody polls are
 * deleted by the server once they expire.
 */

import { jest } from '@jest/globals';
import { createHash, randomBytes } from 'crypto';

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
    async upsertEntity(entity: Entity) {
      rows().set(key(entity.partitionKey, entity.rowKey), { ...entity, etag: `e${++etagCounter}` });
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
  attachSessionToInstallNonce,
  consumeInstallNonce,
  purgeExpiredInstallNonces,
  maybePurgeExpiredInstallNonces,
  resetInstallNoncePurgeStateForTests,
} from '../services/tableStorage.js';

const nonceRows = () => tableFor('mcpInstallNonces');
const rowKeyFor = (nonce: string) => createHash('sha256').update(nonce).digest('hex');
const challenge = () => randomBytes(32).toString('hex');

function record(sessionToken: string, expiresAt = Date.now() + 5 * 60_000) {
  return { sessionToken, userId: 'u1', email: 'u1@example.com', displayName: 'User One', expiresAt };
}

const settle = () => new Promise((r) => setImmediate(r));

beforeEach(async () => {
  nonceRows().clear();
  await resetInstallNoncePurgeStateForTests();
  // Start every test inside a purge interval so only the purges a test asks
  // for run.
  maybePurgeExpiredInstallNonces();
  await settle();
});

describe('install nonce rows hold the session token encrypted', () => {
  it('stores no plaintext token and returns it once on consume', async () => {
    const nonce = challenge();
    const token = randomBytes(32).toString('base64url');
    expect(await attachSessionToInstallNonce(nonce, record(token))).toBe(true);

    const row = nonceRows().get(`nonce|${rowKeyFor(nonce)}`)!;
    expect(row).toBeDefined();
    expect(row).not.toHaveProperty('sessionToken');
    expect(JSON.stringify(row)).not.toContain(token);
    expect(row.sessionTokenCiphertext).toEqual(expect.any(String));
    expect(row.sessionTokenIv).toEqual(expect.any(String));
    expect(row.sessionTokenAuthTag).toEqual(expect.any(String));

    const consumed = await consumeInstallNonce(nonce);
    expect(consumed).toMatchObject({ sessionToken: token, userId: 'u1', email: 'u1@example.com' });
    expect(nonceRows().size).toBe(0);
    expect(await consumeInstallNonce(nonce)).toBe('pending');
  });

  it('refuses an envelope moved in from another nonce row, and deletes it', async () => {
    const victim = challenge();
    const attacker = challenge();
    await attachSessionToInstallNonce(victim, record('victim-token'));
    await attachSessionToInstallNonce(attacker, record('attacker-token'));

    const victimRow = nonceRows().get(`nonce|${rowKeyFor(victim)}`)!;
    const attackerKey = `nonce|${rowKeyFor(attacker)}`;
    nonceRows().set(attackerKey, {
      ...nonceRows().get(attackerKey)!,
      sessionTokenCiphertext: victimRow.sessionTokenCiphertext,
      sessionTokenIv: victimRow.sessionTokenIv,
      sessionTokenAuthTag: victimRow.sessionTokenAuthTag,
    });

    expect(await consumeInstallNonce(attacker)).toBeNull();
    expect(nonceRows().has(attackerKey)).toBe(false);
  });

  it('never hands out a plaintext row written before encryption', async () => {
    const nonce = challenge();
    const key = `nonce|${rowKeyFor(nonce)}`;
    nonceRows().set(key, {
      partitionKey: 'nonce',
      rowKey: rowKeyFor(nonce),
      sessionToken: 'legacy-plaintext-token',
      userId: 'u1',
      email: 'u1@example.com',
      displayName: 'User One',
      expiresAt: Date.now() + 60_000,
      etag: 'legacy',
    });

    expect(await consumeInstallNonce(nonce)).toBeNull();
    expect(nonceRows().has(key)).toBe(false);
  });

  it('returns null for an expired row and deletes it', async () => {
    const nonce = challenge();
    await attachSessionToInstallNonce(nonce, record('t', Date.now() - 1));
    expect(await consumeInstallNonce(nonce)).toBeNull();
    expect(nonceRows().size).toBe(0);
  });
});

describe('expired install nonce rows are purged by the server', () => {
  it('deletes expired and timestamp-less rows, keeps live ones', async () => {
    const now = Date.now();
    const live = challenge();
    const abandoned = challenge();
    await attachSessionToInstallNonce(live, record('live', now + 60_000));
    await attachSessionToInstallNonce(abandoned, record('abandoned', now - 60_000));
    nonceRows().set('nonce|no-expiry', { partitionKey: 'nonce', rowKey: 'no-expiry', sessionToken: 'old', etag: 'x' });

    expect(await purgeExpiredInstallNonces(now)).toBe(2);
    expect([...nonceRows().keys()]).toEqual([`nonce|${rowKeyFor(live)}`]);
  });

  it('a callback starts the purge, so an abandoned row goes without anyone polling it', async () => {
    const abandoned = challenge();
    await attachSessionToInstallNonce(abandoned, record('abandoned', Date.now() - 1));
    expect(nonceRows().size).toBe(1);

    await resetInstallNoncePurgeStateForTests(); // interval elapsed
    await attachSessionToInstallNonce(challenge(), record('next-install'));
    await resetInstallNoncePurgeStateForTests(); // wait for the background purge

    expect(nonceRows().has(`nonce|${rowKeyFor(abandoned)}`)).toBe(false);
    expect(nonceRows().size).toBe(1);
  });

  it('runs at most once per interval per process', async () => {
    const first = challenge();
    await attachSessionToInstallNonce(first, record('a', Date.now() - 1));
    await resetInstallNoncePurgeStateForTests();
    maybePurgeExpiredInstallNonces();
    await settle();

    // Inside the interval: a second expired row is left for the next run.
    const second = challenge();
    await attachSessionToInstallNonce(second, record('b', Date.now() - 1));
    await settle();
    expect(nonceRows().has(`nonce|${rowKeyFor(first)}`)).toBe(false);
    expect(nonceRows().has(`nonce|${rowKeyFor(second)}`)).toBe(true);
  });
});
