/**
 * Per-key deny entries for agent keys.
 *
 * A key's entries are tier 2 rows under the subject `key:<keyId>`. Every deny
 * check made while handling a request that authenticated with the key reads
 * them, because the key id rides on the request scope (clientAddress.ts) and
 * loadDeniedPaths adds it. The REAL denyList and clientAddress services are in
 * the loop; Azure Table Storage is an in-memory fake.
 *
 * Covers:
 *   - isPathDenied, filterDeniedPaths and filterDeniedSearchHits apply a key's
 *     entries inside a request marked with that key, on top of the owner's
 *   - the same checks ignore the key's entries for the owner's own session,
 *     for a different key, and outside a request scope
 *   - delegated subjects (owner + caller) still combine with the key
 *   - clearAgentKeyDenyEntries removes the key's rows of every type
 */
import { jest } from '@jest/globals';
import type { HttpRequest } from '@azure/functions';

// ── In-memory Azure Table Storage ───────────────────────────────────────────

interface Row { partitionKey: string; rowKey: string; path: string }
const tables = new Map<string, Row[]>();
function rows(name: string): Row[] {
  if (!tables.has(name)) tables.set(name, []);
  return tables.get(name)!;
}

jest.mock('@azure/data-tables', () => {
  const makeAsyncIterable = (items: unknown[]) => ({
    [Symbol.asyncIterator]: async function* () {
      for (const item of items) yield item;
    },
  });
  return {
    TableClient: {
      fromConnectionString: jest.fn((_conn: string, tableName: string) => ({
        createTable: async () => undefined,
        listEntities: ({ queryOptions }: { queryOptions: { filter: string } }) => {
          const pk = queryOptions.filter.match(/PartitionKey eq '([^']+)'/)?.[1] ?? '';
          return makeAsyncIterable(rows(tableName).filter((e) => e.partitionKey === pk));
        },
        upsertEntity: async (e: Row) => {
          const t = rows(tableName);
          const i = t.findIndex((r) => r.partitionKey === e.partitionKey && r.rowKey === e.rowKey);
          if (i >= 0) t.splice(i, 1, e); else t.push(e);
        },
        deleteEntity: async (pk: string, rk: string) => {
          const t = rows(tableName);
          const i = t.findIndex((r) => r.partitionKey === pk && r.rowKey === rk);
          if (i >= 0) t.splice(i, 1);
        },
      })),
    },
    odata: (strings: TemplateStringsArray, ...values: unknown[]) =>
      strings.reduce((acc, str, i) => `${acc}${str}${i < values.length ? `'${values[i]}'` : ''}`, ''),
  };
});
process.env.AZURE_STORAGE_CONNECTION_STRING = 'UseDevelopmentStorage=true';

import {
  addUserDenyEntry,
  agentKeyDenySubject,
  clearAgentKeyDenyEntries,
  filterDeniedPaths,
  filterDeniedSearchHits,
  isAgentKeyDenySubject,
  isPathDenied,
  listUserDenyEntries,
} from '../services/denyList.js';
import { markRequestAgentKey, runWithClientAddress } from '../services/clientAddress.js';

const TENANT = 'tenant-1';
const OWNER = 'owner-oid';
const KEY = 'abc123def456';

const request = { headers: new Map<string, string>() } as unknown as HttpRequest;

/** Run `fn` as a request that authenticated with agent key `keyId`. */
function asKey<T>(keyId: string, fn: () => Promise<T>): Promise<T> {
  return runWithClientAddress(request, () => {
    markRequestAgentKey(keyId);
    return fn();
  });
}

/** Run `fn` as a request made with the owner's own session. */
function asOwner<T>(fn: () => Promise<T>): Promise<T> {
  return runWithClientAddress(request, fn);
}

beforeEach(async () => {
  tables.clear();
  delete process.env.DEFAULT_MAIL_DENY_FOLDERS;
  delete process.env.DEFAULT_SHAREPOINT_DENY_PATHS;
  await addUserDenyEntry(OWNER, 'mail', 'Personal');
  await addUserDenyEntry(agentKeyDenySubject(KEY), 'mail', 'Board');
  await addUserDenyEntry(agentKeyDenySubject(KEY), 'calendar', 'Family');
  await addUserDenyEntry(agentKeyDenySubject(KEY), 'sharepoint', '/Shared Documents/Payroll');
});

describe('the key subject', () => {
  it('is distinct from any user id', () => {
    expect(agentKeyDenySubject(KEY)).toBe(`key:${KEY}`);
    expect(isAgentKeyDenySubject(agentKeyDenySubject(KEY))).toBe(true);
    expect(isAgentKeyDenySubject('8f1c2d3e-0000-4000-8000-000000000000')).toBe(false);
  });
});

describe('isPathDenied under an agent key', () => {
  it("applies the key's own entries on top of the owner's", async () => {
    await asKey(KEY, async () => {
      expect(await isPathDenied(TENANT, OWNER, 'mail', 'Board')).toBe(true);
      expect(await isPathDenied(TENANT, OWNER, 'mail', 'Board/2026')).toBe(true);
      expect(await isPathDenied(TENANT, OWNER, 'mail', 'Personal')).toBe(true);
      expect(await isPathDenied(TENANT, OWNER, 'mail', 'Inbox')).toBe(false);
      expect(await isPathDenied(TENANT, OWNER, 'calendar', 'Family')).toBe(true);
    });
  });

  it("leaves the owner's own session unaffected by the key's entries", async () => {
    await asOwner(async () => {
      expect(await isPathDenied(TENANT, OWNER, 'mail', 'Board')).toBe(false);
      expect(await isPathDenied(TENANT, OWNER, 'calendar', 'Family')).toBe(false);
      expect(await isPathDenied(TENANT, OWNER, 'mail', 'Personal')).toBe(true);
    });
  });

  it("does not apply one key's entries to another key", async () => {
    await asKey('fedcba987654', async () => {
      expect(await isPathDenied(TENANT, OWNER, 'mail', 'Board')).toBe(false);
    });
  });

  it('ignores key entries outside a request scope', async () => {
    expect(await isPathDenied(TENANT, OWNER, 'mail', 'Board')).toBe(false);
  });

  it('combines with delegated subjects', async () => {
    await addUserDenyEntry('mailbox-owner', 'mail', 'Legal');
    await asKey(KEY, async () => {
      expect(await isPathDenied(TENANT, [OWNER, 'mailbox-owner'], 'mail', 'Legal')).toBe(true);
      expect(await isPathDenied(TENANT, [OWNER, 'mailbox-owner'], 'mail', 'Board')).toBe(true);
    });
  });
});

describe('listing and search filters under an agent key', () => {
  it("filterDeniedPaths drops the key's folders by name", async () => {
    const folders = [
      { id: '1', name: 'Inbox' },
      { id: '2', name: 'Board' },
      { id: '3', name: 'Personal' },
    ];
    const asKeyResult = await asKey(KEY, () => filterDeniedPaths(TENANT, OWNER, 'mail', folders));
    expect(asKeyResult.map((f) => f.name)).toEqual(['Inbox']);
    const asOwnerResult = await asOwner(() => filterDeniedPaths(TENANT, OWNER, 'mail', folders));
    expect(asOwnerResult.map((f) => f.name)).toEqual(['Inbox', 'Board']);
  });

  it("filterDeniedPaths drops the key's calendars by name", async () => {
    const calendars = [{ id: 'c1', name: 'Calendar' }, { id: 'c2', name: 'Family' }];
    const result = await asKey(KEY, () => filterDeniedPaths(TENANT, OWNER, 'calendar', calendars));
    expect(result.map((c) => c.name)).toEqual(['Calendar']);
  });

  it("filterDeniedSearchHits drops hits under the key's folder", async () => {
    const hits = [
      { path: '/Shared Documents/Payroll/2026.xlsx', name: '2026.xlsx' },
      { path: '/Shared Documents/Projects/plan.docx', name: 'plan.docx' },
    ];
    const asKeyResult = await asKey(KEY, () => filterDeniedSearchHits(TENANT, OWNER, hits));
    expect(asKeyResult.map((h) => h.name)).toEqual(['plan.docx']);
    const asOwnerResult = await asOwner(() => filterDeniedSearchHits(TENANT, OWNER, hits));
    expect(asOwnerResult).toHaveLength(2);
  });
});

describe('clearAgentKeyDenyEntries', () => {
  it("removes the key's rows of every type and nothing else", async () => {
    await clearAgentKeyDenyEntries(KEY);
    const subject = agentKeyDenySubject(KEY);
    expect(await listUserDenyEntries(subject, 'mail')).toEqual([]);
    expect(await listUserDenyEntries(subject, 'calendar')).toEqual([]);
    expect(await listUserDenyEntries(subject, 'sharepoint')).toEqual([]);
    expect((await listUserDenyEntries(OWNER, 'mail')).map((e) => e.path)).toEqual(['Personal']);
  });
});
