/**
 * Tests for auditLog.ts — focusing on queryAuditLog newest-first semantics.
 *
 * Key invariant: regardless of how many rows exist in the partition, queryAuditLog
 * must always return the NEWEST matching entries up to `limit`, never the oldest.
 */

import { jest } from '@jest/globals';

// ── Mock Azure Table Storage ─────────────────────────────────────────────────

type ListEntitiesOpts = { queryOptions?: { filter?: string } };
const mockListEntities = jest.fn<(opts: ListEntitiesOpts) => AsyncIterable<Record<string, unknown>>>();
const mockUpsertEntity = jest.fn<(entity: unknown, mode: string) => Promise<void>>();
const mockCreateTable = jest.fn<() => Promise<void>>();

jest.mock('@azure/data-tables', () => ({
  TableClient: {
    fromConnectionString: () => ({
      listEntities: (...args: unknown[]) => mockListEntities(args[0] as ListEntitiesOpts),
      upsertEntity: (...args: unknown[]) => mockUpsertEntity(args[0], args[1] as string),
    }),
  },
  TableServiceClient: {
    fromConnectionString: () => ({
      createTable: () => mockCreateTable(),
    }),
  },
}));

process.env.AZURE_STORAGE_CONNECTION_STRING =
  'DefaultEndpointsProtocol=https;AccountName=test;AccountKey=dGVzdA==;EndpointSuffix=core.windows.net';

import { queryAuditLog, reverseKeyFromMs } from '../services/auditLog.js';

// ── Fixtures ─────────────────────────────────────────────────────────────────

const MAX_TIMESTAMP_MS = 253402300799999;

function makeEntity(
  ms: number,
  overrides: Partial<Record<string, unknown>> = {}
): Record<string, unknown> {
  return {
    partitionKey: 'tenant1',
    rowKey: `${String(MAX_TIMESTAMP_MS - ms).padStart(15, '0')}_abc1234`,
    tenantId: 'tenant1',
    userId: 'user1',
    userEmail: 'user@example.com',
    deviceLabel: null,
    operation: 'sharepoint.read_file',
    resource: '/sites/it-hub/budget.xlsx',
    result: 'allowed',
    reason: null,
    source: 'http',
    ip: null,
    timestamp: new Date(ms).toISOString(),
    ...overrides,
  };
}

/**
 * Build N entities in newest-first order (ascending reverse key = newest first),
 * spaced 1 second apart starting from baseMs.
 */
function buildEntities(
  count: number,
  baseMs: number,
  overrides: Partial<Record<string, unknown>> = {}
): Record<string, unknown>[] {
  return Array.from({ length: count }, (_, i) => makeEntity(baseMs - i * 1000, overrides));
}

function setEntities(entities: Record<string, unknown>[]) {
  mockListEntities.mockReturnValue({
    [Symbol.asyncIterator]: async function* () {
      for (const e of entities) yield e;
    },
  } as AsyncIterable<Record<string, unknown>>);
}

// ── Setup ─────────────────────────────────────────────────────────────────────

beforeEach(() => {
  jest.clearAllMocks();
  mockCreateTable.mockResolvedValue(undefined);
  mockUpsertEntity.mockResolvedValue(undefined);
});

// ── reverseKeyFromMs ──────────────────────────────────────────────────────────

describe('reverseKeyFromMs', () => {
  it('produces a 15-character zero-padded string', () => {
    const key = reverseKeyFromMs(1_750_000_000_000);
    expect(key).toHaveLength(15);
    expect(/^\d+$/.test(key)).toBe(true);
  });

  it('newer timestamps produce numerically smaller reverse keys', () => {
    const now = 1_750_000_000_000;
    const older = now - 60_000;
    expect(Number(reverseKeyFromMs(now))).toBeLessThan(Number(reverseKeyFromMs(older)));
  });

  it('lexicographic ascending order of reverse keys = descending chronological order', () => {
    const times = [1_750_000_000_000, 1_750_001_000_000, 1_749_999_000_000];
    const keys = times.map(reverseKeyFromMs);
    const sorted = [...keys].sort(); // ascending lexicographic
    // sorted[0] = smallest key = newest timestamp
    expect(sorted[0]).toBe(reverseKeyFromMs(1_750_001_000_000)); // newest
    expect(sorted[2]).toBe(reverseKeyFromMs(1_749_999_000_000)); // oldest
  });
});

// ── queryAuditLog — newest-first semantics ────────────────────────────────────

describe('queryAuditLog — newest-first semantics', () => {
  const BASE_MS = 1_750_000_000_000;

  it('returns all entries when total <= limit, newest first', async () => {
    setEntities(buildEntities(5, BASE_MS));

    const result = await queryAuditLog('tenant1', { limit: 10 });

    expect(result).toHaveLength(5);
    expect(result[0].timestamp).toBe(new Date(BASE_MS).toISOString());
    expect(result[4].timestamp).toBe(new Date(BASE_MS - 4000).toISOString());
  });

  it('returns exactly limit NEWEST entries when total > limit (regression: oldest-first bug)', async () => {
    // 300 entries — well above the default limit of 200
    setEntities(buildEntities(300, BASE_MS));

    const result = await queryAuditLog('tenant1', { limit: 10 });

    expect(result).toHaveLength(10);
    // Must be the 10 NEWEST
    for (let i = 0; i < 10; i++) {
      expect(result[i].timestamp).toBe(new Date(BASE_MS - i * 1000).toISOString());
    }
    // Must NOT include old entries
    const oldestMs = BASE_MS - 299 * 1000;
    expect(result.some(e => e.timestamp === new Date(oldestMs).toISOString())).toBe(false);
  });

  it('handles thousands of entries — always returns newest (regression: scan-cap still returned oldest)', async () => {
    // 1,000 entries — matches real high-volume production scenario
    setEntities(buildEntities(1_000, BASE_MS));

    const result = await queryAuditLog('tenant1', { limit: 5 });

    expect(result).toHaveLength(5);
    for (let i = 0; i < 5; i++) {
      expect(result[i].timestamp).toBe(new Date(BASE_MS - i * 1000).toISOString());
    }
  });

  it('applies startDate early-exit: stops collecting once entries are older than startDate', async () => {
    setEntities(buildEntities(50, BASE_MS));

    // Only the 10 newest entries are within range
    const startDate = new Date(BASE_MS - 9000).toISOString();
    const result = await queryAuditLog('tenant1', { startDate, limit: 200 });

    expect(result).toHaveLength(10);
    expect(result[0].timestamp).toBe(new Date(BASE_MS).toISOString());
    expect(result[9].timestamp).toBe(new Date(BASE_MS - 9000).toISOString());
  });

  it('applies endDate filter: excludes entries at or after endDate', async () => {
    setEntities(buildEntities(20, BASE_MS));

    // Exclude the 5 newest
    const endDate = new Date(BASE_MS - 4999).toISOString(); // exclusive upper bound
    const result = await queryAuditLog('tenant1', { endDate, limit: 200 });

    for (const e of result) {
      expect(e.timestamp < endDate).toBe(true);
    }
    // Newest entry in result should be baseMs - 5000
    expect(result[0].timestamp).toBe(new Date(BASE_MS - 5000).toISOString());
  });

  it('filters by operation client-side and returns newest matching entries', async () => {
    // Mix of operations — every 10th entry is mail.search_mail; rest are sharepoint
    const entities = Array.from({ length: 100 }, (_, i) =>
      makeEntity(BASE_MS - i * 1000, {
        operation: i % 10 === 0 ? 'mail.search_mail' : 'sharepoint.read_file',
      })
    );
    setEntities(entities);

    const result = await queryAuditLog('tenant1', { operation: 'mail.search', limit: 3 });

    expect(result).toHaveLength(3);
    for (const e of result) {
      expect(e.operation).toContain('mail.search');
    }
    // Newest mail.search entry is at i=0 (baseMs)
    expect(result[0].timestamp).toBe(new Date(BASE_MS).toISOString());
  });

  it('operation filter with more entries than scan cap returns newest within cap', async () => {
    // limit=5 → OPERATION_SCAN_CAP = min(50, 10000) = 50
    // Provide 200 entries where only every 5th matches; matches beyond cap are older
    const entities = Array.from({ length: 200 }, (_, i) =>
      makeEntity(BASE_MS - i * 1000, {
        operation: i % 5 === 0 ? 'mail.search_mail' : 'sharepoint.read_file',
      })
    );
    setEntities(entities);

    const result = await queryAuditLog('tenant1', { operation: 'mail.search', limit: 5 });

    // Should return entries from the newest cap window, all matching
    expect(result.length).toBeGreaterThan(0);
    expect(result.length).toBeLessThanOrEqual(5);
    for (const e of result) {
      expect(e.operation).toContain('mail.search');
    }
    // All returned entries must be newer than entries beyond the scan cap
    const capBoundaryMs = BASE_MS - 49 * 1000; // 50th entry (0-indexed)
    for (const e of result) {
      expect(e.timestamp >= new Date(capBoundaryMs).toISOString()).toBe(true);
    }
  });

  it('includes endDate RowKey ge filter in OData to skip newer rows', async () => {
    setEntities([]);

    const endDate = new Date(BASE_MS - 5000).toISOString();
    await queryAuditLog('tenant1', { endDate, limit: 10 });

    const callArg = mockListEntities.mock.calls[0]?.[0] as ListEntitiesOpts | undefined;
    const filter = callArg?.queryOptions?.filter ?? '';
    expect(filter).toMatch(/RowKey ge '/);
  });

  it('includes userEmail and result in OData filter', async () => {
    setEntities([]);

    await queryAuditLog('tenant1', { userEmail: 'alice@example.com', result: 'denied', limit: 10 });

    const callArg = mockListEntities.mock.calls[0]?.[0] as ListEntitiesOpts | undefined;
    const filter = callArg?.queryOptions?.filter ?? '';
    expect(filter).toContain("userEmail eq 'alice@example.com'");
    expect(filter).toContain("result eq 'denied'");
  });

  it('escapes single quotes in OData filter values', async () => {
    setEntities([]);

    await queryAuditLog("tenant'with'quotes", { userEmail: "o'malley@example.com", limit: 10 });

    const callArg = mockListEntities.mock.calls[0]?.[0] as ListEntitiesOpts | undefined;
    const filter = callArg?.queryOptions?.filter ?? '';
    expect(filter).toContain("tenant''with''quotes");
    expect(filter).toContain("o''malley@example.com");
  });
});
