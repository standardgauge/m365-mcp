/**
 * Tests for the auditLog retention purge: the AUDIT_LOG_RETENTION_DAYS parser,
 * purgeAuditLog's filter and batching, and the once-a-day trigger in logAccess.
 */

import { jest } from '@jest/globals';

// ── Mock Azure Table Storage ─────────────────────────────────────────────────

type ListEntitiesOpts = { queryOptions?: { filter?: string; select?: string[] } };
type Action = [string, { partitionKey: string; rowKey: string }];
const mockListEntities = jest.fn<(opts: ListEntitiesOpts) => AsyncIterable<Record<string, unknown>>>();
const mockSubmitTransaction = jest.fn<(actions: Action[]) => Promise<void>>();
const mockDeleteEntity = jest.fn<(pk: string, rk: string) => Promise<void>>();
const mockUpsertEntity = jest.fn<(entity: unknown, mode: string) => Promise<void>>();

jest.mock('@azure/data-tables', () => ({
  TableClient: {
    fromConnectionString: () => ({
      listEntities: (...args: unknown[]) => mockListEntities(args[0] as ListEntitiesOpts),
      submitTransaction: (...args: unknown[]) => mockSubmitTransaction(args[0] as Action[]),
      deleteEntity: (...args: unknown[]) => mockDeleteEntity(args[0] as string, args[1] as string),
      upsertEntity: (...args: unknown[]) => mockUpsertEntity(args[0], args[1] as string),
    }),
  },
  TableServiceClient: {
    fromConnectionString: () => ({
      createTable: () => Promise.resolve(),
    }),
  },
}));

process.env.AZURE_STORAGE_CONNECTION_STRING =
  'DefaultEndpointsProtocol=https;AccountName=test;AccountKey=dGVzdA==;EndpointSuffix=core.windows.net';

import {
  auditRetentionDays,
  DEFAULT_AUDIT_RETENTION_DAYS,
  logAccess,
  maybePurgeAuditLog,
  purgeAuditLog,
  resetAuditPurgeStateForTests,
} from '../services/auditLog.js';

const DAY_MS = 24 * 60 * 60 * 1000;
const NOW = Date.parse('2026-10-09T12:00:00.000Z');

function setRows(rows: { partitionKey: string; rowKey: string }[]) {
  mockListEntities.mockReturnValue({
    [Symbol.asyncIterator]: async function* () {
      for (const r of rows) yield r;
    },
  } as AsyncIterable<Record<string, unknown>>);
}

function rows(partitionKey: string, count: number) {
  return Array.from({ length: count }, (_, i) => ({ partitionKey, rowKey: `rk${i}` }));
}

const flush = () => new Promise((resolve) => setImmediate(resolve));

let logSpy: ReturnType<typeof jest.spyOn>;
let errSpy: ReturnType<typeof jest.spyOn>;

beforeEach(() => {
  jest.clearAllMocks();
  delete process.env.AUDIT_LOG_RETENTION_DAYS;
  resetAuditPurgeStateForTests();
  mockSubmitTransaction.mockResolvedValue(undefined);
  mockDeleteEntity.mockResolvedValue(undefined);
  mockUpsertEntity.mockResolvedValue(undefined);
  setRows([]);
  logSpy = jest.spyOn(console, 'log').mockImplementation(() => undefined);
  errSpy = jest.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  logSpy.mockRestore();
  errSpy.mockRestore();
});

describe('auditRetentionDays', () => {
  it('defaults when unset or empty', () => {
    expect(auditRetentionDays()).toBe(DEFAULT_AUDIT_RETENTION_DAYS);
    process.env.AUDIT_LOG_RETENTION_DAYS = '  ';
    expect(auditRetentionDays()).toBe(DEFAULT_AUDIT_RETENTION_DAYS);
  });

  it('reads a positive integer', () => {
    process.env.AUDIT_LOG_RETENTION_DAYS = '730';
    expect(auditRetentionDays()).toBe(730);
  });

  it('treats 0 as keep forever', () => {
    process.env.AUDIT_LOG_RETENTION_DAYS = '0';
    expect(auditRetentionDays()).toBe(0);
  });

  it.each(['-30', '30.5', 'ninety', '1e3'])('falls back to the default for %s rather than disabling', (raw) => {
    process.env.AUDIT_LOG_RETENTION_DAYS = raw;
    expect(auditRetentionDays()).toBe(DEFAULT_AUDIT_RETENTION_DAYS);
    expect(errSpy).toHaveBeenCalled();
  });
});

describe('purgeAuditLog', () => {
  it('filters on the service-side Timestamp cutoff across all partitions', async () => {
    await purgeAuditLog(90, NOW);
    const opts = mockListEntities.mock.calls[0][0];
    const cutoff = new Date(NOW - 90 * DAY_MS).toISOString();
    expect(opts.queryOptions?.filter).toBe(`Timestamp lt datetime'${cutoff}'`);
    expect(opts.queryOptions?.filter).not.toContain('PartitionKey');
  });

  it('deletes in per-partition transactions of at most 100', async () => {
    setRows([...rows('tenantA', 250), ...rows('tenantB', 3)]);
    const deleted = await purgeAuditLog(90, NOW);

    expect(deleted).toBe(253);
    const batches = mockSubmitTransaction.mock.calls.map((c) => c[0]);
    expect(batches.map((b) => b.length).sort((a, b) => a - b)).toEqual([3, 50, 100, 100]);
    for (const batch of batches) {
      expect(new Set(batch.map(([, e]) => e.partitionKey)).size).toBe(1);
      expect(batch.every(([op]) => op === 'delete')).toBe(true);
    }
  });

  it('falls back to single deletes and tolerates rows already gone', async () => {
    setRows(rows('tenantA', 3));
    mockSubmitTransaction.mockRejectedValueOnce(Object.assign(new Error('not found'), { statusCode: 404 }));
    mockDeleteEntity
      .mockResolvedValueOnce(undefined)
      .mockRejectedValueOnce(Object.assign(new Error('not found'), { statusCode: 404 }))
      .mockResolvedValueOnce(undefined);

    await expect(purgeAuditLog(90, NOW)).resolves.toBe(2);
    expect(mockDeleteEntity).toHaveBeenCalledTimes(3);
  });

  it('surfaces a non-404 failure in the fallback', async () => {
    setRows(rows('tenantA', 1));
    mockSubmitTransaction.mockRejectedValueOnce(new Error('boom'));
    mockDeleteEntity.mockRejectedValueOnce(Object.assign(new Error('forbidden'), { statusCode: 403 }));
    await expect(purgeAuditLog(90, NOW)).rejects.toThrow('forbidden');
  });

  it('logs a line on every run, including when nothing is due', async () => {
    await purgeAuditLog(90, NOW);
    expect(logSpy).toHaveBeenCalledWith(expect.stringContaining('Retention purge deleted 0 row(s)'));
  });

  it('does nothing when retention is disabled', async () => {
    await expect(purgeAuditLog(0, NOW)).resolves.toBe(0);
    expect(mockListEntities).not.toHaveBeenCalled();
  });
});

describe('maybePurgeAuditLog', () => {
  it('runs at most once a day per process', async () => {
    maybePurgeAuditLog(NOW);
    await flush();
    maybePurgeAuditLog(NOW + DAY_MS - 1);
    await flush();
    expect(mockListEntities).toHaveBeenCalledTimes(1);

    maybePurgeAuditLog(NOW + DAY_MS);
    await flush();
    expect(mockListEntities).toHaveBeenCalledTimes(2);
  });

  it('uses the configured retention', async () => {
    process.env.AUDIT_LOG_RETENTION_DAYS = '30';
    maybePurgeAuditLog(NOW);
    await flush();
    const cutoff = new Date(NOW - 30 * DAY_MS).toISOString();
    expect(mockListEntities.mock.calls[0][0].queryOptions?.filter).toBe(`Timestamp lt datetime'${cutoff}'`);
  });

  it('is skipped entirely when AUDIT_LOG_RETENTION_DAYS=0', async () => {
    process.env.AUDIT_LOG_RETENTION_DAYS = '0';
    maybePurgeAuditLog(NOW);
    await flush();
    expect(mockListEntities).not.toHaveBeenCalled();
  });

  it('swallows a purge failure', async () => {
    mockListEntities.mockImplementation(() => { throw new Error('table unavailable'); });
    expect(() => maybePurgeAuditLog(NOW)).not.toThrow();
    await flush();
    expect(errSpy).toHaveBeenCalledWith('[auditLog] Retention purge failed:', 'table unavailable');
  });

  it('is started by logAccess', async () => {
    logAccess({
      tenantId: 'tenantA', userId: 'u1', userEmail: 'user@example.com',
      operation: 'mail.search_mail', result: 'allowed', source: 'mcp',
    });
    await flush();
    expect(mockUpsertEntity).toHaveBeenCalledTimes(1);
    expect(mockListEntities).toHaveBeenCalledTimes(1);
  });
});
