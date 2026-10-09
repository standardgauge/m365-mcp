import { TableClient, TableServiceClient } from '@azure/data-tables';

const connectionString = process.env.AZURE_STORAGE_CONNECTION_STRING;
let auditTable: TableClient | null = null;
let auditTableInit = false;

async function ensureAuditTable(): Promise<void> {
  if (auditTableInit || !connectionString) return;
  const svc = TableServiceClient.fromConnectionString(connectionString);
  try { await svc.createTable('auditLog'); } catch { /* already exists */ }
  auditTable = TableClient.fromConnectionString(connectionString, 'auditLog');
  auditTableInit = true;
}

function getAuditTable(): TableClient {
  if (!auditTable) throw new Error('Audit table not initialized');
  return auditTable;
}

export interface AuditEntry {
  tenantId: string;
  userId: string;
  userEmail: string;
  deviceLabel?: string;
  operation: string;       // e.g. 'sharepoint.read_file', 'mail.search_mail'
  resource?: string;       // e.g. siteId/path, messageId
  result: 'allowed' | 'denied';
  reason?: string;         // denial reason if result === 'denied'
  source: 'http' | 'mcp';
  ip?: string;
}

// Year 9999-12-31T23:59:59.999Z in milliseconds — used to compute reverse timestamps.
// RowKey = pad(MAX_MS - Date.now()) + '_' + random so Table Storage ascending scan
// returns the newest entries first, giving correct newest-first semantics without
// scanning the full partition.
const MAX_TIMESTAMP_MS = 253402300799999;

export function reverseKeyFromMs(ms: number): string {
  return String(MAX_TIMESTAMP_MS - ms).padStart(15, '0');
}

function makeReverseRowKey(): string {
  const suffix = Math.random().toString(36).slice(2, 9);
  return `${reverseKeyFromMs(Date.now())}_${suffix}`;
}

/**
 * Fire-and-forget audit log write. Never throws — logs errors to console.error.
 */
export function logAccess(entry: AuditEntry): void {
  const timestamp = new Date().toISOString();
  const rowKey = makeReverseRowKey();

  ensureAuditTable().then(() => {
    return getAuditTable().upsertEntity({
      partitionKey: entry.tenantId,
      rowKey,
      tenantId: entry.tenantId,
      userId: entry.userId,
      userEmail: entry.userEmail,
      deviceLabel: entry.deviceLabel ?? null,
      operation: entry.operation,
      resource: entry.resource ?? null,
      result: entry.result,
      reason: entry.reason ?? null,
      source: entry.source,
      ip: entry.ip ?? null,
      timestamp,
    }, 'Replace');
  }).catch((err: unknown) => {
    console.error('[auditLog] Failed to write audit entry:', err instanceof Error ? err.message : err);
  });

  maybePurgeAuditLog();
}

// ── Retention ────────────────────────────────────────────────────────────────
//
// Rows older than AUDIT_LOG_RETENTION_DAYS are deleted by a purge that runs at
// most once per PURGE_INTERVAL_MS per process, started from logAccess. This is
// not a Functions timer trigger on purpose: timer triggers need
// AzureWebJobsStorage for their schedule monitor and singleton lock, and the
// Container App deliberately runs without it (see infra/container-app.bicep and
// the boot test in ci.yml). The trade-off is that the purge only runs while the
// server is serving calls, which is also the only time new rows are written.

export const DEFAULT_AUDIT_RETENTION_DAYS = 365;
const PURGE_INTERVAL_MS = 24 * 60 * 60 * 1000;
const DAY_MS = 24 * 60 * 60 * 1000;
// Bounds one run. A backlog larger than this is finished by later runs.
const PURGE_MAX_ROWS_PER_RUN = 50_000;
// Table Storage entity group transactions take at most 100 operations.
const PURGE_BATCH_SIZE = 100;

let lastPurgeStartedAt = 0;
let purgeInFlight = false;

/**
 * Retention in days from AUDIT_LOG_RETENTION_DAYS. `0` disables the purge and
 * keeps every row. Unset, empty, negative or non-integer values fall back to
 * DEFAULT_AUDIT_RETENTION_DAYS, so a typo never turns retention off.
 */
export function auditRetentionDays(): number {
  const raw = (process.env.AUDIT_LOG_RETENTION_DAYS ?? '').trim();
  if (raw === '') return DEFAULT_AUDIT_RETENTION_DAYS;
  if (!/^\d+$/.test(raw)) {
    console.error(`[auditLog] Ignoring invalid AUDIT_LOG_RETENTION_DAYS=${JSON.stringify(raw)}; using ${DEFAULT_AUDIT_RETENTION_DAYS}`);
    return DEFAULT_AUDIT_RETENTION_DAYS;
  }
  return Number(raw);
}

function isNotFound(err: unknown): boolean {
  return (err as { statusCode?: number } | null)?.statusCode === 404;
}

async function deleteBatch(table: TableClient, partitionKey: string, rowKeys: string[]): Promise<number> {
  try {
    await table.submitTransaction(rowKeys.map((rowKey) => ['delete', { partitionKey, rowKey }]));
    return rowKeys.length;
  } catch {
    // A transaction fails whole if any row is already gone (another replica
    // purging the same window). Fall back to single deletes that tolerate 404.
    let deleted = 0;
    for (const rowKey of rowKeys) {
      try {
        await table.deleteEntity(partitionKey, rowKey);
        deleted++;
      } catch (err) {
        if (!isNotFound(err)) throw err;
      }
    }
    return deleted;
  }
}

/**
 * Delete audit rows whose `Timestamp` is older than `retentionDays` before
 * `nowMs`, across every tenant partition. Returns the number of rows deleted.
 *
 * The filter names the service-side `Timestamp` system property with a
 * `datetime'…'` literal. The SDK serializes the `timestamp` field logAccess
 * writes as `Timestamp`, which the service owns, so no lowercase `timestamp`
 * column exists to filter on, and raw filter strings are not rewritten by the
 * SDK. Rows are written once and never updated, so `Timestamp` is the write
 * time. Filtering on it rather than the reverse RowKey also covers rows
 * written under any earlier key format.
 */
export async function purgeAuditLog(retentionDays: number, nowMs: number = Date.now()): Promise<number> {
  if (retentionDays <= 0) return 0;
  await ensureAuditTable();
  const table = getAuditTable();

  const cutoff = new Date(nowMs - retentionDays * DAY_MS).toISOString();
  const iterator = table.listEntities<{ partitionKey: string; rowKey: string }>({
    queryOptions: { filter: `Timestamp lt datetime'${cutoff}'`, select: ['partitionKey', 'rowKey'] },
  });

  const pending = new Map<string, string[]>();
  let deleted = 0;
  let seen = 0;

  for await (const entity of iterator) {
    const partitionKey = entity.partitionKey as string;
    const rowKeys = pending.get(partitionKey) ?? [];
    rowKeys.push(entity.rowKey as string);
    pending.set(partitionKey, rowKeys);
    if (rowKeys.length >= PURGE_BATCH_SIZE) {
      deleted += await deleteBatch(table, partitionKey, rowKeys);
      pending.delete(partitionKey);
    }
    if (++seen >= PURGE_MAX_ROWS_PER_RUN) break;
  }
  for (const [partitionKey, rowKeys] of pending) {
    deleted += await deleteBatch(table, partitionKey, rowKeys);
  }

  // Logged on every run, including zero, so Log Analytics shows the purge is alive.
  console.log(`[auditLog] Retention purge deleted ${deleted} row(s) older than ${cutoff} (retention ${retentionDays} days)`);
  return deleted;
}

/**
 * Start a purge if retention is enabled, none is running in this process, and
 * the last one started more than PURGE_INTERVAL_MS ago. Fire-and-forget; never
 * throws.
 */
export function maybePurgeAuditLog(nowMs: number = Date.now()): void {
  if (!connectionString || purgeInFlight) return;
  if (nowMs - lastPurgeStartedAt < PURGE_INTERVAL_MS) return;
  const retentionDays = auditRetentionDays();
  if (retentionDays === 0) return;

  lastPurgeStartedAt = nowMs;
  purgeInFlight = true;
  purgeAuditLog(retentionDays, nowMs).catch((err: unknown) => {
    console.error('[auditLog] Retention purge failed:', err instanceof Error ? err.message : err);
  }).finally(() => {
    purgeInFlight = false;
  });
}

/** Test hook: forget the last purge so the next maybePurgeAuditLog runs. */
export function resetAuditPurgeStateForTests(): void {
  lastPurgeStartedAt = 0;
  purgeInFlight = false;
}

/**
 * Query the audit log for a tenant. Returns entries sorted by timestamp descending
 * (newest first).
 *
 * Row keys are stored as reverse timestamps so ascending Table Storage scans yield
 * newest entries first. This means the limit is always applied to the newest
 * matching records, not the oldest prefix.
 *
 * Operation filtering is client-side (no OData support for substring match). When an
 * operation filter is specified, scanning stops after OPERATION_SCAN_CAP total rows to
 * bound cost; the returned entries are still the newest ones within that cap.
 */
export async function queryAuditLog(
  tenantId: string,
  opts?: {
    userEmail?: string;
    startDate?: string;
    endDate?: string;
    operation?: string;
    result?: 'allowed' | 'denied';
    limit?: number;
  }
): Promise<(AuditEntry & { timestamp: string })[]> {
  await ensureAuditTable();

  const limit = Math.min(opts?.limit ?? 200, 1000);
  // Cap total rows scanned when the operation filter is active to avoid full-partition
  // scans. Entries returned are always the newest within the cap.
  const OPERATION_SCAN_CAP = Math.min(10 * limit, 10_000);

  // Escape single quotes in OData string literals to prevent filter injection
  const esc = (s: string) => s.replace(/'/g, "''");

  // Build OData filter — exact-match fields only; date range uses reverse RowKey
  const filters: string[] = [`PartitionKey eq '${esc(tenantId)}'`];
  if (opts?.userEmail) {
    filters.push(`userEmail eq '${esc(opts.userEmail)}'`);
  }
  if (opts?.result) {
    filters.push(`result eq '${esc(opts.result)}'`);
  }
  if (opts?.endDate) {
    // Entries before endDate have larger reverse keys → skip entries newer than endDate
    // by starting the scan at the reverse of endDate.
    const endMs = new Date(opts.endDate.replace(/[^0-9TZ:.+-]/g, '')).getTime();
    if (!isNaN(endMs)) {
      filters.push(`RowKey ge '${reverseKeyFromMs(endMs)}'`);
    }
  }

  const filterStr = filters.join(' and ');

  const entries: (AuditEntry & { timestamp: string })[] = [];
  let scanned = 0;

  const iterator = getAuditTable().listEntities({
    queryOptions: { filter: filterStr },
  });

  for await (const entity of iterator) {
    scanned++;

    const ts = entity.timestamp as string;

    // Oldest-bound early exit: once we've passed startDate (scanning newest→oldest),
    // all subsequent entries are older — stop now.
    if (opts?.startDate && ts < opts.startDate) break;

    // endDate boundary (inclusive-exclusive, client-side to catch exact-millisecond edge)
    if (opts?.endDate && ts >= opts.endDate) continue;

    // Operation filter — client-side only; apply scan cap to bound cost
    if (opts?.operation) {
      if (!String(entity.operation ?? '').includes(opts.operation)) {
        if (scanned >= OPERATION_SCAN_CAP) break;
        continue;
      }
    }

    entries.push({
      tenantId: entity.partitionKey as string,
      userId: entity.userId as string,
      userEmail: entity.userEmail as string,
      deviceLabel: entity.deviceLabel as string | undefined || undefined,
      operation: entity.operation as string,
      resource: entity.resource as string | undefined || undefined,
      result: entity.result as 'allowed' | 'denied',
      reason: entity.reason as string | undefined || undefined,
      source: entity.source as 'http' | 'mcp',
      ip: entity.ip as string | undefined || undefined,
      timestamp: ts,
    });

    if (entries.length >= limit) break;
  }

  // Scan is already newest-first; no reversal needed.
  return entries;
}
