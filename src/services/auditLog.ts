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
