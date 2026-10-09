/**
 * Audit events → the instance's Log Analytics workspace, through the Azure
 * Monitor Logs Ingestion API and the data collection rule in infra/.
 *
 * This is the authoritative audit record. The auditLog table in storage stays
 * as the admin screen's backing store; every logAccess event goes to both.
 *
 * Configured by three env vars, all set by the Bicep:
 *   AUDIT_LOGS_INGESTION_ENDPOINT  the DCR's logs ingestion endpoint
 *   AUDIT_DCR_IMMUTABLE_ID         the DCR's immutable id (dcr-…)
 *   AUDIT_DCR_STREAM_NAME          optional, defaults to Custom-M365McpAudit
 * With the endpoint or rule id unset the sink is off and logAccess writes the
 * table only (local dev, or an instance whose infra predates the rule).
 *
 * Failure handling is deliberately minimal: a failed upload is logged to
 * console.error and dropped. The server does not fail closed and raises no
 * alert of its own; watching for gaps in the table is the client's security
 * tooling's job (see docs/operations-runbook.md, Audit trail in Log Analytics).
 */
import { LogsIngestionClient } from '@azure/monitor-ingestion';
import { ContainerAppManagedIdentityCredential } from './managedIdentity.js';
import type { AuditEntry } from './auditLog.js';

export { ContainerAppManagedIdentityCredential };

export const DEFAULT_AUDIT_STREAM = 'Custom-M365McpAudit';

/** Upload as soon as this many events are queued. */
export const MAX_BATCH = 100;
/** Otherwise upload this long after the first event in a batch was queued. */
export const FLUSH_INTERVAL_MS = 1000;

/**
 * One row of the M365McpAudit_CL table. Column names and types must match the
 * stream declaration and table schema in infra/audit-ingestion.bicep.
 *
 * TenantId is a reserved column in every Log Analytics table (it holds the
 * workspace id), so the Entra tenant goes in EntraTenantId.
 */
export interface AuditLogAnalyticsRecord {
  TimeGenerated: string;
  EventId: string;
  EntraTenantId: string;
  UserId: string;
  UserEmail: string;
  DeviceLabel?: string;
  Operation: string;
  TargetResource?: string;
  Result: 'allowed' | 'denied';
  Reason?: string;
  Source: 'http' | 'mcp';
  ClientIp?: string;
}

/**
 * Map an audit entry to its Log Analytics row. `eventId` is the auditLog
 * table's RowKey for the same event, so the two records can be joined.
 * Absent optional fields are omitted rather than sent as null.
 */
export function toLogAnalyticsRecord(
  entry: AuditEntry,
  timestamp: string,
  eventId: string,
): AuditLogAnalyticsRecord {
  const record: AuditLogAnalyticsRecord = {
    TimeGenerated: timestamp,
    EventId: eventId,
    EntraTenantId: entry.tenantId,
    UserId: entry.userId,
    UserEmail: entry.userEmail,
    Operation: entry.operation,
    Result: entry.result,
    Source: entry.source,
  };
  if (entry.deviceLabel) record.DeviceLabel = entry.deviceLabel;
  if (entry.resource) record.TargetResource = entry.resource;
  if (entry.reason) record.Reason = entry.reason;
  if (entry.ip) record.ClientIp = entry.ip;
  return record;
}

interface SinkConfig {
  endpoint: string;
  ruleId: string;
  stream: string;
}

function readConfig(): SinkConfig | null {
  const endpoint = process.env.AUDIT_LOGS_INGESTION_ENDPOINT;
  const ruleId = process.env.AUDIT_DCR_IMMUTABLE_ID;
  if (!endpoint || !ruleId) return null;
  return { endpoint, ruleId, stream: process.env.AUDIT_DCR_STREAM_NAME || DEFAULT_AUDIT_STREAM };
}

type Uploader = Pick<LogsIngestionClient, 'upload'>;

let config: SinkConfig | null | undefined;
let client: Uploader | null = null;
let queue: AuditLogAnalyticsRecord[] = [];
let timer: ReturnType<typeof setTimeout> | null = null;

function getConfig(): SinkConfig | null {
  if (config === undefined) config = readConfig();
  return config;
}

function getClient(cfg: SinkConfig): Uploader {
  if (!client) {
    client = new LogsIngestionClient(cfg.endpoint, new ContainerAppManagedIdentityCredential());
  }
  return client;
}

/** True when the ingestion endpoint and rule id are configured. */
export function isAuditLogAnalyticsEnabled(): boolean {
  return getConfig() !== null;
}

/**
 * Queue one record for upload. Never throws and never blocks the caller; a
 * no-op when the sink is not configured.
 */
export function sendToLogAnalytics(record: AuditLogAnalyticsRecord): void {
  if (!getConfig()) return;
  queue.push(record);
  if (queue.length >= MAX_BATCH) {
    void flushAuditLogAnalytics();
  } else if (!timer) {
    timer = setTimeout(() => { void flushAuditLogAnalytics(); }, FLUSH_INTERVAL_MS);
    timer.unref?.();
  }
}

/** Upload everything queued so far. Resolves once the upload settles; never rejects. */
export async function flushAuditLogAnalytics(): Promise<void> {
  if (timer) {
    clearTimeout(timer);
    timer = null;
  }
  const cfg = getConfig();
  if (!cfg || queue.length === 0) return;
  const batch = queue;
  queue = [];
  try {
    await getClient(cfg).upload(cfg.ruleId, cfg.stream, batch as unknown as Record<string, unknown>[]);
  } catch (err: unknown) {
    console.error(
      `[auditLog] Failed to send ${batch.length} audit event(s) to Log Analytics:`,
      err instanceof Error ? err.message : err,
    );
  }
}

/** Test seam: reset module state and optionally substitute the upload client. */
export function __resetAuditLogAnalyticsForTests(uploader: Uploader | null = null): void {
  if (timer) clearTimeout(timer);
  timer = null;
  queue = [];
  config = undefined;
  client = uploader;
}
