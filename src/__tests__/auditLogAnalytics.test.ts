/**
 * Tests for the Log Analytics audit sink (auditLogAnalytics.ts) and its wiring
 * into logAccess: the row shape, batching, the no-op when unconfigured, and the
 * no-action-on-failure rule.
 */

import { jest } from '@jest/globals';

// ── Mock Azure Table Storage (logAccess still writes the auditLog table) ─────

const mockUpsertEntity = jest.fn<(entity: Record<string, unknown>, mode: string) => Promise<void>>();

jest.mock('@azure/data-tables', () => ({
  TableClient: {
    fromConnectionString: () => ({
      upsertEntity: (...args: unknown[]) => mockUpsertEntity(args[0] as Record<string, unknown>, args[1] as string),
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

import { logAccess, type AuditEntry } from '../services/auditLog.js';
import {
  ContainerAppManagedIdentityCredential,
  DEFAULT_AUDIT_STREAM,
  FLUSH_INTERVAL_MS,
  MAX_BATCH,
  __resetAuditLogAnalyticsForTests,
  flushAuditLogAnalytics,
  isAuditLogAnalyticsEnabled,
  sendToLogAnalytics,
  toLogAnalyticsRecord,
} from '../services/auditLogAnalytics.js';

// ── Fixtures ─────────────────────────────────────────────────────────────────

const ENDPOINT = 'https://audit-dcr-abcd.eastus-1.ingest.monitor.azure.com';
const RULE_ID = 'dcr-00000000000000000000000000000000';

const fullEntry: AuditEntry = {
  tenantId: '11111111-2222-3333-4444-555555555555',
  userId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
  userEmail: 'adele@fabrikam.com',
  deviceLabel: 'laptop',
  operation: 'mail.search_mail',
  resource: 'folder:Inbox',
  result: 'denied',
  reason: 'folder on deny list',
  source: 'mcp',
  ip: '203.0.113.7',
};

const minimalEntry: AuditEntry = {
  tenantId: '11111111-2222-3333-4444-555555555555',
  userId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee',
  userEmail: 'adele@fabrikam.com',
  operation: 'sharepoint.read_file',
  result: 'allowed',
  source: 'http',
};

type UploadArgs = [string, string, Record<string, unknown>[]];
const mockUpload = jest.fn<(...args: UploadArgs) => Promise<void>>();
const uploader = { upload: (...args: unknown[]) => mockUpload(...(args as UploadArgs)) };

function configure(): void {
  process.env.AUDIT_LOGS_INGESTION_ENDPOINT = ENDPOINT;
  process.env.AUDIT_DCR_IMMUTABLE_ID = RULE_ID;
  delete process.env.AUDIT_DCR_STREAM_NAME;
  __resetAuditLogAnalyticsForTests(uploader);
}

function unconfigure(): void {
  delete process.env.AUDIT_LOGS_INGESTION_ENDPOINT;
  delete process.env.AUDIT_DCR_IMMUTABLE_ID;
  delete process.env.AUDIT_DCR_STREAM_NAME;
  __resetAuditLogAnalyticsForTests(uploader);
}

beforeEach(() => {
  jest.clearAllMocks();
  mockUpload.mockResolvedValue(undefined);
  mockUpsertEntity.mockResolvedValue(undefined);
});

afterEach(() => {
  jest.useRealTimers();
  unconfigure();
});

// ── Event shape ──────────────────────────────────────────────────────────────

describe('toLogAnalyticsRecord', () => {
  it('maps every field of a full entry to its table column', () => {
    const record = toLogAnalyticsRecord(fullEntry, '2026-01-02T03:04:05.678Z', '0001_abc');
    expect(record).toEqual({
      TimeGenerated: '2026-01-02T03:04:05.678Z',
      EventId: '0001_abc',
      EntraTenantId: fullEntry.tenantId,
      UserId: fullEntry.userId,
      UserEmail: 'adele@fabrikam.com',
      DeviceLabel: 'laptop',
      Operation: 'mail.search_mail',
      TargetResource: 'folder:Inbox',
      Result: 'denied',
      Reason: 'folder on deny list',
      Source: 'mcp',
      ClientIp: '203.0.113.7',
    });
  });

  it('omits absent optional fields instead of sending null', () => {
    const record = toLogAnalyticsRecord(minimalEntry, '2026-01-02T03:04:05.678Z', '0001_abc');
    expect(Object.keys(record).sort()).toEqual(
      ['EntraTenantId', 'EventId', 'Operation', 'Result', 'Source', 'TimeGenerated', 'UserEmail', 'UserId'],
    );
    expect(Object.values(record)).not.toContain(null);
  });

  it('never uses the reserved TenantId column', () => {
    const record = toLogAnalyticsRecord(fullEntry, '2026-01-02T03:04:05.678Z', 'x');
    expect(record).not.toHaveProperty('TenantId');
  });
});

// ── logAccess wiring ─────────────────────────────────────────────────────────

describe('logAccess → Log Analytics', () => {
  it('sends the event with the same timestamp and EventId as the table row', async () => {
    configure();
    logAccess(fullEntry);
    await flushAuditLogAnalytics();
    await new Promise(r => setImmediate(r)); // let the table write settle

    expect(mockUpload).toHaveBeenCalledTimes(1);
    const [ruleId, stream, rows] = mockUpload.mock.calls[0];
    expect(ruleId).toBe(RULE_ID);
    expect(stream).toBe(DEFAULT_AUDIT_STREAM);
    expect(rows).toHaveLength(1);

    expect(mockUpsertEntity).toHaveBeenCalledTimes(1);
    const tableRow = mockUpsertEntity.mock.calls[0][0];
    expect(rows[0].EventId).toBe(tableRow.rowKey);
    expect(rows[0].TimeGenerated).toBe(tableRow.timestamp);
    expect(rows[0].Operation).toBe('mail.search_mail');
    expect(rows[0].Result).toBe('denied');
  });

  it('honours AUDIT_DCR_STREAM_NAME', async () => {
    configure();
    process.env.AUDIT_DCR_STREAM_NAME = 'Custom-Other';
    __resetAuditLogAnalyticsForTests(uploader);
    logAccess(minimalEntry);
    await flushAuditLogAnalytics();
    expect(mockUpload.mock.calls[0][1]).toBe('Custom-Other');
  });

  it('is a no-op when the endpoint or rule id is unset, and the table write still happens', async () => {
    unconfigure();
    expect(isAuditLogAnalyticsEnabled()).toBe(false);
    logAccess(fullEntry);
    await flushAuditLogAnalytics();
    await new Promise(r => setImmediate(r));
    expect(mockUpload).not.toHaveBeenCalled();
    expect(mockUpsertEntity).toHaveBeenCalledTimes(1);

    process.env.AUDIT_LOGS_INGESTION_ENDPOINT = ENDPOINT; // rule id still missing
    __resetAuditLogAnalyticsForTests(uploader);
    expect(isAuditLogAnalyticsEnabled()).toBe(false);
  });

  it('takes no action on a failed upload beyond logging it', async () => {
    configure();
    mockUpload.mockRejectedValueOnce(new Error('403 Forbidden'));
    const errSpy = jest.spyOn(console, 'error').mockImplementation(() => {});

    expect(() => logAccess(fullEntry)).not.toThrow();
    await expect(flushAuditLogAnalytics()).resolves.toBeUndefined();
    await new Promise(r => setImmediate(r));

    expect(errSpy).toHaveBeenCalledWith(
      expect.stringContaining('Failed to send 1 audit event(s) to Log Analytics'),
      '403 Forbidden',
    );
    // The table write is independent of the Log Analytics outcome.
    expect(mockUpsertEntity).toHaveBeenCalledTimes(1);
    // The failed batch is dropped, not retried on the next flush.
    await flushAuditLogAnalytics();
    expect(mockUpload).toHaveBeenCalledTimes(1);
    errSpy.mockRestore();
  });
});

// ── Batching ─────────────────────────────────────────────────────────────────

describe('batching', () => {
  it('uploads queued events together once the flush interval passes', async () => {
    jest.useFakeTimers();
    configure();
    const r = toLogAnalyticsRecord(minimalEntry, '2026-01-02T03:04:05.678Z', 'a');
    sendToLogAnalytics(r);
    sendToLogAnalytics({ ...r, EventId: 'b' });
    expect(mockUpload).not.toHaveBeenCalled();

    await jest.advanceTimersByTimeAsync(FLUSH_INTERVAL_MS);
    expect(mockUpload).toHaveBeenCalledTimes(1);
    expect(mockUpload.mock.calls[0][2].map(x => x.EventId)).toEqual(['a', 'b']);
  });

  it('uploads immediately once a batch is full', async () => {
    jest.useFakeTimers();
    configure();
    const r = toLogAnalyticsRecord(minimalEntry, '2026-01-02T03:04:05.678Z', 'a');
    for (let i = 0; i < MAX_BATCH; i++) sendToLogAnalytics({ ...r, EventId: String(i) });
    await Promise.resolve();
    expect(mockUpload).toHaveBeenCalledTimes(1);
    expect(mockUpload.mock.calls[0][2]).toHaveLength(MAX_BATCH);

    // Nothing left behind for the timer.
    await jest.advanceTimersByTimeAsync(FLUSH_INTERVAL_MS);
    expect(mockUpload).toHaveBeenCalledTimes(1);
  });
});

// ── Managed identity credential ──────────────────────────────────────────────

describe('ContainerAppManagedIdentityCredential', () => {
  const realFetch = global.fetch;

  afterEach(() => {
    global.fetch = realFetch;
    delete process.env.IDENTITY_ENDPOINT;
    delete process.env.IDENTITY_HEADER;
  });

  it('asks the Container Apps identity endpoint for a Monitor token', async () => {
    process.env.IDENTITY_ENDPOINT = 'http://localhost:42356/msi/token';
    process.env.IDENTITY_HEADER = 'hdr-value';
    const fetchMock = jest.fn<typeof fetch>().mockResolvedValue(
      new Response(JSON.stringify({ access_token: 'tok', expires_on: '1900000000' }), { status: 200 }),
    );
    global.fetch = fetchMock;

    const token = await new ContainerAppManagedIdentityCredential().getToken('https://monitor.azure.com//.default');

    expect(token).toEqual({ token: 'tok', expiresOnTimestamp: 1_900_000_000_000 });
    const [url, init] = fetchMock.mock.calls[0];
    const u = new URL(String(url));
    expect(u.searchParams.get('resource')).toBe('https://monitor.azure.com/');
    expect(u.searchParams.get('api-version')).toBe('2019-08-01');
    expect(u.searchParams.has('client_id')).toBe(false); // system-assigned identity
    expect((init?.headers as Record<string, string>)['X-IDENTITY-HEADER']).toBe('hdr-value');
  });

  it('fails clearly outside a Container App', async () => {
    await expect(new ContainerAppManagedIdentityCredential().getToken('https://monitor.azure.com//.default'))
      .rejects.toThrow(/IDENTITY_ENDPOINT/);
  });

  it('surfaces a non-200 from the identity endpoint', async () => {
    process.env.IDENTITY_ENDPOINT = 'http://localhost:42356/msi/token';
    process.env.IDENTITY_HEADER = 'hdr-value';
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue(new Response('no', { status: 400 }));
    await expect(new ContainerAppManagedIdentityCredential().getToken('https://monitor.azure.com//.default'))
      .rejects.toThrow(/HTTP 400/);
  });
});
