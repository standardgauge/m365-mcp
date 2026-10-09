/**
 * Wire-level check for the retention purge: runs the real @azure/data-tables
 * client against a fake HTTP client and asserts on what reaches the service.
 *
 * The SDK maps the entity field `timestamp` to the service-owned `Timestamp`
 * system property on write, but it does not rewrite raw OData filter strings.
 * So the purge filter has to name `Timestamp` itself; a filter on lowercase
 * `timestamp` targets a column that never exists and deletes nothing.
 */

import { jest } from '@jest/globals';
import type { HttpClient, PipelineRequest, PipelineResponse } from '@azure/core-rest-pipeline';

const requests: { method: string; url: string; body: string }[] = [];

jest.mock('@azure/data-tables', () => {
  const actual = jest.requireActual<typeof import('@azure/data-tables')>('@azure/data-tables');
  const { createHttpHeaders } = jest.requireActual<typeof import('@azure/core-rest-pipeline')>('@azure/core-rest-pipeline');
  const httpClient: HttpClient = {
    async sendRequest(request: PipelineRequest): Promise<PipelineResponse> {
      requests.push({ method: request.method, url: request.url, body: typeof request.body === 'string' ? request.body : '' });
      const isList = request.method === 'GET';
      return {
        request,
        status: isList ? 200 : 204,
        headers: createHttpHeaders({ 'content-type': 'application/json;odata=nometadata' }),
        bodyAsText: isList ? JSON.stringify({ value: [] }) : '',
      };
    },
  };
  return {
    ...actual,
    TableClient: {
      fromConnectionString: (cs: string, table: string) =>
        actual.TableClient.fromConnectionString(cs, table, { httpClient, retryOptions: { maxRetries: 0 } }),
    },
    TableServiceClient: {
      fromConnectionString: (cs: string) =>
        actual.TableServiceClient.fromConnectionString(cs, { httpClient, retryOptions: { maxRetries: 0 } }),
    },
  };
});

process.env.AZURE_STORAGE_CONNECTION_STRING =
  'DefaultEndpointsProtocol=https;AccountName=test;AccountKey=dGVzdA==;EndpointSuffix=core.windows.net';

import { logAccess, purgeAuditLog, resetAuditPurgeStateForTests } from '../services/auditLog.js';

const DAY_MS = 24 * 60 * 60 * 1000;
const NOW = Date.parse('2026-10-09T12:00:00.000Z');
const flush = () => new Promise((resolve) => setImmediate(resolve));

beforeEach(() => {
  requests.length = 0;
  resetAuditPurgeStateForTests();
  jest.spyOn(console, 'log').mockImplementation(() => undefined);
  jest.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  jest.restoreAllMocks();
});

describe('audit log retention at the storage boundary', () => {
  it('writes the entry time as the service Timestamp, not a lowercase timestamp column', async () => {
    process.env.AUDIT_LOG_RETENTION_DAYS = '0';
    logAccess({
      tenantId: 'tenant-1', userId: 'u1', userEmail: 'user@example.com',
      operation: 'mail.search_mail', result: 'allowed', source: 'mcp',
    });
    for (let i = 0; i < 20 && !requests.some((r) => r.method === 'PUT'); i++) await flush();
    delete process.env.AUDIT_LOG_RETENTION_DAYS;

    const upsert = requests.find((r) => r.method === 'PUT');
    expect(upsert).toBeDefined();
    const body = JSON.parse(upsert!.body) as Record<string, unknown>;
    expect(body).toHaveProperty('Timestamp');
    expect(body).not.toHaveProperty('timestamp');
  });

  it('sends a $filter on Timestamp with a datetime literal', async () => {
    await purgeAuditLog(90, NOW);

    const list = requests.find((r) => r.method === 'GET');
    expect(list).toBeDefined();
    const filter = new URL(list!.url).searchParams.get('$filter');
    const cutoff = new Date(NOW - 90 * DAY_MS).toISOString();
    expect(filter).toBe(`Timestamp lt datetime'${cutoff}'`);
  });
});
