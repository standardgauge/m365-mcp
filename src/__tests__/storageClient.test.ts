/**
 * How the server reaches Table Storage: an Entra token for the app's managed
 * identity when AZURE_STORAGE_TABLE_ENDPOINT is set, the account-key
 * connection string otherwise, and Azurite when neither is.
 */
import { jest } from '@jest/globals';
import { createHttpHeaders, type PipelinePolicy, type PipelineRequest } from '@azure/core-rest-pipeline';
import type { AccessToken, TokenCredential } from '@azure/core-auth';
import {
  getTableClient,
  getTableServiceClient,
  isStorageConfigured,
  __resetStorageClientForTests,
} from '../services/storageClient.js';
import {
  CachedTokenCredential,
  ContainerAppManagedIdentityCredential,
  TOKEN_REFRESH_MARGIN_MS,
} from '../services/managedIdentity.js';

const ENV_KEYS = [
  'AZURE_STORAGE_TABLE_ENDPOINT',
  'AZURE_STORAGE_IDENTITY_CLIENT_ID',
  'AZURE_STORAGE_CONNECTION_STRING',
  'IDENTITY_ENDPOINT',
  'IDENTITY_HEADER',
] as const;

const realFetch = global.fetch;

beforeEach(() => {
  for (const k of ENV_KEYS) delete process.env[k];
  __resetStorageClientForTests();
});

afterEach(() => {
  global.fetch = realFetch;
  for (const k of ENV_KEYS) delete process.env[k];
});

/** Answer every request with a 404 after signing, and record what was sent. */
function captureRequests(sent: PipelineRequest[]): PipelinePolicy {
  return {
    name: 'capture',
    sendRequest: async (request) => {
      sent.push(request);
      return {
        request,
        status: 404,
        headers: createHttpHeaders({ 'content-type': 'application/json' }),
        bodyAsText: JSON.stringify({ 'odata.error': { code: 'ResourceNotFound', message: { value: 'no' } } }),
      };
    },
  };
}

function mockIdentityEndpoint(): jest.Mock<typeof fetch> {
  process.env.IDENTITY_ENDPOINT = 'http://localhost:42356/msi/token';
  process.env.IDENTITY_HEADER = 'hdr-value';
  const fetchMock = jest.fn<typeof fetch>().mockImplementation(async () =>
    new Response(JSON.stringify({ access_token: 'mi-token', expires_on: String(Math.floor(Date.now() / 1000) + 3600) }), { status: 200 }),
  );
  global.fetch = fetchMock;
  return fetchMock;
}

describe('storage client selection', () => {
  it('uses the managed identity against the table endpoint when it is set', async () => {
    process.env.AZURE_STORAGE_TABLE_ENDPOINT = 'https://stexample.table.core.windows.net';
    process.env.AZURE_STORAGE_IDENTITY_CLIENT_ID = '00000000-0000-0000-0000-000000000001';
    // A connection string left over from older infra must not win.
    process.env.AZURE_STORAGE_CONNECTION_STRING =
      'DefaultEndpointsProtocol=https;AccountName=old;AccountKey=a2V5;EndpointSuffix=core.windows.net';
    const fetchMock = mockIdentityEndpoint();

    expect(isStorageConfigured()).toBe(true);
    const client = getTableClient('mcpSessions');
    expect(client.url).toBe('https://stexample.table.core.windows.net');
    const sent: PipelineRequest[] = [];
    client.pipeline.addPolicy(captureRequests(sent), { afterPhase: 'Sign' });

    await expect(client.getEntity('session', 'row')).rejects.toBeDefined();

    expect(sent).toHaveLength(1);
    expect(sent[0].url).toContain('https://stexample.table.core.windows.net/mcpSessions');
    expect(sent[0].headers.get('authorization')).toBe('Bearer mi-token');
    const tokenUrl = new URL(String(fetchMock.mock.calls[0][0]));
    expect(tokenUrl.searchParams.get('resource')).toBe('https://storage.azure.com');
    expect(tokenUrl.searchParams.get('client_id')).toBe('00000000-0000-0000-0000-000000000001');
  });

  it('shares one token across clients built per call', async () => {
    process.env.AZURE_STORAGE_TABLE_ENDPOINT = 'https://stexample.table.core.windows.net';
    const fetchMock = mockIdentityEndpoint();

    for (const table of ['GlobalDenyList', 'UserDenyList', 'serviceSettings']) {
      const client = getTableClient(table);
      client.pipeline.addPolicy(captureRequests([]), { afterPhase: 'Sign' });
      await expect(client.getEntity('p', 'r')).rejects.toBeDefined();
    }
    const svc = getTableServiceClient();
    svc.pipeline.addPolicy(captureRequests([]), { afterPhase: 'Sign' });
    await expect(svc.createTable('auditLog')).rejects.toBeDefined();

    expect(fetchMock).toHaveBeenCalledTimes(1);
    // System-assigned identity when no client id is configured.
    expect(new URL(String(fetchMock.mock.calls[0][0])).searchParams.has('client_id')).toBe(false);
  });

  it('falls back to the connection string', () => {
    process.env.AZURE_STORAGE_CONNECTION_STRING =
      'DefaultEndpointsProtocol=https;AccountName=stexample;AccountKey=a2V5;EndpointSuffix=core.windows.net';
    expect(isStorageConfigured()).toBe(true);
    expect(getTableClient('mcpSessions').url).toBe('https://stexample.table.core.windows.net');
  });

  it('falls back to Azurite with nothing configured', () => {
    expect(isStorageConfigured()).toBe(false);
    expect(getTableClient('mcpSessions').url).toContain('127.0.0.1:10002');
  });
});

describe('CachedTokenCredential', () => {
  function inner(expiresInMs: number): { cred: TokenCredential; calls: () => number } {
    let n = 0;
    const cred: TokenCredential = {
      getToken: async (): Promise<AccessToken> => {
        n += 1;
        return { token: `t${n}`, expiresOnTimestamp: Date.now() + expiresInMs };
      },
    };
    return { cred, calls: () => n };
  }

  it('reuses a token until it is close to expiry', async () => {
    const { cred, calls } = inner(TOKEN_REFRESH_MARGIN_MS + 60_000);
    const cached = new CachedTokenCredential(cred);
    await cached.getToken('https://storage.azure.com/.default');
    const again = await cached.getToken('https://storage.azure.com/.default');
    expect(again.token).toBe('t1');
    expect(calls()).toBe(1);
  });

  it('asks again inside the refresh margin', async () => {
    const { cred, calls } = inner(TOKEN_REFRESH_MARGIN_MS - 1);
    const cached = new CachedTokenCredential(cred);
    await cached.getToken('s');
    expect((await cached.getToken('s')).token).toBe('t2');
    expect(calls()).toBe(2);
  });

  it('shares one in-flight request between concurrent callers', async () => {
    const { cred, calls } = inner(3_600_000);
    const cached = new CachedTokenCredential(cred);
    const tokens = await Promise.all([cached.getToken('s'), cached.getToken('s'), cached.getToken('s')]);
    expect(tokens.map((t) => t.token)).toEqual(['t1', 't1', 't1']);
    expect(calls()).toBe(1);
  });

  it('does not cache a failure', async () => {
    let fail = true;
    const cached = new CachedTokenCredential({
      getToken: async () => {
        if (fail) throw new Error('endpoint down');
        return { token: 'ok', expiresOnTimestamp: Date.now() + 3_600_000 };
      },
    });
    await expect(cached.getToken('s')).rejects.toThrow('endpoint down');
    fail = false;
    expect((await cached.getToken('s')).token).toBe('ok');
  });
});

describe('ContainerAppManagedIdentityCredential client id', () => {
  it('selects a user-assigned identity by client id', async () => {
    const fetchMock = mockIdentityEndpoint();
    await new ContainerAppManagedIdentityCredential('cid-1').getToken('https://storage.azure.com/.default');
    expect(new URL(String(fetchMock.mock.calls[0][0])).searchParams.get('client_id')).toBe('cid-1');
  });
});
