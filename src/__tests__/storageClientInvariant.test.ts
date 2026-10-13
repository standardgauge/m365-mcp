/**
 * Every Table Storage client in the server comes from services/storageClient.ts.
 *
 * Production configures storage with AZURE_STORAGE_TABLE_ENDPOINT and a managed
 * identity (infra/container-app.bicep); only local development uses a
 * connection string. A module that called TableClient.fromConnectionString
 * itself therefore worked against Azurite and failed in every deployment:
 * outboundPolicy.ts did exactly that, dialled 127.0.0.1:10002 in production,
 * and because enforceOutboundPolicy fails closed, refused every calendar
 * invite and Teams send on every deploy fork. This test holds the one place
 * where a connection string may be turned into a client.
 */
import { jest } from '@jest/globals';
import { readdirSync, readFileSync, statSync } from 'fs';
import { join, relative } from 'path';

const SRC = join(__dirname, '..');
const ALLOWED = new Set(['services/storageClient.ts']);

function walk(dir: string, out: string[] = []): string[] {
  for (const name of readdirSync(dir)) {
    const p = join(dir, name);
    if (statSync(p).isDirectory()) {
      if (name === '__tests__' || name === 'node_modules') continue;
      walk(p, out);
    } else if (/\.tsx?$/.test(name)) {
      out.push(p);
    }
  }
  return out;
}

describe('Table Storage clients come from storageClient.ts', () => {
  it('no other source file builds a client from a connection string or news one up', () => {
    const offenders: string[] = [];
    for (const file of walk(SRC)) {
      const rel = relative(SRC, file);
      if (ALLOWED.has(rel)) continue;
      const text = readFileSync(file, 'utf8');
      if (/fromConnectionString\(/.test(text) || /new Table(Service)?Client\(/.test(text)) offenders.push(rel);
    }
    expect(offenders).toEqual([]);
  });

  it('outboundPolicy asks the shared client for its table', async () => {
    jest.resetModules();
    const getTableClient = jest.fn(() => ({
      createTable: async () => undefined,
      getEntity: async () => { throw Object.assign(new Error('not found'), { statusCode: 404 }); },
    }));
    jest.doMock('../services/storageClient.js', () => ({ getTableClient }));
    const { getOutboundPolicy } = await import('../services/outboundPolicy.js');
    const policy = await getOutboundPolicy('tenant-a', { scope: 'tenant' });
    expect(getTableClient).toHaveBeenCalledWith('OutboundPolicy');
    expect(policy).toMatchObject({ calendarInvites: 'allow', eventResponses: 'allow', teamsMessages: 'allow' });
  });
});
