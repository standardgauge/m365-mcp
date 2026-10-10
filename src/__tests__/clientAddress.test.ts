/**
 * The client address recorded on audit rows.
 *
 * Only the X-Forwarded-For entry the ingress appended may be recorded:
 * everything to its left was written by the client. The address is kept whole
 * (the rate limiter's /64 bucketing is not applied), and it reaches every row
 * written while a request is handled through withSecurity's per-request scope,
 * including MCP rows whose call sites never see the request.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

const mockUpsertEntity = jest.fn<(entity: Record<string, unknown>, mode: string) => Promise<void>>();

jest.mock('@azure/data-tables', () => ({
  TableClient: {
    fromConnectionString: () => ({
      upsertEntity: (...args: unknown[]) => mockUpsertEntity(args[0] as Record<string, unknown>, args[1] as string),
      listEntities: () => ({ [Symbol.asyncIterator]: async function* () { /* empty */ } }),
    }),
  },
  TableServiceClient: {
    fromConnectionString: () => ({ createTable: async () => undefined }),
  },
}));

process.env.AZURE_STORAGE_CONNECTION_STRING =
  'DefaultEndpointsProtocol=https;AccountName=test;AccountKey=dGVzdA==;EndpointSuffix=core.windows.net';
process.env.AUDIT_LOG_RETENTION_DAYS = '0';

import {
  auditClientAddress,
  currentClientAddress,
  runWithClientAddress,
} from '../services/clientAddress.js';
import { clientAddress } from '../services/rateLimit.js';
import { readdirSync, readFileSync } from 'fs';
import { join } from 'path';
import { auditSnapshot, AUDIT_SNAPSHOT_MAX_CHARS, logAccess } from '../services/auditLog.js';
import { withSecurity } from '../services/securityHeaders.js';

function req(xff?: string): HttpRequest {
  const headers = new Map<string, string>();
  if (xff !== undefined) headers.set('x-forwarded-for', xff);
  return { headers } as unknown as HttpRequest;
}

const flush = () => new Promise((resolve) => setImmediate(resolve));

async function writtenIp(): Promise<unknown> {
  for (let i = 0; i < 5 && mockUpsertEntity.mock.calls.length === 0; i++) await flush();
  return mockUpsertEntity.mock.calls[0]?.[0].ip;
}

const ENTRY = {
  tenantId: 't', userId: 'u', userEmail: 'adele@fabrikam.com',
  operation: 'auth.login', result: 'allowed' as const, source: 'http' as const,
};

beforeEach(() => {
  mockUpsertEntity.mockReset();
  mockUpsertEntity.mockResolvedValue(undefined);
  delete process.env.RATE_LIMIT_TRUSTED_PROXY_HOPS;
});

describe('auditClientAddress', () => {
  it('takes the entry the ingress appended, not the one the client sent', () => {
    expect(auditClientAddress(req('198.51.100.66, 203.0.113.9'))).toBe('203.0.113.9');
  });

  it('honours RATE_LIMIT_TRUSTED_PROXY_HOPS', () => {
    process.env.RATE_LIMIT_TRUSTED_PROXY_HOPS = '2';
    expect(auditClientAddress(req('198.51.100.66, 203.0.113.9, 192.0.2.1'))).toBe('203.0.113.9');
  });

  it('keeps an IPv6 address whole where the rate limiter buckets it by /64', () => {
    const r = req('2001:db8:1:2:3:4:5:6');
    expect(auditClientAddress(r)).toBe('2001:db8:1:2:3:4:5:6');
    expect(clientAddress(r)).toBe('2001:db8:1:2::/64');
  });

  it.each([
    ['203.0.113.9:51234', '203.0.113.9'],
    ['[2001:db8::7]:443', '2001:db8::7'],
    ['::ffff:203.0.113.9', '203.0.113.9'],
  ])('reduces %s to %s', (entry, expected) => {
    expect(auditClientAddress(req(entry))).toBe(expected);
  });

  it('records nothing when there is no header or the entry is not an address', () => {
    expect(auditClientAddress(req())).toBeUndefined();
    expect(auditClientAddress(req('   '))).toBeUndefined();
    expect(auditClientAddress(req('<script>'))).toBeUndefined();
  });
});

describe('per-request scope', () => {
  it('is empty outside a request', () => {
    expect(currentClientAddress()).toBeUndefined();
  });

  it('survives awaits inside the handler', async () => {
    const seen = await runWithClientAddress(req('198.51.100.66, 203.0.113.9'), async () => {
      await flush();
      return currentClientAddress();
    });
    expect(seen).toBe('203.0.113.9');
  });

  it('gives a row written inside withSecurity the request address', async () => {
    const handler = withSecurity(async () => {
      logAccess(ENTRY);
      return { status: 200 };
    });
    await handler(req('198.51.100.66, 203.0.113.9'), {} as InvocationContext);
    expect(await writtenIp()).toBe('203.0.113.9');
  });

  it('lets an explicit address win', async () => {
    await runWithClientAddress(req('203.0.113.9'), async () => logAccess({ ...ENTRY, ip: '192.0.2.44' }));
    expect(await writtenIp()).toBe('192.0.2.44');
  });

  it('writes null outside a request', async () => {
    logAccess(ENTRY);
    expect(await writtenIp()).toBeNull();
  });

  // The scope is the only thing that gives an MCP row its address, so a route
  // registered without withSecurity would silently write rows with none.
  it('wraps every registered route in withSecurity', () => {
    const root = join(__dirname, '..', 'functions');
    const files = readdirSync(root, { recursive: true, encoding: 'utf8' }).filter((f) => f.endsWith('.ts'));
    let routes = 0;
    for (const f of files) {
      const src = readFileSync(join(root, f), 'utf8');
      for (const m of src.matchAll(/app\.http\([^]*?handler:\s*([A-Za-z]+)/g)) {
        routes++;
        expect([f, m[1]]).toEqual([f, 'withSecurity']);
      }
    }
    expect(routes).toBeGreaterThan(50);
  });

  // An explicit ip on a row wins over the scope, so a call site that reads the
  // header itself would record the client-written entries again.
  it('reads X-Forwarded-For only in clientAddress.ts', () => {
    const root = join(__dirname, '..');
    const files = readdirSync(root, { recursive: true, encoding: 'utf8' })
      .filter((f) => f.endsWith('.ts') && !f.startsWith('__tests__'));
    const readers = files.filter((f) => /x-forwarded-for['"]\s*\)/i.test(readFileSync(join(root, f), 'utf8')));
    expect(readers).toEqual([join('services', 'clientAddress.ts')]);
  });

  it('stores before and after on the row', async () => {
    logAccess({ ...ENTRY, operation: 'policy.services.set', before: '["mail"]', after: '["mail","calendar"]' });
    for (let i = 0; i < 5 && mockUpsertEntity.mock.calls.length === 0; i++) await flush();
    expect(mockUpsertEntity.mock.calls[0][0]).toMatchObject({ before: '["mail"]', after: '["mail","calendar"]' });
  });
});

describe('auditSnapshot', () => {
  it('serialises a value as JSON, and absence as null', () => {
    expect(auditSnapshot(['mail'])).toBe('["mail"]');
    expect(auditSnapshot(undefined)).toBe('null');
  });

  it('cuts a value past the cap and says how much was cut', () => {
    const out = auditSnapshot('x'.repeat(AUDIT_SNAPSHOT_MAX_CHARS + 100));
    expect(out.startsWith('"xxx')).toBe(true);
    expect(out).toMatch(/…\[truncated 102 chars\]$/);
    expect(out.length).toBeLessThan(AUDIT_SNAPSHOT_MAX_CHARS + 40);
  });
});
