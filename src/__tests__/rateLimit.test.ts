/**
 * Per-address limiter on the unauthenticated public endpoints.
 *
 * Covers the window arithmetic, the client-address derivation (rightmost
 * X-Forwarded-For entry by default, so a caller cannot choose its own bucket by
 * prepending addresses), the IPv6 /64 grouping, the memory bound, the env
 * overrides, and that each of the four routes is registered behind it.
 */

import { jest } from '@jest/globals';
import type { HttpHandler, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import {
  clientAddress,
  DEFAULT_LIMITS,
  limitFor,
  MAX_TRACKED_ADDRESSES,
  normalizeAddress,
  resetRateLimits,
  take,
  trackedAddressCount,
  withRateLimit,
} from '../services/rateLimit.js';

function req(xff?: string): HttpRequest {
  const headers = new Map<string, string>();
  if (xff !== undefined) headers.set('x-forwarded-for', xff);
  return { headers } as unknown as HttpRequest;
}

const ctx = { warn: jest.fn() } as unknown as InvocationContext;

const ENV_KEYS = [
  'RATE_LIMIT_LOGIN_PER_MINUTE',
  'RATE_LIMIT_DEVICE_PER_MINUTE',
  'RATE_LIMIT_INSTALL_POLL_PER_MINUTE',
  'RATE_LIMIT_MCP_PER_MINUTE',
  'RATE_LIMIT_TRUSTED_PROXY_HOPS',
];

beforeEach(() => {
  resetRateLimits();
  for (const k of ENV_KEYS) delete process.env[k];
  (ctx.warn as jest.Mock).mockClear();
});

describe('take', () => {
  it('allows up to the limit in a window, then reports seconds to reset', () => {
    const t0 = 1_000_000;
    for (let i = 0; i < 3; i++) expect(take('k', 3, t0)).toBe(0);
    expect(take('k', 3, t0 + 15_000)).toBe(45);
    expect(take('k', 3, t0 + 59_900)).toBe(1);
  });

  it('starts a fresh window after a minute', () => {
    const t0 = 1_000_000;
    for (let i = 0; i < 4; i++) take('k', 3, t0);
    expect(take('k', 3, t0 + 60_000)).toBe(0);
  });

  it('keeps buckets independent', () => {
    take('a', 1, 0);
    expect(take('a', 1, 0)).toBeGreaterThan(0);
    expect(take('b', 1, 0)).toBe(0);
  });

  it('never tracks more than the address cap', () => {
    for (let i = 0; i < MAX_TRACKED_ADDRESSES + 50; i++) take(`k${i}`, 5, 0);
    expect(trackedAddressCount()).toBe(MAX_TRACKED_ADDRESSES);
  });
});

describe('clientAddress', () => {
  it('uses the entry the ingress appended, not one the client supplied', () => {
    expect(clientAddress(req('198.51.100.7, 203.0.113.9'))).toBe('203.0.113.9');
  });

  it('honours more trusted hops behind an extra proxy', () => {
    process.env.RATE_LIMIT_TRUSTED_PROXY_HOPS = '2';
    expect(clientAddress(req('198.51.100.7, 203.0.113.9, 192.0.2.1'))).toBe('203.0.113.9');
  });

  it('falls back to the leftmost entry when there are fewer entries than hops', () => {
    process.env.RATE_LIMIT_TRUSTED_PROXY_HOPS = '3';
    expect(clientAddress(req('203.0.113.9'))).toBe('203.0.113.9');
  });

  it('buckets a request with no forwarded header as unknown', () => {
    expect(clientAddress(req())).toBe('unknown');
    expect(clientAddress({} as HttpRequest)).toBe('unknown');
  });
});

describe('normalizeAddress', () => {
  it.each([
    ['203.0.113.9', '203.0.113.9'],
    ['203.0.113.9:51234', '203.0.113.9'],
    ['::ffff:203.0.113.9', '203.0.113.9'],
    ['2001:db8:aa:bb:1:2:3:4', '2001:db8:aa:bb::/64'],
    ['2001:db8:aa:bb::99', '2001:db8:aa:bb::/64'],
    ['[2001:db8:aa:bb::99]:443', '2001:db8:aa:bb::/64'],
    ['2001:0DB8::1', '2001:db8:0:0::/64'],
  ])('%s → %s', (input, expected) => {
    expect(normalizeAddress(input)).toBe(expected);
  });
});

describe('limitFor', () => {
  it('uses the defaults when unset or invalid', () => {
    expect(limitFor('device')).toBe(DEFAULT_LIMITS.device);
    process.env.RATE_LIMIT_DEVICE_PER_MINUTE = 'lots';
    expect(limitFor('device')).toBe(DEFAULT_LIMITS.device);
    process.env.RATE_LIMIT_DEVICE_PER_MINUTE = '-1';
    expect(limitFor('device')).toBe(DEFAULT_LIMITS.device);
  });

  it('reads the per-route override, including the hyphenated route', () => {
    process.env.RATE_LIMIT_INSTALL_POLL_PER_MINUTE = '7';
    expect(limitFor('install-poll')).toBe(7);
    process.env.RATE_LIMIT_MCP_PER_MINUTE = '0';
    expect(limitFor('mcp')).toBe(0);
  });
});

describe('withRateLimit', () => {
  const ok: HttpHandler = async () => ({ status: 200 });

  async function call(h: HttpHandler, xff: string): Promise<HttpResponseInit> {
    return (await h(req(xff), ctx)) as HttpResponseInit;
  }

  it('returns 429 with Retry-After once the address is over the limit', async () => {
    process.env.RATE_LIMIT_LOGIN_PER_MINUTE = '2';
    const inner = jest.fn(ok);
    const h = withRateLimit('login', inner as HttpHandler);
    expect((await call(h, '203.0.113.9')).status).toBe(200);
    expect((await call(h, '203.0.113.9')).status).toBe(200);
    const refused = await call(h, '203.0.113.9');
    expect(refused.status).toBe(429);
    expect(Number((refused.headers as Record<string, string>)['Retry-After'])).toBeGreaterThan(0);
    expect(inner).toHaveBeenCalledTimes(2);
    expect((await call(h, '203.0.113.10')).status).toBe(200);
  });

  it('cannot be dodged by prepending addresses to X-Forwarded-For', async () => {
    process.env.RATE_LIMIT_DEVICE_PER_MINUTE = '1';
    const h = withRateLimit('device', ok);
    await call(h, '10.0.0.1, 203.0.113.9');
    expect((await call(h, '10.0.0.2, 203.0.113.9')).status).toBe(429);
  });

  it('logs the first refusal in a window only', async () => {
    process.env.RATE_LIMIT_LOGIN_PER_MINUTE = '1';
    const h = withRateLimit('login', ok);
    for (let i = 0; i < 4; i++) await call(h, '203.0.113.9');
    expect(ctx.warn).toHaveBeenCalledTimes(1);
  });

  it('answers the MCP route with a JSON-RPC error body', async () => {
    process.env.RATE_LIMIT_MCP_PER_MINUTE = '1';
    const h = withRateLimit('mcp', ok);
    await call(h, '203.0.113.9');
    const refused = await call(h, '203.0.113.9');
    expect(refused.status).toBe(429);
    expect(refused.jsonBody).toMatchObject({ jsonrpc: '2.0', error: { code: -32000 } });
  });

  it('keeps routes in separate buckets', async () => {
    process.env.RATE_LIMIT_LOGIN_PER_MINUTE = '1';
    process.env.RATE_LIMIT_DEVICE_PER_MINUTE = '1';
    await call(withRateLimit('login', ok), '203.0.113.9');
    expect((await call(withRateLimit('device', ok), '203.0.113.9')).status).toBe(200);
  });

  it('passes everything through when the route limit is 0', async () => {
    process.env.RATE_LIMIT_LOGIN_PER_MINUTE = '0';
    const h = withRateLimit('login', ok);
    for (let i = 0; i < 50; i++) expect((await call(h, '203.0.113.9')).status).toBe(200);
    expect(trackedAddressCount()).toBe(0);
  });
});

describe('route wiring', () => {
  // Each public unauthenticated route must register its handler through
  // withRateLimit. A source check, because the handlers' own imports
  // (MSAL, Graph, Table Storage) make importing all four here expensive.
  it.each([
    ['functions/auth/login.ts', 'login'],
    ['functions/auth/deviceLogin.ts', 'device'],
    ['functions/auth/installPoll.ts', 'install-poll'],
    ['functions/mcp/mcpEndpoint.ts', 'mcp'],
  ])('%s is behind the %s limiter', async (file, route) => {
    const { readFileSync } = await import('fs');
    const { join } = await import('path');
    const src = readFileSync(join(__dirname, '..', file), 'utf8');
    expect(src).toContain(`withSecurity(withRateLimit('${route}',`);
  });
});
