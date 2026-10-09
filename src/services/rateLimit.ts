import type { HttpHandler, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { isIP } from 'net';

// Per-client-address request limiter for the unauthenticated public endpoints
// (login, device, install-poll, mcp).
//
// Container Apps ingress has no native rate limit, and these routes answer
// anyone. `device` is the expensive one: every call starts an MSAL device-code
// flow that polls Entra in the background for up to fifteen minutes.
//
// Fixed one-minute window, counted in this replica's memory. With N replicas the
// effective ceiling for one address is up to N times the configured limit; that
// is accepted. The point is to make the endpoints cost something to hammer, not
// to meter them exactly, and a shared store would put a Table Storage round trip
// in front of every MCP call.
//
// Configuration, all optional:
//   RATE_LIMIT_<ROUTE>_PER_MINUTE  requests per address per minute for that
//                                  route (LOGIN, DEVICE, INSTALL_POLL, MCP).
//                                  0 turns the limiter off for the route.
//   RATE_LIMIT_TRUSTED_PROXY_HOPS  how many proxies in front of the app append
//                                  to X-Forwarded-For. Default 1, the Container
//                                  Apps ingress. Set 2 behind Front Door or an
//                                  Application Gateway.

export type RateLimitedRoute = 'login' | 'device' | 'install-poll' | 'mcp';

// Defaults sized for a whole office behind one NAT address, not one user.
// install-poll: the install scripts poll every 2s, so 30/min per install.
// mcp: one request per tool call; an agent can burst, several share an address.
export const DEFAULT_LIMITS: Record<RateLimitedRoute, number> = {
  login: 30,
  device: 10,
  'install-poll': 120,
  mcp: 1200,
};

const WINDOW_MS = 60_000;

// Bounds memory when an attacker rotates addresses. Past this, the oldest
// window is evicted; an evicted address simply starts a fresh window.
export const MAX_TRACKED_ADDRESSES = 10_000;

interface Window {
  start: number;
  count: number;
}

const windows = new Map<string, Window>();

function envName(route: RateLimitedRoute): string {
  return `RATE_LIMIT_${route.replace('-', '_').toUpperCase()}_PER_MINUTE`;
}

export function limitFor(route: RateLimitedRoute): number {
  const raw = process.env[envName(route)];
  if (raw === undefined || raw.trim() === '') return DEFAULT_LIMITS[route];
  const n = Number(raw);
  if (!Number.isInteger(n) || n < 0) return DEFAULT_LIMITS[route];
  return n;
}

function trustedHops(): number {
  const n = Number(process.env.RATE_LIMIT_TRUSTED_PROXY_HOPS ?? '1');
  return Number.isInteger(n) && n >= 1 ? n : 1;
}

// Reduce one X-Forwarded-For entry to a bucket key: drop any port, and key
// IPv6 by its /64 so one host cannot rotate through its own prefix.
export function normalizeAddress(entry: string): string {
  let addr = entry.trim().toLowerCase();
  if (addr.startsWith('[')) {
    // [v6] or [v6]:port
    const close = addr.indexOf(']');
    if (close > 0) addr = addr.slice(1, close);
  } else if (addr.indexOf(':') === addr.lastIndexOf(':') && addr.includes(':')) {
    // v4:port (a bare v6 address always has at least two colons)
    addr = addr.slice(0, addr.indexOf(':'));
  }
  if (isIP(addr) !== 6) return addr;

  // IPv4-mapped IPv6 is the IPv4 address.
  if (addr.startsWith('::ffff:') && isIP(addr.slice(7)) === 4) return addr.slice(7);

  const [head, tail] = addr.split('::');
  const left = head ? head.split(':') : [];
  const right = tail ? tail.split(':') : [];
  const groups = tail === undefined ? left : [...left, ...Array(Math.max(0, 8 - left.length - right.length)).fill('0'), ...right];
  return groups.slice(0, 4).map((g) => g.replace(/^0+(?=.)/, '')).join(':') + '::/64';
}

// The client address as seen by the outermost trusted proxy. Each proxy appends
// the peer it received the connection from, so with H trusted hops the client
// is the H-th entry from the right. Entries further left were supplied by the
// client and are ignored: keying on them would let a caller pick its own bucket.
export function clientAddress(request: HttpRequest): string {
  const header = request.headers?.get('x-forwarded-for');
  if (!header) return 'unknown';
  const entries = header.split(',').map((s) => s.trim()).filter(Boolean);
  if (entries.length === 0) return 'unknown';
  const idx = Math.max(0, entries.length - trustedHops());
  return normalizeAddress(entries[idx]);
}

/**
 * Count one request against `key`. Returns 0 when allowed, otherwise the
 * seconds until the window resets.
 */
export function take(key: string, limit: number, now: number = Date.now()): number {
  let w = windows.get(key);
  if (!w || now - w.start >= WINDOW_MS) {
    windows.delete(key);
    w = { start: now, count: 0 };
    windows.set(key, w);
    while (windows.size > MAX_TRACKED_ADDRESSES) {
      const oldest = windows.keys().next().value as string;
      windows.delete(oldest);
    }
  }
  w.count += 1;
  if (w.count <= limit) return 0;
  return Math.max(1, Math.ceil((w.start + WINDOW_MS - now) / 1000));
}

export function resetRateLimits(): void {
  windows.clear();
}

export function trackedAddressCount(): number {
  return windows.size;
}

/**
 * Wrap a handler so requests over the route's per-address limit get 429 with
 * Retry-After and never reach it. Place inside withSecurity so the 429 still
 * carries the security headers.
 */
export function withRateLimit(route: RateLimitedRoute, handler: HttpHandler): HttpHandler {
  return async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    const limit = limitFor(route);
    if (limit === 0) return (await handler(request, context)) ?? {};

    const address = clientAddress(request);
    const retryAfter = take(`${route}|${address}`, limit);
    if (retryAfter === 0) return (await handler(request, context)) ?? {};

    // Log once per window per address, on the first refusal, not on every one.
    if (windows.get(`${route}|${address}`)?.count === limit + 1) {
      context.warn(`rate limit: ${route} refused ${address} over ${limit}/min`);
    }
    const message = 'Too many requests';
    return {
      status: 429,
      headers: { 'Retry-After': String(retryAfter) },
      jsonBody:
        route === 'mcp'
          ? { jsonrpc: '2.0', id: null, error: { code: -32000, message } }
          : { error: message, retryAfterSeconds: retryAfter },
    };
  };
}
