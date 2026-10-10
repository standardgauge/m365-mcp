import type { HttpRequest } from '@azure/functions';
import { AsyncLocalStorage } from 'async_hooks';
import { isIP } from 'net';

// The client address a request came from, for the rate limiter and the audit
// log.
//
// Container Apps ingress appends the address it received the connection from
// to X-Forwarded-For. Everything to the left of that entry was sent by the
// client and can say anything, so neither consumer may read it. With H trusted
// proxies in front of the app the client is the H-th entry from the right.
//
//   RATE_LIMIT_TRUSTED_PROXY_HOPS  how many proxies append to X-Forwarded-For.
//                                  Default 1, the Container Apps ingress. Set 2
//                                  behind Front Door or an Application Gateway.
//                                  Named for the rate limiter, which came first;
//                                  it governs the audit address too.

export function trustedProxyHops(): number {
  const n = Number(process.env.RATE_LIMIT_TRUSTED_PROXY_HOPS ?? '1');
  return Number.isInteger(n) && n >= 1 ? n : 1;
}

/** The X-Forwarded-For entry the outermost trusted proxy appended, port and all. */
export function forwardedClientEntry(request: HttpRequest): string | undefined {
  const header = request.headers?.get('x-forwarded-for');
  if (!header) return undefined;
  const entries = header.split(',').map((s) => s.trim()).filter(Boolean);
  if (entries.length === 0) return undefined;
  return entries[Math.max(0, entries.length - trustedProxyHops())];
}

/**
 * Reduce one X-Forwarded-For entry to a bare address: drop brackets and any
 * port, lower-case, and turn IPv4-mapped IPv6 into the IPv4 address.
 */
export function stripAddress(entry: string): string {
  let addr = entry.trim().toLowerCase();
  if (addr.startsWith('[')) {
    // [v6] or [v6]:port
    const close = addr.indexOf(']');
    if (close > 0) addr = addr.slice(1, close);
  } else if (addr.indexOf(':') === addr.lastIndexOf(':') && addr.includes(':')) {
    // v4:port (a bare v6 address always has at least two colons)
    addr = addr.slice(0, addr.indexOf(':'));
  }
  if (addr.startsWith('::ffff:') && isIP(addr.slice(7)) === 4) return addr.slice(7);
  return addr;
}

/**
 * The address to record on an audit row: the full ingress-appended address,
 * not the rate limiter's /64 bucket. Undefined when there is no header or the
 * selected entry is not an IP address, so nothing a client wrote lands in the
 * column.
 */
export function auditClientAddress(request: HttpRequest): string | undefined {
  const entry = forwardedClientEntry(request);
  if (!entry) return undefined;
  const addr = stripAddress(entry);
  return isIP(addr) ? addr : undefined;
}

// ── Per-request scope ────────────────────────────────────────────────────────
//
// withSecurity wraps every route and runs its handler inside this scope, so any
// audit row written while the request is being handled carries the address
// without each call site having to pass it down. That includes MCP tool calls
// several frames below the HTTP handler and the device-code sign-in, which
// completes in a promise the request started.

const scope = new AsyncLocalStorage<{ clientAddress?: string }>();

export function runWithClientAddress<T>(request: HttpRequest, fn: () => T): T {
  return scope.run({ clientAddress: auditClientAddress(request) }, fn);
}

/** The audit address of the request being handled, if any. */
export function currentClientAddress(): string | undefined {
  return scope.getStore()?.clientAddress;
}
