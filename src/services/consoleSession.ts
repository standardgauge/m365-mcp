import type { HttpRequest } from '@azure/functions';
import { timingSafeEqual } from 'crypto';
import { macWithSessionKey } from './credentialCrypto.js';

/**
 * The console session: the browser-only credential the admin SPA uses for
 * /api/manage/*.
 *
 * Every client authenticates with the same kind of session token, and the
 * install flow hands that token to the MCP client. If /api/manage/* accepted it
 * too, a Global Administrator's MCP config, keychain or desktop extension would
 * hold an admin-API credential (threat model 7.3). So the manage routes accept
 * only a request that carries both:
 *
 *   - the `mcp_session` cookie, and
 *   - an `mcp_console` cookie minted by the OAuth callback and bound by MAC to
 *     that exact session token.
 *
 * The console token is minted nowhere except the interactive callback, so
 * holding a session token, by bearer header or by forging a Cookie header,
 * is not enough to reach the admin API. Because the MAC covers the session
 * token, a console cookie is useless next to any other session, and it dies
 * with its session: logout, or any other deletion of the session row, ends it.
 *
 * It is short-lived: 30 minutes idle, renewed by /api/auth/me, and never past
 * 8 hours from the sign-in that minted it. After that the SPA sends the user
 * back through sign-in, which with an existing Microsoft session is a redirect
 * bounce rather than a prompt.
 *
 * Stateless by design: there is no table to write, and nothing to clean up.
 */

export const CONSOLE_COOKIE = 'mcp_console';

/** The console cookie is only ever needed by /api routes. */
export const CONSOLE_COOKIE_PATH = '/api';

/** Idle timeout. /api/auth/me pushes expiry out by this much on each call. */
export const CONSOLE_IDLE_MS = 30 * 60 * 1000;

/** Hard cap from the sign-in that minted the token. Renewal never passes it. */
export const CONSOLE_MAX_LIFETIME_MS = 8 * 60 * 60 * 1000;

const MAC_PURPOSE = 'console-session-v1';

export interface ConsoleClaims {
  /** Unix ms of the sign-in that minted the token. Survives renewal. */
  issuedAt: number;
  /** Unix ms after which the token is refused. */
  expiresAt: number;
}

function mac(sessionToken: string, issuedAt: number, expiresAt: number): string {
  return macWithSessionKey(MAC_PURPOSE, `${sessionToken}|${issuedAt}|${expiresAt}`);
}

function expiryFor(issuedAt: number, now: number): number {
  return Math.min(now + CONSOLE_IDLE_MS, issuedAt + CONSOLE_MAX_LIFETIME_MS);
}

export interface ConsoleToken extends ConsoleClaims {
  /** The cookie value. */
  value: string;
}

/** Mint a console token bound to `sessionToken`, at sign-in. */
export function mintConsoleToken(sessionToken: string, now = Date.now()): ConsoleToken {
  return encode(sessionToken, { issuedAt: now, expiresAt: expiryFor(now, now) });
}

/** Re-issue a verified token with its idle expiry pushed out; issuedAt is kept. */
export function renewConsoleToken(sessionToken: string, claims: ConsoleClaims, now = Date.now()): ConsoleToken {
  return encode(sessionToken, { issuedAt: claims.issuedAt, expiresAt: expiryFor(claims.issuedAt, now) });
}

function encode(sessionToken: string, c: ConsoleClaims): ConsoleToken {
  return { ...c, value: `${c.issuedAt}.${c.expiresAt}.${mac(sessionToken, c.issuedAt, c.expiresAt)}` };
}

/**
 * Verify a console token against the session token it must be bound to.
 * Returns its claims, or null when it is malformed, forged, bound to a
 * different session, expired, or claims a lifetime longer than the cap.
 */
export function verifyConsoleToken(
  value: string | null | undefined,
  sessionToken: string | null | undefined,
  now = Date.now(),
): ConsoleClaims | null {
  if (!value || !sessionToken) return null;
  const m = /^(\d{1,16})\.(\d{1,16})\.([0-9a-f]{64})$/.exec(value);
  if (!m) return null;
  const issuedAt = Number(m[1]);
  const expiresAt = Number(m[2]);

  const expected = Buffer.from(mac(sessionToken, issuedAt, expiresAt), 'hex');
  const given = Buffer.from(m[3], 'hex');
  if (expected.length !== given.length || !timingSafeEqual(expected, given)) return null;

  if (now >= expiresAt) return null;
  if (expiresAt - issuedAt > CONSOLE_MAX_LIFETIME_MS) return null;
  if (now - issuedAt >= CONSOLE_MAX_LIFETIME_MS) return null;
  return { issuedAt, expiresAt };
}

/** Set-Cookie attributes for the console cookie. maxAge in seconds. */
export function consoleCookie(token: Pick<ConsoleToken, 'value' | 'expiresAt'>, now = Date.now()) {
  return {
    name: CONSOLE_COOKIE,
    value: token.value,
    httpOnly: true,
    secure: true,
    sameSite: 'Strict' as const,
    path: CONSOLE_COOKIE_PATH,
    maxAge: Math.max(0, Math.floor((token.expiresAt - now) / 1000)),
  };
}

/** Set-Cookie that clears the console cookie. */
export function expiredConsoleCookie() {
  return consoleCookie({ value: '', expiresAt: 0 }, 0);
}

/** Read one cookie by name from the request's Cookie header. */
export function readCookie(request: HttpRequest, name: string): string | null {
  const header = request.headers.get('cookie') ?? '';
  for (const part of header.split(';')) {
    const eq = part.indexOf('=');
    if (eq < 0) continue;
    if (part.slice(0, eq).trim() === name) return part.slice(eq + 1).trim();
  }
  return null;
}

// ── Origin check ────────────────────────────────────────────────────────────

function firstHeaderValue(request: HttpRequest, name: string): string | null {
  const raw = request.headers.get(name);
  if (!raw) return null;
  const first = raw.split(',')[0].trim();
  return first || null;
}

function originOf(url: string): string | null {
  try {
    const u = new URL(url);
    if (u.protocol !== 'https:' && u.protocol !== 'http:') return null;
    return u.origin.toLowerCase();
  } catch {
    return null;
  }
}

/**
 * The origins a legitimate admin SPA request can come from: this instance's
 * own origin as the browser addressed it, plus FRONTEND_URL's origin when that
 * is an absolute URL (local development serves the SPA from the Vite dev
 * server, whose proxy rewrites Host).
 */
export function allowedOrigins(request: HttpRequest, env: NodeJS.ProcessEnv = process.env): Set<string> {
  const out = new Set<string>();
  const proto = firstHeaderValue(request, 'x-forwarded-proto') ?? 'https';
  const host = firstHeaderValue(request, 'x-forwarded-host') ?? firstHeaderValue(request, 'host');
  if (host) {
    const own = originOf(`${proto}://${host}`);
    if (own) out.add(own);
  }
  const frontend = env.FRONTEND_URL?.trim();
  if (frontend && /^https?:\/\//i.test(frontend)) {
    const fe = originOf(frontend);
    if (fe) out.add(fe);
  }
  return out;
}

export type OriginCheck = { ok: true } | { ok: false; reason: string };

/**
 * Accept only a request a browser made from this instance's own pages.
 *
 * - An `Origin` header, when present, must be one of allowedOrigins. That
 *   includes `null`, which is refused.
 * - `Sec-Fetch-Site`, when present, must be `same-origin`. A same-site
 *   sibling subdomain is not enough: SameSite cookies would reach it, which
 *   is why the cookie attribute alone is not the CSRF defence. A GET may also
 *   be `none`, a navigation the user typed or bookmarked.
 * - A state-changing method must carry a matching `Origin`. Browsers always
 *   send one on a non-GET request.
 * - A GET must carry at least one of the two, so a request with neither, which
 *   no current browser sends, is refused.
 *
 * Non-browser clients can set these headers to anything, so this is not
 * authentication. It stops a page on another origin from riding the browser's
 * cookies; the console cookie is what keeps non-browser clients out.
 */
export function checkBrowserOrigin(request: HttpRequest, env: NodeJS.ProcessEnv = process.env): OriginCheck {
  const method = (request.method ?? 'GET').toUpperCase();
  const safe = method === 'GET' || method === 'HEAD';
  const origin = request.headers.get('origin') ?? null;
  const site = request.headers.get('sec-fetch-site')?.toLowerCase() ?? null;

  if (origin !== null) {
    if (!allowedOrigins(request, env).has(origin.trim().toLowerCase())) {
      return { ok: false, reason: 'origin not allowed' };
    }
  }
  if (site !== null && site !== 'same-origin' && !(safe && site === 'none')) {
    return { ok: false, reason: `sec-fetch-site ${site}` };
  }
  if (!safe && origin === null) {
    return { ok: false, reason: 'missing Origin on a state-changing request' };
  }
  if (safe && origin === null && site === null) {
    return { ok: false, reason: 'neither Origin nor Sec-Fetch-Site present' };
  }
  return { ok: true };
}
