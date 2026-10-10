/**
 * The console session: the browser-only credential /api/manage/* requires
 * (threat model section 7, rows 7.3 and 7.4, and row 3.4).
 *
 * Covers:
 *   - console token: bound to one session token, MAC'd, idle expiry, 8-hour cap
 *     that renewal never passes, tampering refused
 *   - Origin check: cross-origin, same-site sibling, null origin, a write with
 *     no Origin, and a GET with neither header are all refused
 *   - authenticateConsoleRequest: the MCP client's session token is not an
 *     admin credential, whether presented as bearer, x-session-token or a
 *     forged mcp_session cookie, without the console cookie bound to it
 *   - every /api/manage/* route authenticates through authenticateConsoleRequest
 *   - logout is POST only, refuses a cross-origin request, clears mcp_console,
 *     and never sends Microsoft back to a host-reserved path
 *   - the admin SPA has no inline executable script and a CSP that refuses one
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import { randomBytes } from 'crypto';
import * as fs from 'fs';
import * as path from 'path';
import type { UserSession } from '../services/tokenCache.js';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

const mockGetSessionByToken = jest.fn<(token: string) => Promise<UserSession | undefined>>();
const mockDeleteAllUserSessions = jest.fn<(userId: string) => Promise<void>>();

jest.mock('../services/tokenCache.js', () => ({
  getSessionByToken: (...args: unknown[]) => mockGetSessionByToken(args[0] as string),
  deleteAllUserSessions: (...args: unknown[]) => mockDeleteAllUserSessions(args[0] as string),
  getValidAccessToken: jest.fn(),
  SESSION_TTL_MS: 7 * 24 * 60 * 60 * 1000,
  storeSession: jest.fn(async () => undefined),
  isAbsoluteLifetimeExceeded: () => false,
}));

jest.mock('../services/graphClient.js', () => ({
  acquireTokenSilent: jest.fn(),
  createGraphClient: jest.fn(),
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import { authenticateConsoleRequest } from '../services/authMiddleware.js';
import {
  CONSOLE_IDLE_MS,
  CONSOLE_MAX_LIFETIME_MS,
  checkBrowserOrigin,
  mintConsoleToken,
  renewConsoleToken,
  verifyConsoleToken,
} from '../services/consoleSession.js';
import { adminSpaHeaders } from '../services/securityHeaders.js';
import { reservedPrefixFor } from '../services/frontendUrl.js';
import '../functions/auth/logout.js';
import { injectRuntimeConfig } from '../functions/admin/serveAdmin.js';

const HOST = 'mcp.example.com';
const SELF = `https://${HOST}`;
const SESSION_TOKEN = 'a'.repeat(64);
const OTHER_TOKEN = 'b'.repeat(64);

const SESSION: UserSession = {
  userId: 'user-1',
  homeAccountId: 'user-1.tenant-1',
  displayName: 'Admin User',
  email: 'admin@example.com',
  tenantId: 'tenant-1',
  accessToken: 'graph-token',
  expiresAt: Date.now() + 3_600_000,
  sessionToken: SESSION_TOKEN,
  sessionCreatedAt: Date.now(),
  sessionAbsoluteCreatedAt: Date.now(),
};

function req(method: string, headers: Record<string, string>): HttpRequest {
  const map = new Map<string, string>(Object.entries({ host: HOST, ...headers }).map(([k, v]) => [k.toLowerCase(), v]));
  return { method, headers: map } as unknown as HttpRequest;
}

/** A same-origin fetch from the admin SPA. */
function browserHeaders(cookie: string, method = 'GET'): Record<string, string> {
  const h: Record<string, string> = { cookie, 'sec-fetch-site': 'same-origin' };
  if (method !== 'GET') h.origin = SELF;
  return h;
}

beforeEach(() => {
  jest.clearAllMocks();
  delete process.env.FRONTEND_URL;
  mockGetSessionByToken.mockImplementation(async (t) => (t === SESSION_TOKEN ? { ...SESSION } : undefined));
});

describe('console token', () => {
  const T0 = 1_800_000_000_000;

  it('verifies against the session token it was minted for, and no other', () => {
    const tok = mintConsoleToken(SESSION_TOKEN, T0);
    expect(verifyConsoleToken(tok.value, SESSION_TOKEN, T0 + 1000)).toEqual({ issuedAt: T0, expiresAt: T0 + CONSOLE_IDLE_MS });
    expect(verifyConsoleToken(tok.value, OTHER_TOKEN, T0 + 1000)).toBeNull();
    expect(verifyConsoleToken(tok.value, null, T0 + 1000)).toBeNull();
  });

  it('expires after the idle window', () => {
    const tok = mintConsoleToken(SESSION_TOKEN, T0);
    expect(verifyConsoleToken(tok.value, SESSION_TOKEN, T0 + CONSOLE_IDLE_MS)).toBeNull();
  });

  it('refuses a token whose expiry was edited', () => {
    const tok = mintConsoleToken(SESSION_TOKEN, T0);
    const [iat, , sig] = tok.value.split('.');
    const forged = `${iat}.${T0 + CONSOLE_MAX_LIFETIME_MS}.${sig}`;
    expect(verifyConsoleToken(forged, SESSION_TOKEN, T0 + CONSOLE_IDLE_MS + 1)).toBeNull();
  });

  it('refuses malformed values', () => {
    for (const v of ['', 'x', '1.2', `1.2.${'g'.repeat(64)}`, `${T0}.${T0 + 1}.${'0'.repeat(63)}`]) {
      expect(verifyConsoleToken(v, SESSION_TOKEN, T0)).toBeNull();
    }
  });

  it('renewal pushes the idle expiry out but never past the 8-hour cap', () => {
    const tok = mintConsoleToken(SESSION_TOKEN, T0);
    const later = T0 + CONSOLE_IDLE_MS - 1000;
    const renewed = renewConsoleToken(SESSION_TOKEN, tok, later);
    expect(renewed.issuedAt).toBe(T0);
    expect(renewed.expiresAt).toBe(later + CONSOLE_IDLE_MS);

    const nearCap = T0 + CONSOLE_MAX_LIFETIME_MS - 60_000;
    const last = renewConsoleToken(SESSION_TOKEN, tok, nearCap);
    expect(last.expiresAt).toBe(T0 + CONSOLE_MAX_LIFETIME_MS);
    expect(verifyConsoleToken(last.value, SESSION_TOKEN, T0 + CONSOLE_MAX_LIFETIME_MS)).toBeNull();
  });
});

describe('checkBrowserOrigin', () => {
  it('accepts a same-origin fetch, GET and POST', () => {
    expect(checkBrowserOrigin(req('GET', { 'sec-fetch-site': 'same-origin' })).ok).toBe(true);
    expect(checkBrowserOrigin(req('POST', { origin: SELF, 'sec-fetch-site': 'same-origin' })).ok).toBe(true);
    // Older browsers without Fetch Metadata still send Origin on writes.
    expect(checkBrowserOrigin(req('DELETE', { origin: SELF })).ok).toBe(true);
  });

  it('accepts a typed or bookmarked GET navigation', () => {
    expect(checkBrowserOrigin(req('GET', { 'sec-fetch-site': 'none' })).ok).toBe(true);
    expect(checkBrowserOrigin(req('POST', { origin: SELF, 'sec-fetch-site': 'none' })).ok).toBe(false);
  });

  it('refuses another origin, including a same-site sibling subdomain', () => {
    expect(checkBrowserOrigin(req('POST', { origin: 'https://evil.example.net', 'sec-fetch-site': 'cross-site' })).ok).toBe(false);
    expect(checkBrowserOrigin(req('POST', { origin: 'https://other.example.com', 'sec-fetch-site': 'same-site' })).ok).toBe(false);
    expect(checkBrowserOrigin(req('GET', { 'sec-fetch-site': 'same-site' })).ok).toBe(false);
    expect(checkBrowserOrigin(req('POST', { origin: 'null' })).ok).toBe(false);
  });

  it('refuses a write with no Origin, and a GET with neither header', () => {
    expect(checkBrowserOrigin(req('POST', { 'sec-fetch-site': 'same-origin' })).ok).toBe(false);
    expect(checkBrowserOrigin(req('GET', {})).ok).toBe(false);
  });

  it('honours the forwarded host and FRONTEND_URL', () => {
    expect(checkBrowserOrigin(req('POST', { host: 'internal:80', 'x-forwarded-host': HOST, 'x-forwarded-proto': 'https', origin: SELF })).ok).toBe(true);
    process.env.FRONTEND_URL = 'http://localhost:5173';
    expect(checkBrowserOrigin(req('POST', { host: '127.0.0.1:7071', origin: 'http://localhost:5173' })).ok).toBe(true);
  });
});

describe('authenticateConsoleRequest', () => {
  const consoleCookie = () => mintConsoleToken(SESSION_TOKEN).value;

  it('accepts the SPA: mcp_session plus a console cookie bound to it, same origin', async () => {
    const auth = await authenticateConsoleRequest(
      req('POST', browserHeaders(`mcp_session=${SESSION_TOKEN}; mcp_console=${consoleCookie()}`, 'POST')),
    );
    expect(auth?.userId).toBe('user-1');
    expect(auth?.sessionToken).toBe(SESSION_TOKEN);
  });

  it('refuses the MCP client token as a bearer or x-session-token header', async () => {
    expect(await authenticateConsoleRequest(req('GET', { authorization: `Bearer ${SESSION_TOKEN}`, 'sec-fetch-site': 'same-origin' }))).toBeNull();
    expect(await authenticateConsoleRequest(req('GET', { 'x-session-token': SESSION_TOKEN, 'sec-fetch-site': 'same-origin' }))).toBeNull();
    expect(mockGetSessionByToken).not.toHaveBeenCalled();
  });

  it('refuses the MCP client token replayed as a forged mcp_session cookie', async () => {
    const auth = await authenticateConsoleRequest(req('POST', browserHeaders(`mcp_session=${SESSION_TOKEN}`, 'POST')));
    expect(auth).toBeNull();
    expect(mockGetSessionByToken).not.toHaveBeenCalled();
  });

  it('refuses a console cookie next to a different session', async () => {
    mockGetSessionByToken.mockImplementation(async () => ({ ...SESSION }));
    const auth = await authenticateConsoleRequest(
      req('GET', browserHeaders(`mcp_session=${OTHER_TOKEN}; mcp_console=${consoleCookie()}`)),
    );
    expect(auth).toBeNull();
  });

  it('refuses a valid pair from another origin', async () => {
    const auth = await authenticateConsoleRequest(
      req('POST', {
        cookie: `mcp_session=${SESSION_TOKEN}; mcp_console=${consoleCookie()}`,
        origin: 'https://evil.example.net',
        'sec-fetch-site': 'cross-site',
      }),
    );
    expect(auth).toBeNull();
  });

  it('refuses once the session itself is gone (logout)', async () => {
    mockGetSessionByToken.mockResolvedValue(undefined);
    const auth = await authenticateConsoleRequest(
      req('GET', browserHeaders(`mcp_session=${SESSION_TOKEN}; mcp_console=${consoleCookie()}`)),
    );
    expect(auth).toBeNull();
  });
});

describe('every /api/manage/* route requires the console session', () => {
  const functionsDir = path.resolve(__dirname, '../functions');
  const files: string[] = [];
  const walk = (dir: string) => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, entry.name);
      if (entry.isDirectory()) walk(p);
      else if (p.endsWith('.ts')) files.push(p);
    }
  };
  walk(functionsDir);
  const manageFiles = files.filter((f) => /route:\s*'api\/manage\//.test(fs.readFileSync(f, 'utf-8')));

  it('finds the manage routes', () => {
    expect(manageFiles.length).toBeGreaterThanOrEqual(9);
  });

  it.each(manageFiles.map((f) => [path.relative(functionsDir, f), f]))('%s', (_rel, file) => {
    const src = fs.readFileSync(file, 'utf-8');
    expect(src).toMatch(/authenticateConsoleRequest\(/);
    expect(src).not.toMatch(/\bauthenticateRequest(AllowExpired)?\(/);
  });
});

describe('logout', () => {
  interface Reg {
    methods: string[];
    handler: (req: HttpRequest, ctx: InvocationContext) => Promise<{
      status: number;
      headers?: Record<string, string>;
      cookies?: Array<{ name: string; maxAge?: number; path?: string }>;
    }>;
  }
  const httpMock = app.http as unknown as jest.Mock<(name: string, opts: Reg) => void>;
  const reg = httpMock.mock.calls.find((c) => c[0] === 'authLogout')?.[1];
  if (!reg) throw new Error('authLogout handler was not registered');
  const handler = reg.handler;
  const ctx = { warn: jest.fn(), error: jest.fn() } as unknown as InvocationContext;

  it('is registered for POST only', () => {
    expect(reg.methods).toEqual(['POST']);
  });

  it('refuses a cross-origin POST and leaves the session alone', async () => {
    const res = await reg.handler(
      req('POST', { cookie: `mcp_session=${SESSION_TOKEN}`, origin: 'https://evil.example.net', 'sec-fetch-site': 'cross-site' }),
      ctx,
    );
    expect(res.status).toBe(403);
    expect(mockDeleteAllUserSessions).not.toHaveBeenCalled();
  });

  it('ends the session and clears both session cookies on a same-origin POST', async () => {
    const res = await reg.handler(req('POST', browserHeaders(`mcp_session=${SESSION_TOKEN}`, 'POST')), ctx);
    expect(res.status).toBe(303);
    expect(res.headers?.Location).toMatch(/^https:\/\/login\.microsoftonline\.com\//);
    expect(mockDeleteAllUserSessions).toHaveBeenCalledWith('user-1');
    const cleared = Object.fromEntries((res.cookies ?? []).map((c) => [c.name, c]));
    expect(cleared.mcp_session?.maxAge).toBe(0);
    expect(cleared.mcp_console).toMatchObject({ maxAge: 0, path: '/api' });
  });

  async function postLogoutTarget(frontendUrl: string | undefined): Promise<string> {
    const saved = process.env.FRONTEND_URL;
    if (frontendUrl === undefined) delete process.env.FRONTEND_URL;
    else process.env.FRONTEND_URL = frontendUrl;
    try {
      const res = await handler(req('POST', browserHeaders(`mcp_session=${SESSION_TOKEN}`, 'POST')), ctx);
      const target = new URL(res.headers?.Location ?? '').searchParams.get('post_logout_redirect_uri');
      if (!target) throw new Error('no post_logout_redirect_uri');
      return target;
    } finally {
      if (saved === undefined) delete process.env.FRONTEND_URL;
      else process.env.FRONTEND_URL = saved;
    }
  }

  it.each([undefined, '/', `${SELF}/`, `${SELF}/admin`, `${SELF}/admin/settings`, `${SELF}/runtime`, '/admin'])(
    'sends Microsoft back to the SPA root, never a host-reserved path (FRONTEND_URL=%s)',
    async (frontendUrl) => {
      const target = await postLogoutTarget(frontendUrl);
      expect(new URL(target).origin).toBe(SELF);
      expect(reservedPrefixFor(target)).toBeNull();
      expect(new URL(target).pathname).toBe('/');
    },
  );

  it('honours a non-reserved absolute FRONTEND_URL, such as the Vite dev server', async () => {
    expect(await postLogoutTarget('http://localhost:5173/')).toBe('http://localhost:5173/');
  });
});

describe('admin SPA inline script', () => {
  it('CSP script-src refuses inline script', () => {
    const csp = adminSpaHeaders['Content-Security-Policy'];
    const scriptSrc = csp.split(';').map((d) => d.trim()).find((d) => d.startsWith('script-src'));
    expect(scriptSrc).toBe("script-src 'self'");
    expect(csp).not.toContain("'unsafe-eval'");
  });

  it('injects runtime config as a non-executable JSON data block, escaped', () => {
    process.env.MCP_INSTANCE_NAME = 'Fabrikam </script><script>alert(1)</script>';
    try {
      const html = injectRuntimeConfig('<html><head></head><body></body></html>');
      const scripts = html.match(/<script\b[^>]*>/g) ?? [];
      expect(scripts).toEqual(['<script type="application/json" id="runtime-config">']);
      expect(html.match(/<\/script>/g)).toHaveLength(1);
      const json = html.slice(html.indexOf('>', html.indexOf('<script')) + 1, html.indexOf('</script>'));
      expect(JSON.parse(json)).toEqual({ instanceName: 'Fabrikam </script><script>alert(1)</script>' });
    } finally {
      delete process.env.MCP_INSTANCE_NAME;
    }
  });
});
