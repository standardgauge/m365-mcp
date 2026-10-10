/**
 * Tests for the post-OAuth redirect target of /api/auth/callback.
 *
 * The callback ends the login flow with a 302. Where it points is load-bearing
 * and was wrong in production on all three tenants: it redirected to
 * `<host>/admin`, and the Azure Functions host reserves /admin/* for its own
 * key-protected admin API, so the response never reached the serveAdmin
 * catch-all and every login terminated on an empty 404.
 *
 * The admin SPA is served from the site root. Nothing in the app reads a
 * userId query parameter off this redirect — the SPA resolves identity from
 * /api/auth/me against the mcp_session cookie — so the GUID is not appended.
 *
 * Covers:
 *   - FRONTEND_URL unset → redirects to the site root
 *   - FRONTEND_URL set   → redirects to exactly that value
 *   - the redirect never targets a host-reserved /admin path
 *   - the redirect carries no userId query parameter
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import { randomBytes } from 'crypto';
import { verifyConsoleToken } from '../services/consoleSession.js';

// The callback mints a console session, which is MAC'd under the session key.
process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');

const mockAcquireTokenByCode = jest.fn<() => Promise<unknown>>();
const mockCreateGraphClient = jest.fn<() => unknown>();
const mockStoreSession = jest.fn<() => Promise<void>>();

jest.mock('../services/graphClient.js', () => ({
  acquireTokenByCode: () => mockAcquireTokenByCode(),
  createGraphClient: () => mockCreateGraphClient(),
}));

jest.mock('../services/tokenCache.js', () => ({
  storeSession: () => mockStoreSession(),
  SESSION_TTL_MS: 7 * 24 * 60 * 60 * 1000,
}));

jest.mock('../services/tenantUtils.js', () => ({
  extractTenantId: () => TENANT,
}));

jest.mock('../services/tableStorage.js', () => ({
  attachSessionToInstallNonce: jest.fn(async () => true),
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/auth/callback.js';

interface Cookie {
  name: string;
  value: string;
  httpOnly?: boolean;
  secure?: boolean;
  sameSite?: string;
  path?: string;
  maxAge?: number;
}
interface HttpRegistration {
  handler: (
    req: HttpRequest,
    context: InvocationContext
  ) => Promise<{ status: number; headers?: Record<string, string>; jsonBody?: unknown; cookies?: Cookie[] }>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const reg = httpMock.mock.calls.find((c) => c[0] === 'callback');
if (!reg) throw new Error('callback handler was not registered');
const handler = reg[1].handler;

const TENANT = 'tenant-abc';
const USER_ID = 'a5361c9c-ed96-44d9-a007-67df1e77ee62';
const STATE = 'state-token-xyz';

const ctx = { error: jest.fn(), warn: jest.fn() } as unknown as InvocationContext;

/** A callback request that passes state validation and carries an auth code. */
function req(): HttpRequest {
  return {
    method: 'GET',
    query: new Map<string, string>([
      ['code', 'auth-code-123'],
      ['state', STATE],
    ]),
    headers: new Map<string, string>([
      ['cookie', `oauth_state=${STATE}; oauth_pkce=verifier-abc; oauth_nonce=nonce-abc`],
    ]),
  } as unknown as HttpRequest;
}

const ORIGINAL_FRONTEND_URL = process.env.FRONTEND_URL;
const ORIGINAL_TENANT_ID = process.env.AZURE_TENANT_ID;

beforeEach(() => {
  jest.clearAllMocks();
  delete process.env.AZURE_TENANT_ID; // skip the foreign-tenant rejection branch
  mockAcquireTokenByCode.mockResolvedValue({
    accessToken: 'graph-access-token',
    homeAccountId: `${USER_ID}.${TENANT}`,
    expiresOn: new Date(Date.now() + 3_600_000),
  });
  mockCreateGraphClient.mockReturnValue({
    api: () => ({
      select: () => ({
        get: async () => ({
          id: USER_ID,
          displayName: 'Test User',
          mail: 'nate@example.com',
          userPrincipalName: 'nate@example.com',
        }),
      }),
    }),
  });
  mockStoreSession.mockResolvedValue(undefined);
});

afterAll(() => {
  if (ORIGINAL_FRONTEND_URL === undefined) delete process.env.FRONTEND_URL;
  else process.env.FRONTEND_URL = ORIGINAL_FRONTEND_URL;
  if (ORIGINAL_TENANT_ID === undefined) delete process.env.AZURE_TENANT_ID;
  else process.env.AZURE_TENANT_ID = ORIGINAL_TENANT_ID;
});

describe('callback — post-OAuth redirect target', () => {
  it('redirects to the site root when FRONTEND_URL is unset', async () => {
    delete process.env.FRONTEND_URL;
    const res = await handler(req(), ctx);
    expect(res.status).toBe(302);
    expect(res.headers?.Location).toBe('/');
  });

  it('redirects to FRONTEND_URL verbatim when it is set', async () => {
    process.env.FRONTEND_URL = 'https://mcp.example.com/';
    const res = await handler(req(), ctx);
    expect(res.status).toBe(302);
    expect(res.headers?.Location).toBe('https://mcp.example.com/');
  });

  it('never defaults to /admin, which the Functions host reserves', async () => {
    delete process.env.FRONTEND_URL;
    const res = await handler(req(), ctx);
    const location = res.headers?.Location ?? '';
    expect(location).not.toMatch(/\/admin(\/|$|\?)/);
  });

  it('does not leak the user GUID into the redirect URL', async () => {
    delete process.env.FRONTEND_URL;
    const res = await handler(req(), ctx);
    const location = res.headers?.Location ?? '';
    expect(location).not.toContain('userId=');
    expect(location).not.toContain(USER_ID);
  });
});

describe('callback — console session', () => {
  it('sets an HttpOnly, SameSite=Strict mcp_console cookie bound to the new session token', async () => {
    delete process.env.FRONTEND_URL;
    const res = await handler(req(), ctx);
    const session = res.cookies?.find((c) => c.name === 'mcp_session');
    const consoleCookie = res.cookies?.find((c) => c.name === 'mcp_console');
    expect(session?.value).toMatch(/^[0-9a-f]{64}$/);
    expect(consoleCookie).toMatchObject({ httpOnly: true, secure: true, sameSite: 'Strict', path: '/api' });
    expect(consoleCookie?.maxAge).toBeGreaterThan(0);
    expect(consoleCookie?.maxAge).toBeLessThanOrEqual(30 * 60);
    expect(verifyConsoleToken(consoleCookie?.value, session?.value)).not.toBeNull();
    // Bound to this session: useless beside any other session token.
    expect(verifyConsoleToken(consoleCookie?.value, 'f'.repeat(64))).toBeNull();
  });
});
