/**
 * The mcp_session cookie authenticates only a request from this instance's own
 * pages (threat model row 3.4, gap G15).
 *
 * SameSite=Lax keeps the cookie off cross-site POSTs, but a page on a sibling
 * subdomain under the same registrable domain is same-site, so the browser
 * attaches the cookie to its requests. authenticateRequest and
 * authenticateRequestAllowExpired therefore take the cookie only past the
 * Origin check, which covers /api/mail/settings, the tool routes, /api/mcp and
 * /api/auth/refresh in one place. Bearer and x-session-token are unaffected:
 * a page on another origin cannot set them without a CORS preflight.
 */

import { jest } from '@jest/globals';
import type { HttpRequest } from '@azure/functions';
import type { UserSession } from '../services/tokenCache.js';

const mockGetSessionByToken = jest.fn<(token: string) => Promise<UserSession | undefined>>();

jest.mock('../services/tokenCache.js', () => ({
  getSessionByToken: (...args: unknown[]) => mockGetSessionByToken(args[0] as string),
  getValidAccessToken: jest.fn(),
  SESSION_TTL_MS: 7 * 24 * 60 * 60 * 1000,
  storeSession: jest.fn(async () => undefined),
  isAbsoluteLifetimeExceeded: () => false,
}));

jest.mock('../services/graphClient.js', () => ({
  acquireTokenSilent: jest.fn(),
  createGraphClient: jest.fn(),
}));

import { authenticateRequest, authenticateRequestAllowExpired } from '../services/authMiddleware.js';

const HOST = 'mcp.example.com';
const SELF = `https://${HOST}`;
const TOKEN = 'a'.repeat(64);

const SESSION: UserSession = {
  userId: 'user-1',
  homeAccountId: 'user-1.tenant-1',
  displayName: 'Adele Vance',
  email: 'adele@fabrikam.com',
  tenantId: 'tenant-1',
  accessToken: 'graph-token',
  expiresAt: Date.now() + 3_600_000,
  sessionToken: TOKEN,
  sessionCreatedAt: Date.now(),
  sessionAbsoluteCreatedAt: Date.now(),
  kind: 'browser',
};

function req(method: string, headers: Record<string, string>): HttpRequest {
  return { method, headers: new Map(Object.entries({ host: HOST, ...headers })) } as unknown as HttpRequest;
}

const COOKIE = `mcp_session=${TOKEN}`;

/** Requests a browser sends when a page on another origin rides the cookie. */
const FOREIGN: Array<[string, string, Record<string, string>]> = [
  ['sibling subdomain POST', 'POST', { cookie: COOKIE, origin: 'https://evil.example.com', 'sec-fetch-site': 'same-site' }],
  ['sibling subdomain PUT', 'PUT', { cookie: COOKIE, origin: 'https://evil.example.com', 'sec-fetch-site': 'same-site' }],
  ['sibling subdomain GET', 'GET', { cookie: COOKIE, 'sec-fetch-site': 'same-site' }],
  ['cross-site POST', 'POST', { cookie: COOKIE, origin: 'https://attacker.test', 'sec-fetch-site': 'cross-site' }],
  ['opaque origin POST', 'POST', { cookie: COOKIE, origin: 'null' }],
  ['write with no Origin', 'POST', { cookie: COOKIE, 'sec-fetch-site': 'same-origin' }],
  ['GET with neither header', 'GET', { cookie: COOKIE }],
];

beforeEach(() => {
  jest.clearAllMocks();
  delete process.env.FRONTEND_URL;
  mockGetSessionByToken.mockImplementation(async (t) => (t === TOKEN ? { ...SESSION } : undefined));
});

describe.each([
  ['authenticateRequest', authenticateRequest],
  ['authenticateRequestAllowExpired', authenticateRequestAllowExpired],
])('%s and the mcp_session cookie', (_name, authenticate) => {
  it('accepts a same-origin fetch from the SPA, GET and write', async () => {
    expect((await authenticate(req('GET', { cookie: COOKIE, 'sec-fetch-site': 'same-origin' })))?.userId).toBe('user-1');
    expect((await authenticate(req('PUT', { cookie: COOKIE, origin: SELF, 'sec-fetch-site': 'same-origin' })))?.userId).toBe('user-1');
  });

  it.each(FOREIGN)('refuses a %s, without looking the token up', async (_label, method, headers) => {
    expect(await authenticate(req(method, headers))).toBeNull();
    expect(mockGetSessionByToken).not.toHaveBeenCalled();
  });

  it('still accepts the token as a bearer or x-session-token, which a foreign page cannot set', async () => {
    expect((await authenticate(req('POST', { authorization: `Bearer ${TOKEN}` })))?.userId).toBe('user-1');
    expect((await authenticate(req('POST', { 'x-session-token': TOKEN })))?.userId).toBe('user-1');
  });

  it('a header token still authenticates when a refused cookie rides along', async () => {
    const auth = await authenticate(req('POST', {
      authorization: `Bearer ${TOKEN}`,
      cookie: `mcp_session=${'b'.repeat(64)}`,
      origin: 'https://evil.example.com',
      'sec-fetch-site': 'same-site',
    }));
    expect(auth?.userId).toBe('user-1');
    expect(mockGetSessionByToken).toHaveBeenCalledTimes(1);
    expect(mockGetSessionByToken).toHaveBeenCalledWith(TOKEN);
  });

  it('an unknown bearer does not open the door to a cookie from another origin', async () => {
    const auth = await authenticate(req('POST', {
      authorization: 'Bearer graph-access-token',
      cookie: COOKIE,
      origin: 'https://evil.example.com',
      'sec-fetch-site': 'same-site',
    }));
    expect(auth).toBeNull();
    expect(mockGetSessionByToken).not.toHaveBeenCalledWith(TOKEN);
  });
});
