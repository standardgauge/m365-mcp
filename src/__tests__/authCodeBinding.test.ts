/**
 * The authorization-code flow binds each redemption to the browser that
 * started it with PKCE (S256) and an OpenID Connect nonce, and login does not
 * return exception text (threat model section 1, rows 1.2 and 1.7).
 *
 * MSAL is mocked out here; msalAuthFlows.test.ts runs the same values through
 * the real library against FakeEntra, which enforces the verifier and echoes
 * the nonce. This file covers the handlers: what login puts in cookies and the
 * authorize URL, what the callback reads back and refuses without.
 */

import { createHash } from 'crypto';
import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

type UrlBinding = { codeChallenge: string; nonce: string };
type RedemptionBinding = { codeVerifier: string; nonce: string };

const mockGetAuthCodeUrl = jest.fn<(state: string, binding: UrlBinding) => Promise<string>>();
const mockAcquireTokenByCode = jest.fn<(code: string, binding: RedemptionBinding) => Promise<unknown>>();

jest.mock('../services/graphClient.js', () => ({
  getAuthCodeUrl: (state: string, binding: UrlBinding) => mockGetAuthCodeUrl(state, binding),
  acquireTokenByCode: (code: string, binding: RedemptionBinding) => mockAcquireTokenByCode(code, binding),
  createGraphClient: () => ({
    api: () => ({
      select: () => ({
        get: async () => ({
          id: 'user-id',
          displayName: 'Adele Vance',
          mail: 'adele.vance@fabrikam.com',
          userPrincipalName: 'adele.vance@fabrikam.com',
        }),
      }),
    }),
  }),
}));

jest.mock('../services/tokenCache.js', () => ({
  storeSession: jest.fn(async () => undefined),
  SESSION_TTL_MS: 7 * 24 * 60 * 60 * 1000,
}));
jest.mock('../services/tableStorage.js', () => ({
  attachSessionToInstallNonce: jest.fn(async () => true),
}));
jest.mock('../services/securityHeaders.js', () => ({
  withSecurity: (handler: unknown) => handler,
}));
jest.mock('../services/rateLimit.js', () => ({
  withRateLimit: (_name: string, handler: unknown) => handler,
}));
jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/auth/login.js';
import '../functions/auth/callback.js';

interface Cookie {
  name: string;
  value: string;
  httpOnly?: boolean;
  secure?: boolean;
  maxAge?: number;
}
type Handler = (
  req: HttpRequest,
  ctx: InvocationContext
) => Promise<{ status: number; jsonBody?: unknown; cookies?: Cookie[] }>;

function handlerFor(name: string): Handler {
  const calls = (app.http as unknown as jest.Mock).mock.calls as Array<[string, { handler: Handler }]>;
  const reg = calls.find(([n]) => n === name);
  if (!reg) throw new Error(`${name} handler was not registered`);
  return reg[1].handler;
}

const login = handlerFor('login');
const callback = handlerFor('callback');
const ctx = { error: jest.fn(), warn: jest.fn(), log: jest.fn() } as unknown as InvocationContext;

function loginRequest(): HttpRequest {
  return {
    query: new Map<string, string>(),
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

function callbackRequest(cookie: string): HttpRequest {
  return {
    query: new Map<string, string>([
      ['code', 'auth-code-123'],
      ['state', 'state-xyz'],
    ]),
    headers: new Map<string, string>([['cookie', cookie]]),
  } as unknown as HttpRequest;
}

function cookie(cookies: Cookie[] | undefined, name: string): Cookie | undefined {
  return cookies?.find((c) => c.name === name);
}

const ORIGINAL_TENANT_ID = process.env.AZURE_TENANT_ID;

beforeEach(() => {
  jest.clearAllMocks();
  delete process.env.AZURE_TENANT_ID;
  mockGetAuthCodeUrl.mockResolvedValue('https://login.microsoftonline.com/authorize');
  mockAcquireTokenByCode.mockResolvedValue({
    accessToken: 'graph-access-token',
    homeAccountId: 'user-id.tenant-id',
    expiresOn: new Date(Date.now() + 3_600_000),
  });
});

afterAll(() => {
  if (ORIGINAL_TENANT_ID === undefined) delete process.env.AZURE_TENANT_ID;
  else process.env.AZURE_TENANT_ID = ORIGINAL_TENANT_ID;
});

describe('login', () => {
  it('sends the S256 challenge of the verifier it keeps in an HttpOnly cookie', async () => {
    const res = await login(loginRequest(), ctx);

    expect(res.status).toBe(302);
    const verifier = cookie(res.cookies, 'oauth_pkce');
    expect(verifier).toMatchObject({ httpOnly: true, secure: true, maxAge: 600 });
    // RFC 7636 section 4.1: 43 to 128 characters from the unreserved set.
    expect(verifier!.value).toMatch(/^[A-Za-z0-9\-._~]{43,128}$/);

    const [, binding] = mockGetAuthCodeUrl.mock.calls[0];
    expect(binding.codeChallenge).toBe(
      createHash('sha256').update(verifier!.value).digest('base64url')
    );
  });

  it('sends the nonce it keeps in an HttpOnly cookie', async () => {
    const res = await login(loginRequest(), ctx);

    const nonce = cookie(res.cookies, 'oauth_nonce');
    expect(nonce).toMatchObject({ httpOnly: true, secure: true, maxAge: 600 });
    expect(nonce!.value).toMatch(/^[a-f0-9]{32}$/);
    expect(mockGetAuthCodeUrl.mock.calls[0][1].nonce).toBe(nonce!.value);
  });

  it('uses a fresh verifier and nonce on every login', async () => {
    const a = await login(loginRequest(), ctx);
    const b = await login(loginRequest(), ctx);

    expect(cookie(a.cookies, 'oauth_pkce')!.value).not.toBe(cookie(b.cookies, 'oauth_pkce')!.value);
    expect(cookie(a.cookies, 'oauth_nonce')!.value).not.toBe(cookie(b.cookies, 'oauth_nonce')!.value);
  });

  it('returns a generic error without the exception text', async () => {
    mockGetAuthCodeUrl.mockRejectedValue(
      new Error('endpoints_resolution_error: tenant 72f988bf not found at authority')
    );

    const res = await login(loginRequest(), ctx);

    expect(res.status).toBe(500);
    expect(res.jsonBody).toEqual({ error: 'Failed to initiate login flow' });
    expect(JSON.stringify(res.jsonBody)).not.toContain('72f988bf');
  });
});

describe('callback', () => {
  const STATE_COOKIE = 'oauth_state=state-xyz';

  it('redeems the code with the verifier and nonce from the login cookies', async () => {
    const res = await callback(
      callbackRequest(`${STATE_COOKIE}; oauth_pkce=verifier-abc; oauth_nonce=nonce-abc`),
      ctx
    );

    expect(res.status).toBe(302);
    expect(mockAcquireTokenByCode).toHaveBeenCalledWith('auth-code-123', {
      codeVerifier: 'verifier-abc',
      nonce: 'nonce-abc',
    });
  });

  it('clears the verifier and nonce cookies once the code is redeemed', async () => {
    const res = await callback(
      callbackRequest(`${STATE_COOKIE}; oauth_pkce=verifier-abc; oauth_nonce=nonce-abc`),
      ctx
    );

    expect(cookie(res.cookies, 'oauth_pkce')).toMatchObject({ value: '', maxAge: 0, httpOnly: true });
    expect(cookie(res.cookies, 'oauth_nonce')).toMatchObject({ value: '', maxAge: 0, httpOnly: true });
  });

  it.each([
    ['the verifier', `${STATE_COOKIE}; oauth_nonce=nonce-abc`],
    ['the nonce', `${STATE_COOKIE}; oauth_pkce=verifier-abc`],
    ['both', STATE_COOKIE],
  ])('refuses without redeeming when %s cookie is missing', async (_what, cookieHeader) => {
    const res = await callback(callbackRequest(cookieHeader), ctx);

    expect(res.status).toBe(403);
    expect(mockAcquireTokenByCode).not.toHaveBeenCalled();
  });

  it('returns a generic error when redemption fails, such as on a nonce mismatch', async () => {
    mockAcquireTokenByCode.mockRejectedValue(new Error('nonce_mismatch: Nonce mismatch error.'));

    const res = await callback(
      callbackRequest(`${STATE_COOKIE}; oauth_pkce=verifier-abc; oauth_nonce=nonce-abc`),
      ctx
    );

    expect(res.status).toBe(500);
    expect(res.jsonBody).toEqual({ error: 'Authentication callback failed' });
  });
});
