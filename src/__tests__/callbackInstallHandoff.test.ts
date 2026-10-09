/**
 * Tests for /api/auth/callback when an installer started the sign-in.
 *
 * The install_nonce cookie names a challenge, and whoever wrote the sign-in link
 * chose it. The callback must not hand the new session to that challenge on its
 * own; it records a pending handoff and sends the browser to install-confirm,
 * where the user has to enter the installer's code.
 *
 * Covers:
 *   - no session is attached to the install nonce by the callback
 *   - a pending handoff is recorded with the challenge and a keyed hash of the
 *     session token, never the token itself
 *   - the browser is redirected to install-confirm with the handoff cookie
 *   - an ordinary sign-in is unaffected
 *   - a failure to record the handoff degrades to an ordinary sign-in
 */

import { jest } from '@jest/globals';
import { createHash } from 'crypto';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { InstallHandoffRecord, InstallNonceRecord } from '../services/tableStorage.js';

const mockAcquireTokenByCode = jest.fn<() => Promise<unknown>>();
const mockCreateGraphClient = jest.fn<() => unknown>();

jest.mock('../services/graphClient.js', () => ({
  acquireTokenByCode: () => mockAcquireTokenByCode(),
  createGraphClient: () => mockCreateGraphClient(),
}));

jest.mock('../services/tokenCache.js', () => ({
  storeSession: async () => undefined,
  SESSION_TTL_MS: 7 * 24 * 60 * 60 * 1000,
}));

jest.mock('../services/tenantUtils.js', () => ({
  extractTenantId: () => 'tenant-abc',
}));

const attachMock = jest.fn<(nonce: string, rec: InstallNonceRecord) => Promise<boolean>>();
const createHandoffMock = jest.fn<(id: string, rec: InstallHandoffRecord) => Promise<void>>();
jest.mock('../services/tableStorage.js', () => ({
  attachSessionToInstallNonce: (nonce: string, rec: InstallNonceRecord) => attachMock(nonce, rec),
  createInstallHandoff: (id: string, rec: InstallHandoffRecord) => createHandoffMock(id, rec),
}));

const fakeHash = (t: string) => createHash('sha256').update('k:' + t).digest('hex');
jest.mock('../services/credentialCrypto.js', () => ({
  hashSessionToken: (t: string) => fakeHash(t),
}));

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/auth/callback.js';

interface Res {
  status: number;
  headers?: Record<string, string>;
  cookies?: Array<{ name: string; value: string; path?: string; httpOnly?: boolean; maxAge?: number }>;
}
const httpMock = app.http as unknown as jest.Mock<
  (name: string, opts: { handler: (r: HttpRequest, c: InvocationContext) => Promise<Res> }) => void
>;
const reg = httpMock.mock.calls.find((c) => c[0] === 'callback');
if (!reg) throw new Error('callback handler was not registered');
const handler = reg[1].handler;

const STATE = 'state-token-xyz';
const CHALLENGE = createHash('sha256').update('a'.repeat(32)).digest('hex');
const ctx = { error: jest.fn(), warn: jest.fn() } as unknown as InvocationContext;

function req(extraCookies = ''): HttpRequest {
  return {
    method: 'GET',
    query: new Map<string, string>([
      ['code', 'auth-code-123'],
      ['state', STATE],
    ]),
    headers: new Map<string, string>([['cookie', `oauth_state=${STATE}${extraCookies}`]]),
  } as unknown as HttpRequest;
}

beforeEach(() => {
  jest.clearAllMocks();
  delete process.env.AZURE_TENANT_ID;
  delete process.env.FRONTEND_URL;
  mockAcquireTokenByCode.mockResolvedValue({
    accessToken: 'graph-access-token',
    homeAccountId: 'user-1.tenant-abc',
    expiresOn: new Date(Date.now() + 3_600_000),
  });
  mockCreateGraphClient.mockReturnValue({
    api: () => ({
      select: () => ({
        get: async () => ({ id: 'user-1', displayName: 'Adele Vance', mail: 'adele@fabrikam.com' }),
      }),
    }),
  });
  createHandoffMock.mockResolvedValue(undefined);
});

describe('callback — installer sign-in', () => {
  it('does not attach the session to the install nonce', async () => {
    await handler(req(`; install_nonce=${CHALLENGE}; device_label=ADELE-LAPTOP`), ctx);
    expect(attachMock).not.toHaveBeenCalled();
  });

  it('records a pending handoff holding a hash of the session token, not the token', async () => {
    const res = await handler(req(`; install_nonce=${CHALLENGE}; device_label=ADELE-LAPTOP`), ctx);
    expect(createHandoffMock).toHaveBeenCalledTimes(1);
    const [id, rec] = createHandoffMock.mock.calls[0];
    expect(id).toMatch(/^[a-f0-9]{64}$/);
    expect(rec.challenge).toBe(CHALLENGE);
    expect(rec.userId).toBe('user-1');
    expect(rec.deviceLabel).toBe('ADELE-LAPTOP');
    expect(rec.attempts).toBe(0);
    const sessionToken = res.cookies?.find((c) => c.name === 'mcp_session')?.value ?? '';
    expect(sessionToken).toMatch(/^[a-f0-9]{64}$/);
    expect(rec.sessionTokenHash).toBe(fakeHash(sessionToken));
    expect(JSON.stringify(rec)).not.toContain(sessionToken);
  });

  it('sends the browser to install-confirm with the handoff id in an HttpOnly cookie', async () => {
    const res = await handler(req(`; install_nonce=${CHALLENGE}`), ctx);
    expect(res.status).toBe(302);
    expect(res.headers?.Location).toBe('/api/auth/install-confirm');
    const cookie = res.cookies?.find((c) => c.name === 'install_handoff');
    expect(cookie?.value).toBe(createHandoffMock.mock.calls[0][0]);
    expect(cookie?.httpOnly).toBe(true);
    expect(cookie?.path).toBe('/api/auth/install-confirm');
    expect(cookie?.maxAge).toBe(300);
  });

  it('leaves an ordinary sign-in alone', async () => {
    const res = await handler(req(), ctx);
    expect(createHandoffMock).not.toHaveBeenCalled();
    expect(res.headers?.Location).toBe('/');
    expect(res.cookies?.find((c) => c.name === 'install_handoff')).toBeUndefined();
  });

  it('ignores a malformed install_nonce cookie', async () => {
    const res = await handler(req('; install_nonce=not-a-challenge'), ctx);
    expect(createHandoffMock).not.toHaveBeenCalled();
    expect(res.headers?.Location).toBe('/');
  });

  it('falls back to an ordinary sign-in if the handoff cannot be recorded', async () => {
    createHandoffMock.mockRejectedValue(new Error('storage down'));
    const res = await handler(req(`; install_nonce=${CHALLENGE}`), ctx);
    expect(res.status).toBe(302);
    expect(res.headers?.Location).toBe('/');
    expect(attachMock).not.toHaveBeenCalled();
  });
});
