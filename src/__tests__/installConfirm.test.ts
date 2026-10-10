/**
 * Tests for the install handoff confirmation.
 *
 * A sign-in link carries the installer's challenge, and whoever wrote the link
 * holds the verifier that install-poll wants. So the OAuth callback must not
 * hand the signed-in session to the challenge on its own: the browser has to
 * enter the code the installer shows. These tests hold:
 *
 *   - the code derivation, and that install.sh, install.ps1 and the extension
 *     all print the code the server checks;
 *   - GET /api/auth/install-confirm renders the request without the code;
 *   - POST acts only on the session this sign-in created, only with the right
 *     code, once, and hands the installer a new client session rather than the
 *     browser's own token;
 *   - POST refuses a request from another origin, a same-site sibling included;
 *   - wrong codes are counted and discard the handoff at the limit;
 *   - "This wasn't me" discards it.
 */

import { jest } from '@jest/globals';
import { createHash } from 'crypto';
import * as crypto from 'crypto';
import { execFileSync } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';
import * as vm from 'vm';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { InstallNonceRecord, StoredInstallHandoff } from '../services/tableStorage.js';

// ── Mocks ───────────────────────────────────────────────────────────────────

const httpMock = jest.fn();
jest.mock('@azure/functions', () => ({ app: { http: httpMock } }));

const getInstallHandoffMock = jest.fn<(id: string) => Promise<StoredInstallHandoff | null>>();
const recordFailureMock = jest.fn<(id: string, s: StoredInstallHandoff) => Promise<number | null>>();
const deleteHandoffMock = jest.fn<(id: string, etag?: string) => Promise<boolean>>();
const attachMock = jest.fn<(nonce: string, rec: InstallNonceRecord) => Promise<boolean>>();
jest.mock('../services/tableStorage.js', () => ({
  getInstallHandoff: (id: string) => getInstallHandoffMock(id),
  recordInstallHandoffFailure: (id: string, s: StoredInstallHandoff) => recordFailureMock(id, s),
  deleteInstallHandoff: (id: string, etag?: string) => deleteHandoffMock(id, etag),
  attachSessionToInstallNonce: (nonce: string, rec: InstallNonceRecord) => attachMock(nonce, rec),
}));

import type { UserSession } from '../services/tokenCache.js';
const getSessionByTokenMock = jest.fn<(t: string) => Promise<UserSession | undefined>>();
const storeSessionMock = jest.fn<(s: UserSession) => Promise<void>>();
const deleteSessionByKeyMock = jest.fn<(k: string, u: string) => Promise<void>>();
jest.mock('../services/tokenCache.js', () => ({
  getSessionByToken: (t: string) => getSessionByTokenMock(t),
  storeSession: (s: UserSession) => storeSessionMock(s),
  deleteSessionByKey: (k: string, u: string) => deleteSessionByKeyMock(k, u),
}));

// Stand-in for the keyed HMAC: deterministic, and distinct per token.
const fakeHash = (t: string) => createHash('sha256').update('k:' + t).digest('hex');
jest.mock('../services/credentialCrypto.js', () => ({
  hashSessionToken: (t: string) => fakeHash(t),
}));

jest.mock('../services/securityHeaders.js', () => ({
  withSecurity: (handler: unknown) => handler,
}));

// The update payload is signed; extensionUpdate.test.ts covers that.
// Here only the served code matters, so the signer is a stand-in.
jest.mock('../services/extensionSigning.js', () => ({
  extensionPublicKey: () => 'test-public-key',
  signExtensionPayload: (payload: unknown) => ({ payload: JSON.stringify(payload), signature: '' }),
}));

// Served code takes its origin from OAUTH_REDIRECT_URI, not the Host header.
process.env.OAUTH_REDIRECT_URI = 'https://mcp.example.com/api/auth/callback';

import '../functions/auth/installConfirm.js';
import '../functions/install/installEndpoint.js';
import {
  confirmationCodeMatches,
  formatConfirmationCode,
  installConfirmationCode,
} from '../services/installConfirm.js';

// ── Helpers ─────────────────────────────────────────────────────────────────

interface Res {
  status: number;
  body?: string;
  headers?: Record<string, string>;
  cookies?: Array<{ name: string; value: string; maxAge?: number }>;
}
type Handler = (req: HttpRequest, ctx: InvocationContext) => Promise<Res>;

function handlerFor(name: string): Handler {
  const reg = httpMock.mock.calls.find((c) => c[0] === name);
  if (!reg) throw new Error(`Handler '${name}' not registered`);
  return (reg[1] as { handler: Handler }).handler;
}
const confirm = handlerFor('install-confirm');
const extensionUpdate = handlerFor('extensionUpdate');

const VERIFIER = 'a'.repeat(32);
const CHALLENGE = createHash('sha256').update(VERIFIER).digest('hex');
const HANDOFF_ID = 'b'.repeat(64);
const SESSION_TOKEN = 'c'.repeat(64);
const CODE = formatConfirmationCode(installConfirmationCode(CHALLENGE));

function stored(overrides: Partial<StoredInstallHandoff['record']> = {}): StoredInstallHandoff {
  return {
    etag: 'W/"etag-1"',
    record: {
      challenge: CHALLENGE,
      sessionTokenHash: fakeHash(SESSION_TOKEN),
      userId: 'user-1',
      email: 'adele@fabrikam.com',
      displayName: 'Adele Vance',
      deviceLabel: 'ADELE-LAPTOP',
      attempts: 0,
      expiresAt: Date.now() + 60_000,
      ...overrides,
    },
  };
}

const HOST = 'mcp.example.com';
// What a browser sends when the confirmation page's own form posts.
const FORM_POST_HEADERS: Record<string, string> = {
  host: HOST,
  origin: `https://${HOST}`,
  'sec-fetch-site': 'same-origin',
};

function req(opts: {
  method?: string;
  cookie?: string;
  form?: Record<string, string>;
  headers?: Record<string, string>;
}): HttpRequest {
  const body = new URLSearchParams(opts.form ?? {}).toString();
  const method = opts.method ?? 'GET';
  const headers = new Map<string, string>(
    Object.entries(opts.headers ?? (method === 'POST' ? FORM_POST_HEADERS : { host: HOST }))
  );
  if (opts.cookie) headers.set('cookie', opts.cookie);
  return { method, headers, text: async () => body } as unknown as HttpRequest;
}

function browserSession(overrides: Partial<UserSession> = {}): UserSession {
  return {
    userId: 'user-1',
    homeAccountId: 'home-1',
    displayName: 'Adele Vance',
    email: 'adele@fabrikam.com',
    tenantId: 'tenant-1',
    accessToken: 'graph-access-token',
    expiresAt: Date.now() + 3_600_000,
    sessionToken: SESSION_TOKEN,
    sessionCreatedAt: Date.now() - 1_000,
    sessionAbsoluteCreatedAt: Date.now() - 60_000,
    deviceLabel: 'ADELE-LAPTOP',
    kind: 'browser',
    _storageKey: 'browser-row',
    ...overrides,
  };
}

const BOTH_COOKIES = `install_handoff=${HANDOFF_ID}; mcp_session=${SESSION_TOKEN}`;
const ctx = { log: jest.fn(), warn: jest.fn(), error: jest.fn() } as unknown as InvocationContext;

beforeEach(() => {
  jest.clearAllMocks();
  getInstallHandoffMock.mockResolvedValue(stored());
  deleteHandoffMock.mockResolvedValue(true);
  attachMock.mockResolvedValue(true);
  recordFailureMock.mockResolvedValue(4);
  getSessionByTokenMock.mockResolvedValue(browserSession());
  storeSessionMock.mockImplementation(async (sess) => {
    sess._storageKey = 'client-row';
  });
  deleteSessionByKeyMock.mockResolvedValue(undefined);
});

// ── Code derivation ─────────────────────────────────────────────────────────

describe('confirmation code', () => {
  it('is eight uppercase hex characters derived from the challenge', () => {
    const code = installConfirmationCode(CHALLENGE);
    expect(code).toMatch(/^[0-9A-F]{8}$/);
    const expected = createHash('sha256')
      .update('m365-mcp-install-confirm:' + CHALLENGE)
      .digest('hex')
      .slice(0, 8)
      .toUpperCase();
    expect(code).toBe(expected);
  });

  it('is not a substring of the challenge in the address bar', () => {
    expect(CHALLENGE.toUpperCase()).not.toContain(installConfirmationCode(CHALLENGE));
  });

  it('accepts the code with or without the hyphen, any case, spaces, and O for 0', () => {
    const raw = installConfirmationCode(CHALLENGE);
    expect(confirmationCodeMatches(CHALLENGE, raw)).toBe(true);
    expect(confirmationCodeMatches(CHALLENGE, formatConfirmationCode(raw).toLowerCase())).toBe(true);
    expect(confirmationCodeMatches(CHALLENGE, ` ${raw.slice(0, 4)} ${raw.slice(4)} `)).toBe(true);
    expect(confirmationCodeMatches(CHALLENGE, raw.replace(/0/g, 'O'))).toBe(true);
  });

  it('rejects another challenge\'s code, a truncated code, and empty input', () => {
    const other = createHash('sha256').update('d'.repeat(32)).digest('hex');
    expect(confirmationCodeMatches(CHALLENGE, installConfirmationCode(other))).toBe(false);
    expect(confirmationCodeMatches(CHALLENGE, installConfirmationCode(CHALLENGE).slice(0, 7))).toBe(false);
    expect(confirmationCodeMatches(CHALLENGE, '')).toBe(false);
  });
});

describe('installers print the code the server checks', () => {
  const installDir = path.join(process.cwd(), 'src', 'install');

  it('install.sh derives the same code', () => {
    const template = fs.readFileSync(path.join(installDir, 'install-mcp.sh.template'), 'utf-8');
    const lines = template
      .split('\n')
      .filter((l) => l.startsWith('CONFIRM_HEX=') || l.startsWith('CONFIRM_CODE='));
    expect(lines).toHaveLength(2);
    const out = execFileSync('bash', ['-c', `set -e\n${lines.join('\n')}\nprintf '%s' "$CONFIRM_CODE"`], {
      env: { ...process.env, CHALLENGE },
    }).toString();
    expect(out).toBe(CODE);
    // Shown before the browser opens, so the user has it when the page asks.
    expect(template.indexOf('Your confirmation code')).toBeLessThan(template.indexOf('open "${LOGIN_URL}"'));
  });

  it('install.ps1 derives the same code', () => {
    const template = fs.readFileSync(path.join(installDir, 'install-mcp.ps1.template'), 'utf-8');
    expect(template).toContain('"m365-mcp-install-confirm:$Challenge"');
    expect(template).toContain(".Substring(0, 8).ToUpper()");
    expect(template).toContain("$hex.Substring(0, 4) + '-' + $hex.Substring(4, 4)");
    expect(template).toContain('$confirmCode = Get-ConfirmationCode -Challenge $Challenge');
    let pwsh: string | null = null;
    try {
      pwsh = execFileSync('which', ['pwsh']).toString().trim() || null;
    } catch { /* not installed */ }
    if (!pwsh) return;
    const fn = template.slice(
      template.indexOf('function Get-ConfirmationCode'),
      template.indexOf('function Start-OAuthFlow')
    );
    const out = execFileSync(pwsh, ['-NoProfile', '-Command', `${fn}\nGet-ConfirmationCode -Challenge '${CHALLENGE}'`])
      .toString()
      .trim();
    expect(out).toBe(CODE);
  });

  it('the extension derives the same code and shows it before sign-in', async () => {
    const res = (await extensionUpdate(
      {
        headers: new Map<string, string>([['host', 'mcp.example.com']]),
        query: { get: () => null },
      } as unknown as HttpRequest,
      ctx
    )) as unknown as { jsonBody: { files: Record<string, string> } };
    const serverJs = res.jsonBody.files['server/index.js'];
    expect(() => new vm.Script(serverJs)).not.toThrow();
    const start = serverJs.indexOf('function confirmationCode(challenge)');
    const end = serverJs.indexOf('\n}\n', start) + 2;
    expect(start).toBeGreaterThan(-1);
    const derive = new Function('crypto', `${serverJs.slice(start, end)}; return confirmationCode;`)(crypto) as (
      c: string
    ) => string;
    expect(derive(CHALLENGE)).toBe(CODE);
    // The extension has no window, so it opens a local page carrying the code.
    expect(serverJs).toContain('openInBrowser(pageFile || loginUrl)');
    expect(serverJs).toContain("'&device_label=' + encodeURIComponent(deviceLabel())");
  });
});

// ── GET ─────────────────────────────────────────────────────────────────────

describe('GET /api/auth/install-confirm', () => {
  it('410 without a handoff cookie', async () => {
    const res = await confirm(req({}), ctx);
    expect(res.status).toBe(410);
    expect(getInstallHandoffMock).not.toHaveBeenCalled();
  });

  it('410 when the handoff is unknown or expired', async () => {
    getInstallHandoffMock.mockResolvedValue(null);
    const res = await confirm(req({ cookie: BOTH_COOKIES }), ctx);
    expect(res.status).toBe(410);
  });

  it('shows the account and device label, asks for the code, and never prints it', async () => {
    const res = await confirm(req({ cookie: BOTH_COOKIES }), ctx);
    expect(res.status).toBe(200);
    expect(res.headers?.['Content-Type']).toMatch(/text\/html/);
    expect(res.headers?.['Cache-Control']).toBe('no-store');
    expect(res.headers?.['Content-Security-Policy']).toContain("form-action 'self'");
    expect(res.body).toContain('adele@fabrikam.com');
    expect(res.body).toContain('ADELE-LAPTOP');
    expect(res.body).toContain('name="code"');
    expect(res.body).toContain("This wasn't me");
    const raw = installConfirmationCode(CHALLENGE);
    expect(res.body).not.toContain(raw);
    expect(res.body).not.toContain(CODE);
    expect(attachMock).not.toHaveBeenCalled();
  });

  it('escapes the installer-supplied device label', async () => {
    getInstallHandoffMock.mockResolvedValue(stored({ deviceLabel: '<script>x</script>' }));
    const res = await confirm(req({ cookie: BOTH_COOKIES }), ctx);
    expect(res.body).not.toContain('<script>x');
    expect(res.body).toContain('&lt;script&gt;x');
  });
});

// ── POST ────────────────────────────────────────────────────────────────────

describe('POST /api/auth/install-confirm', () => {
  it('a signer who was only sent the link cannot complete it: wrong code attaches nothing', async () => {
    const res = await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: { code: '0000-0000' } }), ctx);
    expect(res.status).toBe(400);
    expect(res.body).toContain('4 attempts left');
    expect(recordFailureMock).toHaveBeenCalledWith(HANDOFF_ID, expect.objectContaining({ etag: 'W/"etag-1"' }));
    expect(attachMock).not.toHaveBeenCalled();
  });

  it('an empty submission counts as a wrong code', async () => {
    const res = await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: {} }), ctx);
    expect(res.status).toBe(400);
    expect(attachMock).not.toHaveBeenCalled();
  });

  it('discards the handoff at the attempt limit', async () => {
    recordFailureMock.mockResolvedValue(0);
    const res = await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: { code: '0000-0000' } }), ctx);
    expect(res.status).toBe(400);
    expect(res.body).toContain('Too many attempts');
    expect(res.cookies?.find((c) => c.name === 'install_handoff')?.maxAge).toBe(0);
    expect(attachMock).not.toHaveBeenCalled();
  });

  it('attaches a new client session, never the browser\'s own token, with the right code', async () => {
    const res = await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: { code: CODE } }), ctx);
    expect(res.status).toBe(200);
    expect(res.body).toContain('Connected');
    expect(deleteHandoffMock).toHaveBeenCalledWith(HANDOFF_ID, 'W/"etag-1"');
    expect(getSessionByTokenMock).toHaveBeenCalledWith(SESSION_TOKEN);

    expect(storeSessionMock).toHaveBeenCalledTimes(1);
    const client = storeSessionMock.mock.calls[0][0];
    expect(client.kind).toBe('client');
    expect(client.sessionToken).toMatch(/^[0-9a-f]{64}$/);
    expect(client.sessionToken).not.toBe(SESSION_TOKEN);
    expect(client.userId).toBe('user-1');
    expect(client.homeAccountId).toBe('home-1');
    expect(client.tenantId).toBe('tenant-1');
    expect(client.deviceLabel).toBe('ADELE-LAPTOP');

    expect(attachMock).toHaveBeenCalledTimes(1);
    const [nonce, rec] = attachMock.mock.calls[0];
    expect(nonce).toBe(CHALLENGE);
    expect(rec.sessionToken).toBe(client.sessionToken);
    expect(rec.userId).toBe('user-1');
    expect(res.cookies?.find((c) => c.name === 'install_handoff')?.maxAge).toBe(0);
  });

  it('keeps the sign-in\'s absolute lifetime anchor on the client session', async () => {
    const anchor = Date.now() - 5 * 86_400_000;
    getSessionByTokenMock.mockResolvedValue(browserSession({ sessionAbsoluteCreatedAt: anchor }));
    await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: { code: CODE } }), ctx);
    expect(storeSessionMock.mock.calls[0][0].sessionAbsoluteCreatedAt).toBe(anchor);
  });

  it('attaches nothing when the browser session is gone or storage cannot answer', async () => {
    const { SessionStoreUnavailableError } = await import('../services/sessionStoreError.js');
    for (const outcome of [
      () => getSessionByTokenMock.mockResolvedValue(undefined),
      () => getSessionByTokenMock.mockRejectedValue(new SessionStoreUnavailableError(new Error('503'))),
    ]) {
      jest.clearAllMocks();
      deleteHandoffMock.mockResolvedValue(true);
      outcome();
      const res = await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: { code: CODE } }), ctx);
      expect(res.status).toBe(500);
      expect(res.body).not.toContain('Connected');
      expect(storeSessionMock).not.toHaveBeenCalled();
      expect(attachMock).not.toHaveBeenCalled();
    }
  });

  it('deletes the client session when it cannot be attached', async () => {
    attachMock.mockResolvedValue(false);
    await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: { code: CODE } }), ctx);
    expect(deleteSessionByKeyMock).toHaveBeenCalledWith('client-row', 'user-1');
  });

  it('refuses a POST from another origin, a same-site sibling included, with the right code', async () => {
    const refused: Array<Record<string, string>> = [
      { host: HOST, origin: 'https://evil.example.com', 'sec-fetch-site': 'same-site' },
      { host: HOST, origin: 'https://attacker.test', 'sec-fetch-site': 'cross-site' },
      { host: HOST, origin: `https://${HOST}`, 'sec-fetch-site': 'same-site' },
      { host: HOST, 'sec-fetch-site': 'same-origin' },
      { host: HOST, origin: 'null' },
    ];
    for (const headers of refused) {
      for (const form of [{ code: CODE }, { action: 'cancel' }] as Array<Record<string, string>>) {
        const res = await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form, headers }), ctx);
        expect(res.status).toBe(403);
      }
    }
    expect(deleteHandoffMock).not.toHaveBeenCalled();
    expect(recordFailureMock).not.toHaveBeenCalled();
    expect(attachMock).not.toHaveBeenCalled();
  });

  it('refuses without the session the callback issued, even with the right code', async () => {
    for (const cookie of [`install_handoff=${HANDOFF_ID}`, `install_handoff=${HANDOFF_ID}; mcp_session=${'e'.repeat(64)}`]) {
      const res = await confirm(req({ method: 'POST', cookie, form: { code: CODE } }), ctx);
      expect(res.status).toBe(403);
    }
    expect(attachMock).not.toHaveBeenCalled();
    expect(deleteHandoffMock).not.toHaveBeenCalled();
  });

  it('is one-time: a request that loses the delete race attaches nothing', async () => {
    deleteHandoffMock.mockResolvedValue(false);
    const res = await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: { code: CODE } }), ctx);
    expect(res.status).toBe(410);
    expect(attachMock).not.toHaveBeenCalled();
  });

  it('reports a failed attach rather than claiming success', async () => {
    attachMock.mockResolvedValue(false);
    const res = await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: { code: CODE } }), ctx);
    expect(res.status).toBe(500);
    expect(res.body).not.toContain('Connected');
  });

  it('"This wasn\'t me" discards the handoff', async () => {
    const res = await confirm(req({ method: 'POST', cookie: BOTH_COOKIES, form: { action: 'cancel' } }), ctx);
    expect(res.status).toBe(200);
    expect(res.body).toContain('Nothing was connected');
    expect(deleteHandoffMock).toHaveBeenCalledWith(HANDOFF_ID, undefined);
    expect(attachMock).not.toHaveBeenCalled();
  });
});
