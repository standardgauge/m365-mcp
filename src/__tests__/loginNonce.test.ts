/**
 * Tests for GET /api/auth/login — PKCE install_nonce format enforcement.
 *
 * install_nonce is now the PKCE code_challenge: SHA256(verifier), a 64-hex-char
 * string. The old format (32 hex chars = raw nonce) must now be rejected.
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── Mock Azure Functions app ─────────────────────────────────────────────────

const httpMock = jest.fn();
jest.mock('@azure/functions', () => ({
  app: { http: httpMock },
}));

// ── Mock graphClient (getAuthCodeUrl) ────────────────────────────────────────

const getAuthCodeUrlMock = jest.fn<(state: string) => Promise<string>>();
jest.mock('../services/graphClient.js', () => ({
  getAuthCodeUrl: (state: string) => getAuthCodeUrlMock(state),
}));

// ── Mock securityHeaders (pass-through) ─────────────────────────────────────

jest.mock('../services/securityHeaders.js', () => ({
  withSecurity: (handler: unknown) => handler,
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import '../functions/auth/login.js';

// ── Helpers ──────────────────────────────────────────────────────────────────

interface HttpRegistration {
  handler: (req: HttpRequest, ctx: InvocationContext) => Promise<{
    status: number;
    jsonBody?: unknown;
    cookies?: unknown[];
  }>;
}

function getHandler(): HttpRegistration['handler'] {
  const reg = httpMock.mock.calls.find((c) => c[0] === 'login');
  if (!reg) throw new Error("Handler 'login' not registered");
  return (reg[1] as HttpRegistration).handler;
}

function makeRequest(params: Record<string, string | null>): HttpRequest {
  return {
    query: {
      get: (key: string) => params[key] ?? null,
      has: (key: string) => key in params && params[key] !== null,
    },
    headers: { get: () => null },
  } as unknown as HttpRequest;
}

function makeContext(): InvocationContext {
  return {
    error: jest.fn(),
    warn: jest.fn(),
    log: jest.fn(),
  } as unknown as InvocationContext;
}

// ── Tests ────────────────────────────────────────────────────────────────────

describe('GET /api/auth/login — install_nonce PKCE format', () => {
  let handler: HttpRegistration['handler'];

  beforeAll(() => {
    handler = getHandler();
    getAuthCodeUrlMock.mockResolvedValue(
      'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?state=test'
    );
  });

  it('allows requests without install_nonce', async () => {
    const res = await handler(makeRequest({}), makeContext());
    expect(res.status).toBe(302);
  });

  it('rejects old 32-hex-char raw nonce (pre-PKCE format)', async () => {
    const oldNonce = 'a'.repeat(32);
    const res = await handler(makeRequest({ install_nonce: oldNonce }), makeContext());
    expect(res.status).toBe(400);
    const body = res.jsonBody as { error: string };
    expect(body.error).toContain('64 hex');
  });

  it('accepts 64-hex-char PKCE challenge', async () => {
    const challenge = 'b'.repeat(64);
    const res = await handler(makeRequest({ install_nonce: challenge }), makeContext());
    expect(res.status).toBe(302);
  });

  it('rejects install_nonce with non-hex characters', async () => {
    const nonHex = 'z'.repeat(64);
    const res = await handler(makeRequest({ install_nonce: nonHex }), makeContext());
    expect(res.status).toBe(400);
  });

  it('rejects install_nonce longer than 64 hex chars', async () => {
    const tooLong = 'a'.repeat(65);
    const res = await handler(makeRequest({ install_nonce: tooLong }), makeContext());
    expect(res.status).toBe(400);
  });

  it('rejects install_nonce shorter than 64 hex chars', async () => {
    const tooShort = 'a'.repeat(63);
    const res = await handler(makeRequest({ install_nonce: tooShort }), makeContext());
    expect(res.status).toBe(400);
  });
});
