/**
 * Tests for GET /api/auth/install-poll — PKCE-style nonce binding.
 *
 * The install-poll endpoint now requires the raw verifier (nonce_verifier) rather
 * than the nonce directly. It derives the challenge as SHA256(verifier) and
 * looks up the session under that key. An attacker who intercepts the login URL
 * only sees the challenge and cannot poll without knowing the verifier preimage.
 */

import { jest } from '@jest/globals';
import { createHash } from 'crypto';
import type { HttpRequest, InvocationContext } from '@azure/functions';
import type { InstallNonceRecord } from '../services/tableStorage.js';

// ── Mock Azure Functions app ─────────────────────────────────────────────────

const httpMock = jest.fn();
jest.mock('@azure/functions', () => ({
  app: { http: httpMock },
}));

// ── Mock tableStorage ────────────────────────────────────────────────────────

const consumeInstallNonceMock = jest.fn<
  (nonce: string) => Promise<InstallNonceRecord | 'pending' | null>
>();
jest.mock('../services/tableStorage.js', () => ({
  consumeInstallNonce: (nonce: string) => consumeInstallNonceMock(nonce),
}));

// ── Mock securityHeaders (pass-through) ─────────────────────────────────────

jest.mock('../services/securityHeaders.js', () => ({
  withSecurity: (handler: unknown) => handler,
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import '../functions/auth/installPoll.js';

// ── Helpers ──────────────────────────────────────────────────────────────────

interface HttpRegistration {
  handler: (req: HttpRequest, ctx: InvocationContext) => Promise<{
    status: number;
    jsonBody?: unknown;
  }>;
}

function getHandler(): HttpRegistration['handler'] {
  const reg = httpMock.mock.calls.find((c) => c[0] === 'install-poll');
  if (!reg) throw new Error("Handler 'install-poll' not registered");
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

const sha256 = (input: string) => createHash('sha256').update(input).digest('hex');

// ── Tests ────────────────────────────────────────────────────────────────────

describe('GET /api/auth/install-poll (PKCE binding)', () => {
  let handler: HttpRegistration['handler'];

  beforeAll(() => {
    handler = getHandler();
  });

  beforeEach(() => {
    consumeInstallNonceMock.mockReset();
  });

  it('returns 400 when nonce_verifier is missing', async () => {
    const res = await handler(makeRequest({}), makeContext());
    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('nonce_verifier');
    expect(consumeInstallNonceMock).not.toHaveBeenCalled();
  });

  it('returns 400 when nonce_verifier is wrong length (32 chars required)', async () => {
    const res = await handler(
      makeRequest({ nonce_verifier: 'abc123' }),
      makeContext()
    );
    expect(res.status).toBe(400);
    expect((res.jsonBody as { error: string }).error).toContain('32 hex');
    expect(consumeInstallNonceMock).not.toHaveBeenCalled();
  });

  it('returns 400 for 64-char challenge submitted directly (wrong param)', async () => {
    // Attackers might try submitting the 64-char challenge via the new param.
    const oldStyleChallenge = 'a'.repeat(64);
    const res = await handler(
      makeRequest({ nonce_verifier: oldStyleChallenge }),
      makeContext()
    );
    expect(res.status).toBe(400);
    expect(consumeInstallNonceMock).not.toHaveBeenCalled();
  });

  it('returns 400 for non-hex characters in nonce_verifier', async () => {
    const res = await handler(
      makeRequest({ nonce_verifier: 'zzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzz' }),
      makeContext()
    );
    expect(res.status).toBe(400);
  });

  it('derives challenge as SHA256(verifier) before lookup', async () => {
    const verifier = 'abcdef0123456789abcdef0123456789'; // 32 hex chars
    const expectedChallenge = sha256(verifier);

    consumeInstallNonceMock.mockResolvedValueOnce('pending');

    await handler(makeRequest({ nonce_verifier: verifier }), makeContext());

    expect(consumeInstallNonceMock).toHaveBeenCalledWith(expectedChallenge);
  });

  it('returns 202 when OAuth is still in progress', async () => {
    const verifier = 'deadbeefdeadbeefdeadbeefdeadbeef';
    consumeInstallNonceMock.mockResolvedValueOnce('pending');

    const res = await handler(makeRequest({ nonce_verifier: verifier }), makeContext());
    expect(res.status).toBe(202);
    expect((res.jsonBody as { status: string }).status).toBe('pending');
  });

  it('returns 410 when nonce is expired or already consumed', async () => {
    const verifier = 'cafebabecafebabecafebabecafebabe';
    consumeInstallNonceMock.mockResolvedValueOnce(null);

    const res = await handler(makeRequest({ nonce_verifier: verifier }), makeContext());
    expect(res.status).toBe(410);
  });

  it('returns 200 with session payload on successful poll', async () => {
    const verifier = '11223344556677889900aabbccddeeff';
    const record: InstallNonceRecord = {
      sessionToken: 'tok_abc123',
      userId: 'user-guid-1234',
      email: 'alice@example.com',
      displayName: 'Alice',
      expiresAt: Date.now() + 300_000,
    };
    consumeInstallNonceMock.mockResolvedValueOnce(record);

    const res = await handler(makeRequest({ nonce_verifier: verifier }), makeContext());
    expect(res.status).toBe(200);
    const body = res.jsonBody as typeof record;
    expect(body.sessionToken).toBe(record.sessionToken);
    expect(body.userId).toBe(record.userId);
    expect(body.email).toBe(record.email);
    expect(body.displayName).toBe(record.displayName);
  });

  it('PKCE: attacker submitting wrong verifier reaches wrong challenge slot', async () => {
    // The attacker chose a challenge and embedded it in the victim's login URL.
    // The attacker's verifier hashes to a DIFFERENT challenge → misses the victim's session.
    // (In practice the attacker cannot know the CLI's verifier preimage.)
    const attackerVerifier = 'ffffffffffffffffffffffffffffffff';
    const attackerChallenge = sha256(attackerVerifier);

    // Both challenges return 'pending' — attacker gets nothing
    consumeInstallNonceMock.mockResolvedValue('pending');

    const res = await handler(makeRequest({ nonce_verifier: attackerVerifier }), makeContext());
    expect(res.status).toBe(202); // pending, not 200
    // Server looked up sha256(attackerVerifier), not any other slot
    expect(consumeInstallNonceMock).toHaveBeenCalledWith(attackerChallenge);
  });
});
