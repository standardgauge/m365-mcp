/**
 * Tests for the GET /health liveness endpoint.
 *
 * Covers:
 *   - returns 200 with { status: 'ok', sha } JSON
 *   - sha reflects the GIT_SHA build-arg env var when set
 *   - sha falls back to '' when GIT_SHA is unset
 *   - registered on the literal `health` route (so it wins over the
 *     serveAdmin `{*restOfPath}` SPA catch-all)
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

// ── Mock Azure Functions app ─────────────────────────────────────────────────

const httpMock = jest.fn();

jest.mock('@azure/functions', () => ({
  app: { http: httpMock },
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import '../functions/health.js';

interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<{
    status: number;
    jsonBody?: unknown;
    headers?: Record<string, string>;
  }>;
  route?: string;
  authLevel?: string;
  methods?: string[];
}

function getRegistration(name: string): HttpRegistration {
  const reg = httpMock.mock.calls.find((call) => call[0] === name);
  if (!reg) throw new Error(`Handler '${name}' not registered`);
  return reg[1] as HttpRegistration;
}

const context = {} as InvocationContext;
const request = {} as HttpRequest;

describe('GET /health', () => {
  const originalSha = process.env.GIT_SHA;

  afterEach(() => {
    if (originalSha === undefined) {
      delete process.env.GIT_SHA;
    } else {
      process.env.GIT_SHA = originalSha;
    }
  });

  it('is registered on the literal `health` route as an anonymous GET', () => {
    const reg = getRegistration('health');
    expect(reg.route).toBe('health');
    expect(reg.authLevel).toBe('anonymous');
    expect(reg.methods).toEqual(['GET']);
  });

  it('returns 200 with status ok and the deployed sha', async () => {
    process.env.GIT_SHA = 'abc1234';
    const res = await getRegistration('health').handler(request, context);
    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({ status: 'ok', sha: 'abc1234' });
  });

  it('falls back to an empty sha when GIT_SHA is unset', async () => {
    delete process.env.GIT_SHA;
    const res = await getRegistration('health').handler(request, context);
    expect(res.status).toBe(200);
    expect(res.jsonBody).toEqual({ status: 'ok', sha: '' });
  });
});
