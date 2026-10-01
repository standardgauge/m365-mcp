/**
 * Tests that unmatched /api/* paths get a real 404 from serveAdmin.
 *
 * serveAdmin is the site-wide catch-all and falls back to index.html so the
 * admin SPA's client-side deep links work. That fallback used to apply to
 * /api/* too, which meant a call to a nonexistent API route came back 200
 * text/html. found the operations runbook probing
 * `curl -f https://your-mcp-host.example.com/api/health` — a route that has never
 * existed — and reporting OK. That check could not have failed even with the
 * API completely dead.
 *
 * Covers:
 *   - unmatched /api/* → 404 JSON, never the SPA shell
 *   - the specific path from the runbook that silently passed
 *   - SPA deep links outside /api still fall back to index.html
 */

import { jest } from '@jest/globals';
import type { HttpRequest, InvocationContext } from '@azure/functions';

jest.mock('@azure/functions', () => ({ app: { http: jest.fn() } }));

import { app } from '@azure/functions';
import '../functions/admin/serveAdmin.js';

interface Res {
  status: number;
  jsonBody?: unknown;
  body?: unknown;
  headers?: Record<string, string>;
}
interface HttpRegistration {
  handler: (req: HttpRequest, context: InvocationContext) => Promise<Res>;
}
const httpMock = app.http as unknown as jest.Mock<(name: string, opts: HttpRegistration) => void>;
const reg = httpMock.mock.calls.find((c) => c[0] === 'serveAdmin');
if (!reg) throw new Error('serveAdmin handler was not registered');
const handler = reg[1].handler;

const ctx = { error: jest.fn(), warn: jest.fn() } as unknown as InvocationContext;

function req(restOfPath: string): HttpRequest {
  return {
    method: 'GET',
    params: { restOfPath },
    headers: new Map<string, string>(),
  } as unknown as HttpRequest;
}

beforeEach(() => jest.clearAllMocks());

describe('serveAdmin — unmatched /api/* is not the SPA', () => {
  it('404s the exact path the runbook probed and that always passed', async () => {
    const res = await handler(req('api/health'), ctx);

    expect(res.status).toBe(404);
    expect(res.headers?.['Content-Type']).toMatch(/application\/json/);
    expect(res.jsonBody).toMatchObject({ error: 'Not found' });
  });

  it.each([
    'api',
    'api/',
    'api/does-not-exist',
    'api/auth/nope',
    'API/Health',
  ])('404s unmatched path /%s', async (p) => {
    const res = await handler(req(p), ctx);
    expect(res.status).toBe(404);
  });

  it('never returns the SPA shell for an /api path', async () => {
    const res = await handler(req('api/health'), ctx);
    const rendered = typeof res.body === 'string' ? res.body : '';
    expect(rendered).not.toContain('<!DOCTYPE html>');
    expect(res.headers?.['Content-Type'] ?? '').not.toContain('text/html');
  });

  it('still falls back to the SPA for non-API deep links', async () => {
    // The narrowing must not break client-side routing: anything outside /api
    // still resolves to index.html so the SPA can handle the route itself.
    const res = await handler(req('settings/profile'), ctx);

    expect(res.status).toBe(200);
    expect(res.headers?.['Content-Type']).toMatch(/text\/html/);
    expect(String(res.body)).toContain('<!DOCTYPE html>');
  });
});
