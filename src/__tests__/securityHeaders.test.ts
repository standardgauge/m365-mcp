import type { HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { secureHeaders, adminSpaHeaders, installLandingHeaders, withSecurity } from '../services/securityHeaders.js';

const fakeRequest = {} as HttpRequest;
const fakeContext = {} as InvocationContext;

describe('withSecurity', () => {
  it('adds the full secureHeaders set to a bare response', async () => {
    const handler = async (): Promise<HttpResponseInit> => ({ status: 200, jsonBody: { ok: true } });
    const wrapped = withSecurity(handler);

    const result = (await wrapped(fakeRequest, fakeContext)) as HttpResponseInit;

    expect(result.headers).toMatchObject(secureHeaders);
    expect(result.status).toBe(200);
    expect(result.jsonBody).toEqual({ ok: true });
  });

  it('preserves explicit handler headers (caller wins on collision)', async () => {
    const handler = async (): Promise<HttpResponseInit> => ({
      status: 200,
      headers: {
        'Cache-Control': 'no-store',
        'X-Frame-Options': 'SAMEORIGIN', // intentional override
      },
    });
    const wrapped = withSecurity(handler);

    const result = (await wrapped(fakeRequest, fakeContext)) as HttpResponseInit;

    expect(result.headers).toMatchObject({
      'Cache-Control': 'no-store',
      'X-Frame-Options': 'SAMEORIGIN',
      'Strict-Transport-Security': secureHeaders['Strict-Transport-Security'],
      'X-Content-Type-Options': 'nosniff',
    });
  });

  it('handles handlers that return undefined gracefully', async () => {
    const handler = async (): Promise<HttpResponseInit> => undefined as unknown as HttpResponseInit;
    const wrapped = withSecurity(handler);

    const result = (await wrapped(fakeRequest, fakeContext)) as HttpResponseInit;

    expect(result.headers).toMatchObject(secureHeaders);
  });

  it('includes Content-Security-Policy in the default secureHeaders set', () => {
    expect(secureHeaders['Content-Security-Policy']).toContain("default-src 'none'");
    expect(secureHeaders['Content-Security-Policy']).toContain("frame-ancestors 'none'");
  });

  it('adminSpaHeaders CSP allows MSAL connect targets and blocks framing', () => {
    const csp = adminSpaHeaders['Content-Security-Policy'];
    expect(csp).toContain('https://login.microsoftonline.com');
    expect(csp).toContain('https://graph.microsoft.com');
    expect(csp).toContain("frame-ancestors 'none'");
    expect(csp).toContain("default-src 'none'");
  });

  it('installLandingHeaders CSP allows inline styles and blocks framing', () => {
    const csp = installLandingHeaders['Content-Security-Policy'];
    expect(csp).toContain("style-src 'unsafe-inline'");
    expect(csp).toContain("frame-ancestors 'none'");
    expect(csp).toContain("default-src 'none'");
    // must NOT allow scripts
    expect(csp).not.toContain('script-src');
  });

  it('install landing handler can override the default CSP via installLandingHeaders', async () => {
    const handler = async (): Promise<HttpResponseInit> => ({
      status: 200,
      headers: { 'Content-Type': 'text/html; charset=utf-8', ...installLandingHeaders },
    });
    const wrapped = withSecurity(handler);
    const result = (await wrapped(fakeRequest, fakeContext)) as HttpResponseInit;

    expect((result.headers as Record<string, string>)['Content-Security-Policy']).toBe(
      installLandingHeaders['Content-Security-Policy'],
    );
    expect((result.headers as Record<string, string>)['Strict-Transport-Security']).toBe(
      secureHeaders['Strict-Transport-Security'],
    );
  });

  it('admin SPA handler can override the default CSP via its own headers', async () => {
    const handler = async (): Promise<HttpResponseInit> => ({
      status: 200,
      headers: { 'Content-Type': 'text/html; charset=utf-8', ...adminSpaHeaders },
    });
    const wrapped = withSecurity(handler);
    const result = (await wrapped(fakeRequest, fakeContext)) as HttpResponseInit;

    expect((result.headers as Record<string, string>)['Content-Security-Policy']).toBe(
      adminSpaHeaders['Content-Security-Policy'],
    );
    expect((result.headers as Record<string, string>)['Strict-Transport-Security']).toBe(
      secureHeaders['Strict-Transport-Security'],
    );
  });
});
