/**
 * Tests for FRONTEND_URL resolution and reserved-path rejection.
 *
 * Background:. Every tenant had FRONTEND_URL = `<host>/admin`. The Azure
 * Functions host reserves /admin/* for its own key-protected admin API, so the
 * post-login redirect 404'd before reaching any function. It shipped in March
 * and was found in September because nothing validated the value.
 *
 * Covers:
 *   - reserved paths detected in both absolute and relative form
 *   - prefix matching respects path boundaries (/administration is not /admin)
 *   - safe values pass through untouched
 *   - a reserved value degrades to the site root and reports the problem,
 *     rather than throwing and taking authentication down with it
 */

import { jest } from '@jest/globals';
import {
  reservedPrefixFor,
  resolveFrontendUrl,
  resolveAuthUrlBase,
  DEFAULT_FRONTEND_URL,
} from '../services/frontendUrl.js';

describe('reservedPrefixFor', () => {
  it.each([
    ['https://mcp.example.com/admin', '/admin'],
    ['https://mcp.example.com/admin/', '/admin'],
    ['https://mcp.example.com/admin/settings', '/admin'],
    ['/admin', '/admin'],
    ['/admin/', '/admin'],
    ['/ADMIN', '/admin'],
    ['/runtime/webhooks', '/runtime'],
    ['https://mcp.example.com/admin?userId=abc', '/admin'],
  ])('flags %s as reserved', (input, expected) => {
    expect(reservedPrefixFor(input)).toBe(expected);
  });

  it.each([
    ['https://mcp.example.com/'],
    ['https://mcp.example.com'],
    ['/'],
    ['/administration'],
    ['/admin-portal'],
    ['https://mcp.example.com/settings'],
    [undefined],
    [''],
  ])('does not flag %s', (input) => {
    expect(reservedPrefixFor(input as string | undefined)).toBeNull();
  });
});

describe('resolveFrontendUrl', () => {
  it('passes a safe absolute URL through unchanged', () => {
    const onProblem = jest.fn();
    expect(resolveFrontendUrl('https://mcp.example.com/', onProblem)).toBe(
      'https://mcp.example.com/'
    );
    expect(onProblem).not.toHaveBeenCalled();
  });

  it('falls back to the site root when unset or empty', () => {
    const onProblem = jest.fn();
    expect(resolveFrontendUrl(undefined, onProblem)).toBe(DEFAULT_FRONTEND_URL);
    expect(resolveFrontendUrl('', onProblem)).toBe(DEFAULT_FRONTEND_URL);
    expect(onProblem).not.toHaveBeenCalled();
  });

  it('rejects the exact misconfiguration and reports why', () => {
    const onProblem = jest.fn();
    const result = resolveFrontendUrl('https://mcp.example.com/admin', onProblem);

    expect(result).toBe(DEFAULT_FRONTEND_URL);
    expect(onProblem).toHaveBeenCalledTimes(1);

    const message = onProblem.mock.calls[0][0] as string;
    expect(message).toContain('https://mcp.example.com/admin');
    expect(message).toContain('/admin');
    expect(message).toContain('frontend-url');
  });

  it('degrades rather than throwing, so a bad value cannot break login', () => {
    expect(() => resolveFrontendUrl('/admin', () => {})).not.toThrow();
    expect(resolveFrontendUrl('/admin', () => {})).toBe(DEFAULT_FRONTEND_URL);
  });
});

/**
 * resolveAuthUrlBase — the base for the MCP re-authentication link.
 *
 * Codex flagged on PR #150 that mcpEndpoint built this from FRONTEND_URL and
 * stripped only an exact trailing `/admin`, so `/admin/settings` and
 * `/runtime` still produced login URLs the Functions host intercepts. The fix
 * takes the origin and drops the path, so no configured path — reserved or
 * otherwise — can reach the emitted link.
 */
describe('resolveAuthUrlBase', () => {
  it.each([
    ['https://mcp.example.com/admin'],
    ['https://mcp.example.com/admin/'],
    ['https://mcp.example.com/admin/settings'],
    ['https://mcp.example.com/runtime'],
    ['https://mcp.example.com/runtime/webhooks/foo'],
    ['https://mcp.example.com/some/other/path'],
    ['https://mcp.example.com/admin?userId=abc#frag'],
    ['https://mcp.example.com/'],
    ['https://mcp.example.com'],
  ])('drops the path from %s', (input) => {
    expect(resolveAuthUrlBase(input, undefined, () => {})).toBe('https://mcp.example.com');
  });

  it('never emits a host-reserved login URL for a reserved FRONTEND_URL', () => {
    for (const bad of [
      'https://mcp.example.com/admin/settings',
      'https://mcp.example.com/runtime',
    ]) {
      const url = `${resolveAuthUrlBase(bad, undefined, () => {})}/api/auth/login`;
      expect(url).toBe('https://mcp.example.com/api/auth/login');
      expect(reservedPrefixFor(url)).toBeNull();
    }
  });

  it('reports a reserved FRONTEND_URL even though the link now survives it', () => {
    const onProblem = jest.fn();
    resolveAuthUrlBase('https://mcp.example.com/runtime', undefined, onProblem);

    expect(onProblem).toHaveBeenCalledTimes(1);
    const message = onProblem.mock.calls[0][0] as string;
    expect(message).toContain('https://mcp.example.com/runtime');
    expect(message).toContain('/runtime');
    expect(message).toContain('frontend-url');
  });

  it('stays silent for a safe value', () => {
    const onProblem = jest.fn();
    expect(resolveAuthUrlBase('https://mcp.example.com', undefined, onProblem)).toBe(
      'https://mcp.example.com'
    );
    expect(onProblem).not.toHaveBeenCalled();
  });

  it('falls back to the OAUTH_REDIRECT_URI origin when FRONTEND_URL has none', () => {
    const redirect = 'https://mcp.example.com/api/auth/callback';
    expect(resolveAuthUrlBase(undefined, redirect, () => {})).toBe('https://mcp.example.com');
    expect(resolveAuthUrlBase('', redirect, () => {})).toBe('https://mcp.example.com');
    // `/` is what resolveFrontendUrl degrades to — site-relative, no origin.
    expect(resolveAuthUrlBase('/', redirect, () => {})).toBe('https://mcp.example.com');
    expect(resolveAuthUrlBase('/admin', redirect, () => {})).toBe('https://mcp.example.com');
  });

  it('returns an empty base when nothing carries an origin', () => {
    expect(resolveAuthUrlBase('', '', () => {})).toBe('');
    expect(resolveAuthUrlBase('/admin', '/api/auth/callback', () => {})).toBe('');
  });

  it('defaults to the process environment when called with no arguments', () => {
    const frontend = process.env.FRONTEND_URL;
    const redirect = process.env.OAUTH_REDIRECT_URI;
    try {
      process.env.FRONTEND_URL = 'https://mcp.example.com/admin/settings';
      delete process.env.OAUTH_REDIRECT_URI;
      expect(resolveAuthUrlBase(undefined, undefined, () => {})).toBe(
        'https://mcp.example.com'
      );
    } finally {
      if (frontend === undefined) delete process.env.FRONTEND_URL;
      else process.env.FRONTEND_URL = frontend;
      if (redirect === undefined) delete process.env.OAUTH_REDIRECT_URI;
      else process.env.OAUTH_REDIRECT_URI = redirect;
    }
  });
});
