/**
 * Tests for services/readiness.ts, the post-deploy operator checks the admin
 * console reports.
 *
 * Covers:
 *   - both flags unset: both checks open, report not ready
 *   - both flags `true`: both checks enforced, report ready
 *   - one flag set: only that check is enforced, report not ready
 *   - `enforced` follows the storage layer's parse (case-insensitive `true`,
 *     anything else is open) and `value` shows the raw setting beside it
 */

import { collectReadinessChecks } from '../services/readiness.js';

const FLAGS = ['MCP_ENVELOPE_REQUIRE_AAD', 'MCP_SESSION_REQUIRE_IDENTITY_BINDING'] as const;
const saved: Record<string, string | undefined> = {};

beforeEach(() => {
  for (const flag of FLAGS) {
    saved[flag] = process.env[flag];
    delete process.env[flag];
  }
});

afterEach(() => {
  for (const flag of FLAGS) {
    const value = saved[flag];
    if (value === undefined) delete process.env[flag];
    else process.env[flag] = value;
  }
});

function byId(id: string) {
  const check = collectReadinessChecks().checks.find((c) => c.id === id);
  if (!check) throw new Error(`no check ${id}`);
  return check;
}

describe('collectReadinessChecks', () => {
  it('reports both checks open and not ready when neither flag is set', () => {
    const report = collectReadinessChecks();
    expect(report.ready).toBe(false);
    expect(report.checks.map((c) => [c.id, c.setting, c.value, c.enforced])).toEqual([
      ['envelope-row-binding', 'MCP_ENVELOPE_REQUIRE_AAD', null, false],
      ['session-identity-binding', 'MCP_SESSION_REQUIRE_IDENTITY_BINDING', null, false],
    ]);
  });

  it('reports ready only when both flags are true', () => {
    process.env.MCP_ENVELOPE_REQUIRE_AAD = 'true';
    process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING = 'true';
    const report = collectReadinessChecks();
    expect(report.ready).toBe(true);
    expect(report.checks.every((c) => c.enforced)).toBe(true);
  });

  it('keeps the report not ready while only one flag is set', () => {
    process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING = 'true';
    const report = collectReadinessChecks();
    expect(report.ready).toBe(false);
    expect(byId('session-identity-binding').enforced).toBe(true);
    expect(byId('envelope-row-binding').enforced).toBe(false);
  });

  it('follows the storage layer parse: TRUE enforces, a near miss stays open with its raw value shown', () => {
    process.env.MCP_ENVELOPE_REQUIRE_AAD = 'TRUE';
    process.env.MCP_SESSION_REQUIRE_IDENTITY_BINDING = 'yes';
    expect(byId('envelope-row-binding')).toMatchObject({ value: 'TRUE', enforced: true });
    expect(byId('session-identity-binding')).toMatchObject({ value: 'yes', enforced: false });
  });

  it('names what is at stake while a check is open', () => {
    expect(byId('envelope-row-binding').open).toMatch(/2\.2/);
    expect(byId('session-identity-binding').open).toMatch(/2\.3/);
  });
});
