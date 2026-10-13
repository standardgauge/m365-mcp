/**
 * Pure parts of the agent API key service (services/agentKeys.ts): key
 * shape, the policy a user may request, the tenant policy's effect on a key,
 * and how a key narrows the SharePoint allow-list.
 */
import { randomBytes } from 'crypto';

process.env.MCP_SESSION_HMAC_KEY = randomBytes(32).toString('hex');
process.env.MCP_DATA_ENCRYPTION_KEY = randomBytes(32).toString('hex');

import {
  AGENT_KEY_PREFIX,
  agentKeyDeviceLabel,
  effectiveAgentKeyPolicy,
  generateAgentKeyId,
  generateAgentKeyToken,
  keyAllowsService,
  keyIsReadOnly,
  looksLikeAgentKey,
  narrowAllowedSites,
  normalizeTenantAgentKeyPolicy,
  validateRequestedPolicy,
} from '../services/agentKeys.js';
import type { AgentKeyContext } from '../services/agentKeys.js';

describe('agent key tokens', () => {
  it('mints a prefixed, high-entropy token that is recognised as a key', () => {
    const t = generateAgentKeyToken();
    expect(t.startsWith(AGENT_KEY_PREFIX)).toBe(true);
    expect(t.length).toBeGreaterThan(AGENT_KEY_PREFIX.length + 40);
    expect(looksLikeAgentKey(t)).toBe(true);
    expect(generateAgentKeyToken()).not.toEqual(t);
  });

  it('does not mistake a session token or the bare prefix for a key', () => {
    expect(looksLikeAgentKey(randomBytes(32).toString('hex'))).toBe(false);
    expect(looksLikeAgentKey(AGENT_KEY_PREFIX)).toBe(false);
    expect(looksLikeAgentKey('')).toBe(false);
  });

  it('gives each key a short public id and a device label that names it', () => {
    const id = generateAgentKeyId();
    expect(id).toMatch(/^[0-9a-f]{12}$/);
    expect(agentKeyDeviceLabel('nightly report', id)).toBe(`agent-key:nightly report#${id}`);
  });
});

describe('validateRequestedPolicy', () => {
  const grantable = ['mail', 'sharepoint', 'calendar'];

  it('accepts a subset of the grantable services and normalises it', () => {
    const r = validateRequestedPolicy({ services: ['sharepoint', 'sharepoint', 'mail'], readOnlyServices: ['mail'] }, grantable);
    expect(r).toEqual({ ok: true, policy: { services: ['sharepoint', 'mail'], readOnlyServices: ['mail'], allowedSites: [] } });
  });

  it('refuses an empty or missing service list', () => {
    expect(validateRequestedPolicy({ services: [] }, grantable)).toMatchObject({ ok: false });
    expect(validateRequestedPolicy(undefined, grantable)).toMatchObject({ ok: false });
  });

  it('refuses a service the user cannot reach themselves', () => {
    const r = validateRequestedPolicy({ services: ['teams'] }, grantable);
    expect(r).toMatchObject({ ok: false, error: expect.stringContaining('teams') });
  });

  it('refuses an unknown service name', () => {
    expect(validateRequestedPolicy({ services: ['exchange'] }, grantable)).toMatchObject({ ok: false, error: expect.stringContaining('Unknown') });
  });

  it('requires readOnlyServices to be within services', () => {
    const r = validateRequestedPolicy({ services: ['mail'], readOnlyServices: ['calendar'] }, grantable);
    expect(r).toMatchObject({ ok: false, error: expect.stringContaining('subset') });
  });

  it('accepts site picks only alongside sharepoint, and de-duplicates them', () => {
    const sites = [{ id: 's1', name: 'Finance' }, { id: 's1', name: 'Finance' }, { id: 's2', name: 'HR' }];
    expect(validateRequestedPolicy({ services: ['sharepoint'], allowedSites: sites }, grantable)).toEqual({
      ok: true,
      policy: { services: ['sharepoint'], readOnlyServices: [], allowedSites: [{ id: 's1', name: 'Finance' }, { id: 's2', name: 'HR' }] },
    });
    expect(validateRequestedPolicy({ services: ['mail'], allowedSites: sites }, grantable)).toMatchObject({ ok: false });
    expect(validateRequestedPolicy({ services: ['sharepoint'], allowedSites: [{ name: 'no id' } as unknown as { id: string; name: string }] }, grantable)).toMatchObject({ ok: false });
  });
});

describe('effectiveAgentKeyPolicy', () => {
  const key = { services: ['mail', 'sharepoint', 'calendar'], readOnlyServices: ['calendar'], allowedSites: [{ id: 's1', name: 'A' }] };

  it('is the key policy itself under the default tenant policy', () => {
    expect(effectiveAgentKeyPolicy(key, normalizeTenantAgentKeyPolicy(null))).toEqual(key);
  });

  it('drops services the tenant disables for keys and adds the tenant read-only set', () => {
    const r = effectiveAgentKeyPolicy(key, { allowKeys: true, disabledServices: ['mail'], readOnlyServices: ['sharepoint', 'teams'], maxKeyLifetimeDays: 365 });
    expect(r.services).toEqual(['sharepoint', 'calendar']);
    expect(r.readOnlyServices.sort()).toEqual(['calendar', 'sharepoint']);
  });

  it('never widens: a tenant policy cannot add a service the key lacks', () => {
    const r = effectiveAgentKeyPolicy({ services: ['mail'], readOnlyServices: [], allowedSites: [] }, { allowKeys: true, disabledServices: [], readOnlyServices: [], maxKeyLifetimeDays: 365 });
    expect(r.services).toEqual(['mail']);
  });

  it('answers the per-request questions', () => {
    const ctx: AgentKeyContext = { keyId: 'k', label: 'l', policy: { services: ['mail'], readOnlyServices: ['mail'], allowedSites: [] } };
    expect(keyAllowsService(ctx, 'mail')).toBe(true);
    expect(keyAllowsService(ctx, 'sharepoint')).toBe(false);
    expect(keyAllowsService(undefined, 'sharepoint')).toBe(true);
    expect(keyIsReadOnly(ctx, 'mail')).toBe(true);
    expect(keyIsReadOnly(undefined, 'mail')).toBe(false);
  });
});

describe('narrowAllowedSites', () => {
  const tenant = [{ id: 's1', name: 'A' }, { id: 's2', name: 'B' }];
  const ctx = (sites: Array<{ id: string; name: string }>): AgentKeyContext => ({ keyId: 'k', label: 'l', policy: { services: ['sharepoint'], readOnlyServices: [], allowedSites: sites } });

  it('leaves the tenant list alone without a key, or for a key with no picks', () => {
    expect(narrowAllowedSites(tenant, undefined)).toBe(tenant);
    expect(narrowAllowedSites(tenant, ctx([]))).toBe(tenant);
  });

  it("uses the key's picks when the tenant allows every site", () => {
    expect(narrowAllowedSites([], ctx([{ id: 's9', name: 'Z' }]))).toEqual([{ id: 's9', name: 'Z' }]);
  });

  it('intersects by id when both lists are set', () => {
    expect(narrowAllowedSites(tenant, ctx([{ id: 's2', name: 'B' }, { id: 's9', name: 'Z' }]))).toEqual([{ id: 's2', name: 'B' }]);
  });

  it('matches no site, rather than falling back to the tenant list, when every pick is outside it', () => {
    const r = narrowAllowedSites(tenant, ctx([{ id: 's9', name: 'Z' }]));
    expect(r).toHaveLength(1);
    expect(tenant.some((s) => s.id === r[0].id)).toBe(false);
  });
});

describe('normalizeTenantAgentKeyPolicy', () => {
  it('defaults to keys allowed with no restrictions and a one-year maximum lifetime', () => {
    expect(normalizeTenantAgentKeyPolicy(undefined)).toEqual({ allowKeys: true, disabledServices: [], readOnlyServices: [], maxKeyLifetimeDays: 365 });
  });
  it('keeps only known service keys', () => {
    expect(normalizeTenantAgentKeyPolicy({ allowKeys: false, disabledServices: ['mail', 'bogus'], readOnlyServices: ['teams', 'teams'], maxKeyLifetimeDays: 365 }))
      .toEqual({ allowKeys: false, disabledServices: ['mail'], readOnlyServices: ['teams'], maxKeyLifetimeDays: 365 });
  });
  it('keeps a valid maximum lifetime and falls back to a year otherwise', () => {
    expect(normalizeTenantAgentKeyPolicy({ maxKeyLifetimeDays: 30 }).maxKeyLifetimeDays).toBe(30);
    expect(normalizeTenantAgentKeyPolicy({ maxKeyLifetimeDays: 0 }).maxKeyLifetimeDays).toBe(365);
    expect(normalizeTenantAgentKeyPolicy({ maxKeyLifetimeDays: 9999 }).maxKeyLifetimeDays).toBe(365);
    expect(normalizeTenantAgentKeyPolicy({ maxKeyLifetimeDays: 2.5 }).maxKeyLifetimeDays).toBe(365);
  });
});
