/**
 * Unit tests for the UserEmailSettings service.
 *
 * Covers:
 *   - getUserEmailSettings: returns stored mode, defaults to 'draft' on 404
 *   - getUserEmailSettings: treats any stored value other than 'send' as 'draft'
 *   - setUserEmailSettings: upserts entity with the given mode
 *
 * Enforced draft mode:
 *   - the effective mode is 'draft' while a tenant or user policy applies, whatever is stored
 *   - tenant policy takes precedence over user policy
 *   - only a literal boolean true in the policy row enforces
 *   - a policy read that fails for anything but 404 propagates (fail closed)
 *   - getEmailOutputModeEnforcement / getEmailOutputModePolicy read the right rows
 *   - setEmailOutputModePolicy writes the right row with audit fields, for set and clear
 */

import { jest } from '@jest/globals';

const mockGetEntity = jest.fn<(table: string, partitionKey: string, rowKey: string) => Promise<unknown>>();
const mockUpsertEntity = jest.fn<(table: string, entity: unknown, mode: string) => Promise<void>>();
const mockCreateTable = jest.fn<(table: string) => Promise<void>>();

// The mock routes by table name so the settings table and the policy table can
// be scripted independently in one test.
jest.mock('@azure/data-tables', () => ({
  TableClient: {
    fromConnectionString: (_conn: string, table: string) => ({
      getEntity: (...args: unknown[]) => mockGetEntity(table, args[0] as string, args[1] as string),
      upsertEntity: (...args: unknown[]) => mockUpsertEntity(table, args[0], args[1] as string),
      createTable: () => mockCreateTable(table),
    }),
  },
}));

import {
  getUserEmailSettings,
  setUserEmailSettings,
  getEmailOutputModeEnforcement,
  getEmailOutputModePolicy,
  setEmailOutputModePolicy,
  TENANT_POLICY_ROW_KEY,
} from '../services/userEmailSettings.js';

const SETTINGS_TABLE = 'UserEmailSettings';
const POLICY_TABLE = 'EmailOutputModePolicy';
const TENANT = 'tenant-abc';
const USER = 'user-xyz';
const ADMIN = 'admin-123';
const NOT_FOUND = { statusCode: 404 };

/** Script the three reads getUserEmailSettings makes: stored preference, tenant policy, user policy. */
function script(opts: { stored?: unknown; tenantPolicy?: unknown; userPolicy?: unknown }) {
  mockGetEntity.mockImplementation(async (table, _pk, rk) => {
    const value = table === SETTINGS_TABLE
      ? opts.stored
      : rk === TENANT_POLICY_ROW_KEY
        ? opts.tenantPolicy
        : opts.userPolicy;
    if (value === undefined) throw NOT_FOUND;
    if (value instanceof Error || (typeof value === 'object' && value !== null && 'statusCode' in value)) throw value;
    return value;
  });
}

beforeEach(() => {
  jest.clearAllMocks();
  mockCreateTable.mockResolvedValue(undefined);
  mockUpsertEntity.mockResolvedValue(undefined);
  script({});
});

describe('getUserEmailSettings', () => {
  it('returns "draft" mode from stored entity', async () => {
    script({ stored: { emailOutputMode: 'draft' } });

    const result = await getUserEmailSettings(TENANT, USER);

    expect(result.emailOutputMode).toBe('draft');
    expect(mockGetEntity).toHaveBeenCalledWith(SETTINGS_TABLE, TENANT, USER);
  });

  it('returns "send" mode from stored entity', async () => {
    script({ stored: { emailOutputMode: 'send' } });

    const result = await getUserEmailSettings(TENANT, USER);

    expect(result).toEqual({
      emailOutputMode: 'send',
      preferredEmailOutputMode: 'send',
      enforced: false,
      enforcedBy: null,
    });
  });

  it('defaults to "draft" when no row exists (404)', async () => {
    const result = await getUserEmailSettings(TENANT, USER);

    expect(result.emailOutputMode).toBe('draft');
    expect(result.preferredEmailOutputMode).toBe('draft');
    expect(result.enforced).toBe(false);
  });

  it('treats unrecognised stored value as "draft"', async () => {
    script({ stored: { emailOutputMode: 'unknown-value' } });

    const result = await getUserEmailSettings(TENANT, USER);

    expect(result.emailOutputMode).toBe('draft');
  });

  it('propagates non-404 errors from the settings table', async () => {
    script({ stored: { statusCode: 500 } });

    await expect(getUserEmailSettings(TENANT, USER)).rejects.toEqual({ statusCode: 500 });
  });

  it('reads both policy rows of the tenant partition', async () => {
    await getUserEmailSettings(TENANT, USER);

    expect(mockGetEntity).toHaveBeenCalledWith(POLICY_TABLE, TENANT, TENANT_POLICY_ROW_KEY);
    expect(mockGetEntity).toHaveBeenCalledWith(POLICY_TABLE, TENANT, USER);
  });

  it('tolerates an existing table (409) and propagates other createTable failures', async () => {
    mockCreateTable.mockRejectedValue({ statusCode: 409 });
    await expect(getUserEmailSettings(TENANT, USER)).resolves.toBeDefined();

    mockCreateTable.mockRejectedValue({ statusCode: 503 });
    await expect(getUserEmailSettings(TENANT, USER)).rejects.toEqual({ statusCode: 503 });
  });
});

describe('getUserEmailSettings — enforced draft mode', () => {
  it('resolves to draft under a tenant policy even when the stored preference is send', async () => {
    script({ stored: { emailOutputMode: 'send' }, tenantPolicy: { enforceDraft: true } });

    const result = await getUserEmailSettings(TENANT, USER);

    expect(result).toEqual({
      emailOutputMode: 'draft',
      preferredEmailOutputMode: 'send',
      enforced: true,
      enforcedBy: 'tenant',
    });
  });

  it('resolves to draft under a per-user policy', async () => {
    script({ stored: { emailOutputMode: 'send' }, userPolicy: { enforceDraft: true } });

    const result = await getUserEmailSettings(TENANT, USER);

    expect(result.emailOutputMode).toBe('draft');
    expect(result.enforcedBy).toBe('user');
  });

  it('reports the tenant policy when both apply', async () => {
    script({ stored: { emailOutputMode: 'send' }, tenantPolicy: { enforceDraft: true }, userPolicy: { enforceDraft: true } });

    const result = await getUserEmailSettings(TENANT, USER);

    expect(result.enforcedBy).toBe('tenant');
  });

  it('is not enforced by a cleared policy row (enforceDraft=false)', async () => {
    script({ stored: { emailOutputMode: 'send' }, tenantPolicy: { enforceDraft: false, updatedBy: ADMIN } });

    const result = await getUserEmailSettings(TENANT, USER);

    expect(result.emailOutputMode).toBe('send');
    expect(result.enforced).toBe(false);
  });

  it('is not enforced by a non-boolean value — only a literal true counts', async () => {
    script({ stored: { emailOutputMode: 'send' }, tenantPolicy: { enforceDraft: 'true' }, userPolicy: { enforceDraft: 1 } });

    const result = await getUserEmailSettings(TENANT, USER);

    expect(result.enforced).toBe(false);
  });

  it('fails closed: a policy read that fails for anything but 404 propagates', async () => {
    script({ stored: { emailOutputMode: 'send' }, tenantPolicy: { statusCode: 500 } });

    await expect(getUserEmailSettings(TENANT, USER)).rejects.toEqual({ statusCode: 500 });
  });

  it('fails closed on a per-user policy read error too', async () => {
    script({ stored: { emailOutputMode: 'send' }, userPolicy: new Error('storage unreachable') });

    await expect(getUserEmailSettings(TENANT, USER)).rejects.toThrow('storage unreachable');
  });
});

describe('getEmailOutputModeEnforcement', () => {
  it('returns both policies with their audit fields and the resolution', async () => {
    script({
      tenantPolicy: { enforceDraft: false, updatedAt: 't1', updatedBy: ADMIN },
      userPolicy: { enforceDraft: true, updatedAt: 't2', updatedBy: ADMIN },
    });

    const result = await getEmailOutputModeEnforcement(TENANT, USER);

    expect(result).toEqual({
      enforced: true,
      enforcedBy: 'user',
      tenant: { enforceDraft: false, updatedAt: 't1', updatedBy: ADMIN },
      user: { enforceDraft: true, updatedAt: 't2', updatedBy: ADMIN },
    });
  });

  it('returns not-enforced defaults when neither row exists', async () => {
    const result = await getEmailOutputModeEnforcement(TENANT, USER);

    expect(result).toEqual({
      enforced: false,
      enforcedBy: null,
      tenant: { enforceDraft: false },
      user: { enforceDraft: false },
    });
  });

  it('drops non-string audit fields rather than passing them through', async () => {
    script({ tenantPolicy: { enforceDraft: true, updatedAt: 12345, updatedBy: null } });

    const result = await getEmailOutputModeEnforcement(TENANT, USER);

    expect(result.tenant).toEqual({ enforceDraft: true });
  });
});

describe('getEmailOutputModePolicy', () => {
  it('reads the tenant row for scope=tenant', async () => {
    script({ tenantPolicy: { enforceDraft: true } });

    const result = await getEmailOutputModePolicy(TENANT, { scope: 'tenant' });

    expect(result).toEqual({ enforceDraft: true });
    expect(mockGetEntity).toHaveBeenCalledWith(POLICY_TABLE, TENANT, TENANT_POLICY_ROW_KEY);
    expect(mockGetEntity).not.toHaveBeenCalledWith(POLICY_TABLE, TENANT, USER);
  });

  it('reads the user row for scope=user', async () => {
    script({ userPolicy: { enforceDraft: true } });

    const result = await getEmailOutputModePolicy(TENANT, { scope: 'user', userId: USER });

    expect(result).toEqual({ enforceDraft: true });
    expect(mockGetEntity).toHaveBeenCalledWith(POLICY_TABLE, TENANT, USER);
  });
});

describe('setEmailOutputModePolicy', () => {
  it('upserts the tenant row with audit fields and returns the policy', async () => {
    const before = Date.now();
    const result = await setEmailOutputModePolicy(TENANT, { scope: 'tenant' }, true, ADMIN);

    expect(result.enforceDraft).toBe(true);
    expect(result.updatedBy).toBe(ADMIN);
    expect(Date.parse(result.updatedAt as string)).toBeGreaterThanOrEqual(before);
    expect(mockUpsertEntity).toHaveBeenCalledWith(
      POLICY_TABLE,
      { partitionKey: TENANT, rowKey: TENANT_POLICY_ROW_KEY, enforceDraft: true, updatedAt: result.updatedAt, updatedBy: ADMIN },
      'Replace',
    );
  });

  it('upserts the user row for scope=user', async () => {
    await setEmailOutputModePolicy(TENANT, { scope: 'user', userId: USER }, true, ADMIN);

    expect(mockUpsertEntity).toHaveBeenCalledWith(
      POLICY_TABLE,
      expect.objectContaining({ partitionKey: TENANT, rowKey: USER, enforceDraft: true, updatedBy: ADMIN }),
      'Replace',
    );
  });

  it('clears by writing enforceDraft=false rather than deleting, keeping the audit trail', async () => {
    await setEmailOutputModePolicy(TENANT, { scope: 'tenant' }, false, ADMIN);

    expect(mockUpsertEntity).toHaveBeenCalledWith(
      POLICY_TABLE,
      expect.objectContaining({ rowKey: TENANT_POLICY_ROW_KEY, enforceDraft: false, updatedBy: ADMIN }),
      'Replace',
    );
  });

  it('propagates storage errors', async () => {
    mockUpsertEntity.mockRejectedValue(new Error('storage failure'));

    await expect(setEmailOutputModePolicy(TENANT, { scope: 'tenant' }, true, ADMIN)).rejects.toThrow('storage failure');
  });
});

describe('setUserEmailSettings', () => {
  it('upserts entity with emailOutputMode "send"', async () => {
    await setUserEmailSettings(TENANT, USER, { emailOutputMode: 'send' });

    expect(mockUpsertEntity).toHaveBeenCalledWith(
      SETTINGS_TABLE,
      { partitionKey: TENANT, rowKey: USER, emailOutputMode: 'send' },
      'Replace',
    );
  });

  it('upserts entity with emailOutputMode "draft"', async () => {
    await setUserEmailSettings(TENANT, USER, { emailOutputMode: 'draft' });

    expect(mockUpsertEntity).toHaveBeenCalledWith(
      SETTINGS_TABLE,
      { partitionKey: TENANT, rowKey: USER, emailOutputMode: 'draft' },
      'Replace',
    );
  });

  it('does not consult the policy — enforcement is the caller\'s check', async () => {
    await setUserEmailSettings(TENANT, USER, { emailOutputMode: 'send' });

    expect(mockGetEntity).not.toHaveBeenCalled();
  });

  it('propagates storage errors', async () => {
    mockUpsertEntity.mockRejectedValue(new Error('storage failure'));

    await expect(setUserEmailSettings(TENANT, USER, { emailOutputMode: 'draft' })).rejects.toThrow(
      'storage failure',
    );
  });
});
