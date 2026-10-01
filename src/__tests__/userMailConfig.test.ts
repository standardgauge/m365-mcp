/**
 * Unit tests for the UserMailConfig table service.
 *
 * Covers:
 *   - getMailConfig: returns parsed config, defaults on 404, propagates other errors
 *   - setMailConfig: upserts with merged fields + audit metadata
 *   - isMailIndexingDisabled: convenience flag read; fails closed on storage error
 */

import { jest } from '@jest/globals';

// ── Mock the Azure Table Client ──────────────────────────────────────────────

const mockGetEntity = jest.fn<(partitionKey: string, rowKey: string) => Promise<unknown>>();
const mockUpsertEntity = jest.fn<(entity: unknown, mode: string) => Promise<void>>();
const mockCreateTable = jest.fn<() => Promise<void>>();

jest.mock('@azure/data-tables', () => ({
  TableClient: {
    fromConnectionString: () => ({
      getEntity: (...args: unknown[]) => mockGetEntity(args[0] as string, args[1] as string),
      upsertEntity: (...args: unknown[]) => mockUpsertEntity(args[0], args[1] as string),
      createTable: () => mockCreateTable(),
    }),
  },
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import {
  getMailConfig,
  setMailConfig,
  isMailIndexingDisabled,
} from '../services/userMailConfig.js';

// ── Fixtures ────────────────────────────────────────────────────────────────

const TENANT = 'tenant-abc';
const USER = 'user-xyz';
const ADMIN = 'admin-123';

beforeEach(() => {
  jest.clearAllMocks();
  mockCreateTable.mockResolvedValue(undefined);
});

// ── getMailConfig ────────────────────────────────────────────────────────────

describe('getMailConfig', () => {
  it('returns parsed config when row exists', async () => {
    mockGetEntity.mockResolvedValue({
      config: JSON.stringify({ disable_mail_indexing: true, updated_at: '2026-01-01T00:00:00.000Z', updated_by: ADMIN }),
    });

    const result = await getMailConfig(TENANT, USER);

    expect(result.disable_mail_indexing).toBe(true);
    expect(result.updated_by).toBe(ADMIN);
    expect(mockGetEntity).toHaveBeenCalledWith(TENANT, USER);
  });

  it('returns default config (flag=false) when row does not exist (404)', async () => {
    mockGetEntity.mockRejectedValue({ statusCode: 404 });

    const result = await getMailConfig(TENANT, USER);

    expect(result).toEqual({ disable_mail_indexing: false });
  });

  it('propagates non-404 errors', async () => {
    mockGetEntity.mockRejectedValue({ statusCode: 500 });

    await expect(getMailConfig(TENANT, USER)).rejects.toEqual({ statusCode: 500 });
  });
});

// ── setMailConfig ────────────────────────────────────────────────────────────

describe('setMailConfig', () => {
  it('upserts config with audit metadata when no prior row exists', async () => {
    mockGetEntity.mockRejectedValue({ statusCode: 404 });
    mockUpsertEntity.mockResolvedValue(undefined);

    await setMailConfig(TENANT, USER, { disable_mail_indexing: true }, ADMIN);

    expect(mockUpsertEntity).toHaveBeenCalledWith(
      expect.objectContaining({
        partitionKey: TENANT,
        rowKey: USER,
      }),
      'Replace',
    );

    // Verify JSON payload contains the flag + audit fields
    const call = (mockUpsertEntity.mock.calls[0] as [{ config: string }, string])[0];
    const stored = JSON.parse(call.config);
    expect(stored.disable_mail_indexing).toBe(true);
    expect(stored.updated_by).toBe(ADMIN);
    expect(stored.updated_at).toBeTruthy();
  });

  it('merges with existing config — does not overwrite unrelated fields', async () => {
    mockGetEntity.mockResolvedValue({
      config: JSON.stringify({ disable_mail_indexing: false, custom_field: 'keep-me' }),
    });
    mockUpsertEntity.mockResolvedValue(undefined);

    await setMailConfig(TENANT, USER, { disable_mail_indexing: true }, ADMIN);

    const call = (mockUpsertEntity.mock.calls[0] as [{ config: string }, string])[0];
    const stored = JSON.parse(call.config);
    expect(stored.disable_mail_indexing).toBe(true);
    expect(stored.custom_field).toBe('keep-me');
    expect(stored.updated_by).toBe(ADMIN);
  });

  it('can clear the flag (set to false)', async () => {
    mockGetEntity.mockResolvedValue({
      config: JSON.stringify({ disable_mail_indexing: true }),
    });
    mockUpsertEntity.mockResolvedValue(undefined);

    await setMailConfig(TENANT, USER, { disable_mail_indexing: false }, ADMIN);

    const call = (mockUpsertEntity.mock.calls[0] as [{ config: string }, string])[0];
    const stored = JSON.parse(call.config);
    expect(stored.disable_mail_indexing).toBe(false);
    expect(stored.updated_by).toBe(ADMIN);
  });
});

// ── isMailIndexingDisabled ───────────────────────────────────────────────────

describe('isMailIndexingDisabled', () => {
  it('returns true when flag is set', async () => {
    mockGetEntity.mockResolvedValue({
      config: JSON.stringify({ disable_mail_indexing: true }),
    });

    expect(await isMailIndexingDisabled(TENANT, USER)).toBe(true);
  });

  it('returns false when flag is not set', async () => {
    mockGetEntity.mockResolvedValue({
      config: JSON.stringify({ disable_mail_indexing: false }),
    });

    expect(await isMailIndexingDisabled(TENANT, USER)).toBe(false);
  });

  it('returns false when no config row exists (404)', async () => {
    mockGetEntity.mockRejectedValue({ statusCode: 404 });

    expect(await isMailIndexingDisabled(TENANT, USER)).toBe(false);
  });

  it('fails closed — returns true on storage error', async () => {
    mockGetEntity.mockRejectedValue({ statusCode: 503, message: 'Service unavailable' });

    expect(await isMailIndexingDisabled(TENANT, USER)).toBe(true);
  });
});
