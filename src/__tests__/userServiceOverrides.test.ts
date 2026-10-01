/**
 * Unit tests for the UserServiceOverrides table service.
 *
 * Covers:
 *   - getUserServiceOverrides: returns parsed array, empty on 404
 *   - setUserServiceOverrides: upserts entity, deletes row on empty array
 *   - isServiceDisabledForUser: convenience check
 */

import { jest } from '@jest/globals';

// ── Mock the Azure Table Client ──────────────────────────────────────────────

const mockGetEntity = jest.fn<(partitionKey: string, rowKey: string) => Promise<unknown>>();
const mockUpsertEntity = jest.fn<(entity: unknown, mode: string) => Promise<void>>();
const mockDeleteEntity = jest.fn<(partitionKey: string, rowKey: string) => Promise<void>>();
const mockCreateTable = jest.fn<() => Promise<void>>();

jest.mock('@azure/data-tables', () => ({
  TableClient: {
    fromConnectionString: () => ({
      getEntity: (...args: unknown[]) => mockGetEntity(args[0] as string, args[1] as string),
      upsertEntity: (...args: unknown[]) => mockUpsertEntity(args[0], args[1] as string),
      deleteEntity: (...args: unknown[]) => mockDeleteEntity(args[0] as string, args[1] as string),
      createTable: () => mockCreateTable(),
    }),
  },
}));

// ── Import after mocks ──────────────────────────────────────────────────────

import {
  getUserServiceOverrides,
  setUserServiceOverrides,
  isServiceDisabledForUser,
} from '../services/userServiceOverrides.js';

// ── Fixtures ────────────────────────────────────────────────────────────────

const TENANT = 'tenant-abc';
const USER = 'user-xyz';

beforeEach(() => {
  jest.clearAllMocks();
  mockCreateTable.mockResolvedValue(undefined);
});

// ── getUserServiceOverrides ─────────────────────────────────────────────────

describe('getUserServiceOverrides', () => {
  it('returns parsed disabled services from the table', async () => {
    mockGetEntity.mockResolvedValue({ disabledServices: '["mail","calendar"]' });

    const result = await getUserServiceOverrides(TENANT, USER);

    expect(result).toEqual(['mail', 'calendar']);
    expect(mockGetEntity).toHaveBeenCalledWith(TENANT, USER);
  });

  it('returns empty array when no row exists (404)', async () => {
    mockGetEntity.mockRejectedValue({ statusCode: 404 });

    const result = await getUserServiceOverrides(TENANT, USER);

    expect(result).toEqual([]);
  });

  it('propagates non-404 errors', async () => {
    mockGetEntity.mockRejectedValue({ statusCode: 500 });

    await expect(getUserServiceOverrides(TENANT, USER)).rejects.toEqual({ statusCode: 500 });
  });
});

// ── setUserServiceOverrides ─────────────────────────────────────────────────

describe('setUserServiceOverrides', () => {
  it('upserts entity with JSON-encoded disabled services', async () => {
    mockUpsertEntity.mockResolvedValue(undefined);

    await setUserServiceOverrides(TENANT, USER, ['mail', 'onedrive']);

    expect(mockUpsertEntity).toHaveBeenCalledWith(
      {
        partitionKey: TENANT,
        rowKey: USER,
        disabledServices: '["mail","onedrive"]',
      },
      'Replace',
    );
  });

  it('deletes the row when disabledServices is empty (cleanup)', async () => {
    mockDeleteEntity.mockResolvedValue(undefined);

    await setUserServiceOverrides(TENANT, USER, []);

    expect(mockDeleteEntity).toHaveBeenCalledWith(TENANT, USER);
    expect(mockUpsertEntity).not.toHaveBeenCalled();
  });

  it('ignores 404 when deleting a row that does not exist', async () => {
    mockDeleteEntity.mockRejectedValue({ statusCode: 404 });

    // Should not throw
    await setUserServiceOverrides(TENANT, USER, []);

    expect(mockDeleteEntity).toHaveBeenCalledWith(TENANT, USER);
  });

  it('propagates non-404 errors on delete', async () => {
    mockDeleteEntity.mockRejectedValue({ statusCode: 500 });

    await expect(setUserServiceOverrides(TENANT, USER, [])).rejects.toEqual({ statusCode: 500 });
  });
});

// ── isServiceDisabledForUser ────────────────────────────────────────────────

describe('isServiceDisabledForUser', () => {
  it('returns true when the service is in the disabled list', async () => {
    mockGetEntity.mockResolvedValue({ disabledServices: '["mail","calendar"]' });

    expect(await isServiceDisabledForUser(TENANT, USER, 'mail')).toBe(true);
    expect(await isServiceDisabledForUser(TENANT, USER, 'calendar')).toBe(true);
  });

  it('returns false when the service is not in the disabled list', async () => {
    mockGetEntity.mockResolvedValue({ disabledServices: '["mail"]' });

    expect(await isServiceDisabledForUser(TENANT, USER, 'sharepoint')).toBe(false);
  });

  it('returns false when no overrides exist', async () => {
    mockGetEntity.mockRejectedValue({ statusCode: 404 });

    expect(await isServiceDisabledForUser(TENANT, USER, 'mail')).toBe(false);
  });
});
