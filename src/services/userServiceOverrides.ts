/**
 * Per-user service overrides — allows Global Admins to disable specific
 * services for individual users while leaving them enabled for the rest
 * of the tenant.
 *
 * Table: UserServiceOverrides
 *   PartitionKey: tenantId
 *   RowKey:       userId
 *   disabledServices: JSON string[] (e.g. '["mail","calendar"]')
 *
 * Semantics:
 *   - Tenant-level enabledServices is checked first; if a service is
 *     disabled at tenant level, user overrides are irrelevant.
 *   - User overrides can only *disable* services that the tenant has
 *     enabled — they cannot re-enable tenant-disabled services.
 *   - An empty disabledServices array (or missing row) means no
 *     per-user restrictions.
 */

import { TableClient, TableEntity } from '@azure/data-tables';

const TABLE = 'UserServiceOverrides';
const AZURITE_CONNECTION_STRING = 'UseDevelopmentStorage=true';

function getTableClient(): TableClient {
  const conn = process.env.AZURE_STORAGE_CONNECTION_STRING ?? AZURITE_CONNECTION_STRING;
  const allowInsecureConnection =
    conn === 'UseDevelopmentStorage=true' || conn.includes('DefaultEndpointsProtocol=http;');
  return TableClient.fromConnectionString(conn, TABLE, { allowInsecureConnection });
}

async function ensureTable(): Promise<void> {
  const client = getTableClient();
  try {
    await client.createTable();
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 409) throw err;
  }
}

/**
 * Get the list of services disabled for a specific user.
 * Returns an empty array if no overrides exist (no restrictions).
 */
export async function getUserServiceOverrides(tenantId: string, userId: string): Promise<string[]> {
  await ensureTable();
  const client = getTableClient();
  try {
    const entity = await client.getEntity<{ disabledServices: string }>(tenantId, userId);
    return JSON.parse(entity.disabledServices) as string[];
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status === 404) return [];
    throw err;
  }
}

/**
 * Set the list of disabled services for a specific user.
 * Pass an empty array to clear all overrides.
 */
export async function setUserServiceOverrides(
  tenantId: string,
  userId: string,
  disabledServices: string[],
): Promise<void> {
  await ensureTable();
  const client = getTableClient();

  if (disabledServices.length === 0) {
    // Clean up: delete the row when there are no overrides
    try {
      await client.deleteEntity(tenantId, userId);
    } catch (err: unknown) {
      const status = (err as { statusCode?: number }).statusCode;
      if (status !== 404) throw err;
    }
    return;
  }

  const entity: TableEntity = {
    partitionKey: tenantId,
    rowKey: userId,
    disabledServices: JSON.stringify(disabledServices),
  };
  await client.upsertEntity(entity, 'Replace');
}

/**
 * Check whether a specific service is disabled for a user.
 * Returns true if the service is in the user's disabledServices list.
 */
export async function isServiceDisabledForUser(
  tenantId: string,
  userId: string,
  service: string,
): Promise<boolean> {
  const disabled = await getUserServiceOverrides(tenantId, userId);
  return disabled.includes(service);
}
