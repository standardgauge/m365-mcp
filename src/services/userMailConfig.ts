/**
 * Per-user mail config — supports the disable_mail_indexing flag for
 * IR-adjacent roles whose entire mailbox is composed of investor PII,
 * K-1s, sub docs, or wire instructions.
 *
 * Table: UserMailConfig
 *   PartitionKey: tenantId
 *   RowKey:       userId
 *   config:       JSON-serialised UserMailConfig object
 *
 * Semantics:
 *   - disable_mail_indexing: true short-circuits ALL mail tool calls for
 *     the user before any folder-level deny-list enforcement runs.
 *   - Only Global Admins may set or unset the flag via the admin API.
 *   - Missing row or missing field is treated as flag=false (mail allowed).
 *   - Fail closed: if storage is unreachable the call is denied, not
 *     allowed through (consistent with userServiceOverrides behaviour).
 */

import { TableClient, TableEntity } from '@azure/data-tables';

const TABLE = 'UserMailConfig';
const AZURITE_CONNECTION_STRING = 'UseDevelopmentStorage=true';

export interface UserMailConfig {
  disable_mail_indexing: boolean;
  /** ISO-8601 timestamp when the flag was last set/cleared */
  updated_at?: string;
  /** userId of the admin who last changed the flag */
  updated_by?: string;
}

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
 * Retrieve the full mail config for a user.
 * Returns defaults (flag=false) when no row exists.
 */
export async function getMailConfig(tenantId: string, userId: string): Promise<UserMailConfig> {
  await ensureTable();
  const client = getTableClient();
  try {
    const entity = await client.getEntity<{ config: string }>(tenantId, userId);
    return JSON.parse(entity.config) as UserMailConfig;
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status === 404) return { disable_mail_indexing: false };
    throw err;
  }
}

/**
 * Persist a mail config update for a user.
 * Always stamps updated_at and updated_by on the stored record.
 */
export async function setMailConfig(
  tenantId: string,
  userId: string,
  patch: Partial<UserMailConfig>,
  setBy: string,
): Promise<void> {
  await ensureTable();
  const client = getTableClient();

  // Read-modify-write so we don't overwrite unrelated fields.
  let current: UserMailConfig;
  try {
    const entity = await client.getEntity<{ config: string }>(tenantId, userId);
    current = JSON.parse(entity.config) as UserMailConfig;
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status === 404) {
      current = { disable_mail_indexing: false };
    } else {
      throw err;
    }
  }

  const updated: UserMailConfig = {
    ...current,
    ...patch,
    updated_at: new Date().toISOString(),
    updated_by: setBy,
  };

  const entity: TableEntity = {
    partitionKey: tenantId,
    rowKey: userId,
    config: JSON.stringify(updated),
  };
  await client.upsertEntity(entity, 'Replace');
}

/**
 * Returns true when the user's mailbox indexing is fully disabled.
 * Fails closed — returns true (deny) on storage unavailability.
 */
export async function isMailIndexingDisabled(tenantId: string, userId: string): Promise<boolean> {
  try {
    const cfg = await getMailConfig(tenantId, userId);
    return cfg.disable_mail_indexing === true;
  } catch {
    // Fail closed: treat storage errors as "deny mail access"
    return true;
  }
}
