/**
 * Per-user email output mode — controls whether AI-generated emails are
 * saved as drafts (for review) or sent immediately.
 *
 * Table: UserEmailSettings
 *   PartitionKey: tenantId
 *   RowKey:       userId
 *   emailOutputMode: 'draft' | 'send'
 *
 * Default for new users is 'draft', giving every user an opt-in review
 * step before any outbound email is actually delivered.
 *
 * Enforced draft mode
 * ----------------------------
 * The stored preference above is self-service: the user, or an agent acting
 * as the user through `set_email_output_mode`, can flip it. That means draft
 * mode on its own only guards against *unintended* sends. A prompt-injected
 * agent can call `set_email_output_mode('send')` and then `send_mail`, and
 * nothing stops it.
 *
 * A Global Admin can therefore pin the mode to 'draft' with a policy, either
 * tenant-wide or for one user. While a policy applies, the *effective* mode
 * is 'draft' whatever the stored preference says, and the self-service paths
 * (the MCP tool and POST /api/mail/settings) refuse to change it. Only the
 * admin policy endpoint clears it.
 *
 * Table: EmailOutputModePolicy
 *   PartitionKey: tenantId
 *   RowKey:       TENANT_POLICY_ROW_KEY for the tenant-wide policy, or userId
 *   enforceDraft: boolean
 *   updatedAt / updatedBy: audit trail of the last change
 *
 * Fail closed: a policy read that fails for any reason other than "no row"
 * throws, so the send paths and the mode-change paths that depend on it
 * error out rather than proceed as if nothing were enforced.
 */

import { TableClient, TableEntity } from '@azure/data-tables';
import { getTableClient as getStorageTableClient } from './storageClient.js';

const TABLE = 'UserEmailSettings';
const POLICY_TABLE = 'EmailOutputModePolicy';
/**
 * RowKey of the tenant-wide policy row. Entra user IDs are GUIDs, so this
 * can never collide with a per-user row in the same partition.
 */
export const TENANT_POLICY_ROW_KEY = '__tenant__';

export type EmailOutputMode = 'draft' | 'send';

/** Which policy pins the effective mode to 'draft'. */
export type EmailOutputModeEnforcedBy = 'tenant' | 'user';

export interface UserEmailSettings {
  /**
   * The effective mode: the stored preference, or 'draft' whenever a policy
   * enforces it. Every send path reads this field and nothing else, so the
   * policy applies to REST and MCP alike without each caller re-deriving it.
   */
  emailOutputMode: EmailOutputMode;
  /** The stored self-service preference, regardless of any policy. */
  preferredEmailOutputMode: EmailOutputMode;
  /** True when a tenant-wide or per-user policy pins the mode to 'draft'. */
  enforced: boolean;
  /** Which policy applies; null when none does. Tenant wins over user. */
  enforcedBy: EmailOutputModeEnforcedBy | null;
}

export interface EmailOutputModePolicy {
  enforceDraft: boolean;
  /** ISO-8601 timestamp of the last change */
  updatedAt?: string;
  /** userId of the admin who last changed the policy */
  updatedBy?: string;
}

export interface EmailOutputModeEnforcement {
  enforced: boolean;
  enforcedBy: EmailOutputModeEnforcedBy | null;
  tenant: EmailOutputModePolicy;
  user: EmailOutputModePolicy;
}

export type EmailOutputModePolicyScope =
  | { scope: 'tenant' }
  | { scope: 'user'; userId: string };

function getTableClient(table: string = TABLE): TableClient {
  return getStorageTableClient(table);
}

async function ensureTable(table: string = TABLE): Promise<void> {
  const client = getTableClient(table);
  try {
    await client.createTable();
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 409) throw err;
  }
}

function statusOf(err: unknown): number | undefined {
  return (err as { statusCode?: number }).statusCode;
}

const NO_POLICY: EmailOutputModePolicy = { enforceDraft: false };

/**
 * Read one policy row. A missing row means "not enforced"; any other failure
 * propagates so the caller fails closed.
 */
async function readPolicyRow(tenantId: string, rowKey: string): Promise<EmailOutputModePolicy> {
  const client = getTableClient(POLICY_TABLE);
  try {
    const entity = await client.getEntity<{ enforceDraft?: unknown; updatedAt?: unknown; updatedBy?: unknown }>(tenantId, rowKey);
    // Only a literal boolean true enforces. Anything else stored in the row
    // (or a row from another table shape) reads as "not enforced".
    const policy: EmailOutputModePolicy = { enforceDraft: entity.enforceDraft === true };
    if (typeof entity.updatedAt === 'string') policy.updatedAt = entity.updatedAt;
    if (typeof entity.updatedBy === 'string') policy.updatedBy = entity.updatedBy;
    return policy;
  } catch (err: unknown) {
    if (statusOf(err) === 404) return { ...NO_POLICY };
    throw err;
  }
}

/**
 * Resolve whether draft mode is enforced for a user, and by which policy.
 * The tenant-wide policy takes precedence over the per-user one so that an
 * admin who locks the tenant does not have to audit every user row.
 */
export async function getEmailOutputModeEnforcement(
  tenantId: string,
  userId: string,
): Promise<EmailOutputModeEnforcement> {
  await ensureTable(POLICY_TABLE);
  const [tenant, user] = await Promise.all([
    readPolicyRow(tenantId, TENANT_POLICY_ROW_KEY),
    readPolicyRow(tenantId, userId),
  ]);
  const enforcedBy: EmailOutputModeEnforcedBy | null = tenant.enforceDraft
    ? 'tenant'
    : user.enforceDraft
      ? 'user'
      : null;
  return { enforced: enforcedBy !== null, enforcedBy, tenant, user };
}

/**
 * Read one policy as stored, without resolving precedence. The admin UI uses
 * this to show the tenant-wide switch on its own.
 */
export async function getEmailOutputModePolicy(
  tenantId: string,
  target: EmailOutputModePolicyScope,
): Promise<EmailOutputModePolicy> {
  await ensureTable(POLICY_TABLE);
  return readPolicyRow(tenantId, target.scope === 'tenant' ? TENANT_POLICY_ROW_KEY : target.userId);
}

/**
 * Set or clear the draft-mode policy for the whole tenant or for one user.
 * Clearing writes `enforceDraft: false` rather than deleting the row, so the
 * audit fields record who cleared it and when.
 */
export async function setEmailOutputModePolicy(
  tenantId: string,
  target: EmailOutputModePolicyScope,
  enforceDraft: boolean,
  setBy: string,
): Promise<EmailOutputModePolicy> {
  await ensureTable(POLICY_TABLE);
  const client = getTableClient(POLICY_TABLE);
  const policy: EmailOutputModePolicy = {
    enforceDraft,
    updatedAt: new Date().toISOString(),
    updatedBy: setBy,
  };
  const entity: TableEntity = {
    partitionKey: tenantId,
    rowKey: target.scope === 'tenant' ? TENANT_POLICY_ROW_KEY : target.userId,
    ...policy,
  };
  await client.upsertEntity(entity, 'Replace');
  return policy;
}

/**
 * Read the user's stored preference, ignoring any policy.
 */
async function getPreferredEmailOutputMode(tenantId: string, userId: string): Promise<EmailOutputMode> {
  const client = getTableClient();
  try {
    const entity = await client.getEntity<{ emailOutputMode: string }>(tenantId, userId);
    return entity.emailOutputMode === 'send' ? 'send' : 'draft';
  } catch (err: unknown) {
    if (statusOf(err) === 404) return 'draft';
    throw err;
  }
}

/**
 * The user's email output settings with enforcement already applied.
 * `emailOutputMode` is the effective mode and is what every send path honours.
 */
export async function getUserEmailSettings(tenantId: string, userId: string): Promise<UserEmailSettings> {
  await ensureTable();
  const [preferredEmailOutputMode, enforcement] = await Promise.all([
    getPreferredEmailOutputMode(tenantId, userId),
    getEmailOutputModeEnforcement(tenantId, userId),
  ]);
  return {
    emailOutputMode: enforcement.enforced ? 'draft' : preferredEmailOutputMode,
    preferredEmailOutputMode,
    enforced: enforcement.enforced,
    enforcedBy: enforcement.enforcedBy,
  };
}

/**
 * Persist the user's self-service preference. This does not consult the
 * policy: callers that expose it to the user or to an agent (the MCP tool,
 * POST /api/mail/settings) check `enforced` first and refuse. Keeping the
 * check at the call site is what lets them log the refusal as denied with
 * the caller's identity attached.
 */
export async function setUserEmailSettings(
  tenantId: string,
  userId: string,
  settings: { emailOutputMode: EmailOutputMode },
): Promise<void> {
  await ensureTable();
  const client = getTableClient();
  const entity: TableEntity = {
    partitionKey: tenantId,
    rowKey: userId,
    emailOutputMode: settings.emailOutputMode,
  };
  await client.upsertEntity(entity, 'Replace');
}
