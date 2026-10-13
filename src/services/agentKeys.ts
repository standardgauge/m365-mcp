import { randomBytes, timingSafeEqual } from 'crypto';
import { TableClient, TableEntity } from '@azure/data-tables';
import { getTableClient as getStorageTableClient } from './storageClient.js';
import { hashSessionToken } from './credentialCrypto.js';
import { ALL_SERVICE_KEYS, getAllowedSites } from './serviceSettings.js';

/**
 * Agent API keys: a credential a signed-in user mints for an agent or an app
 * they run, narrower than their own session.
 *
 * A key is never a new identity. It resolves to a session of the user who
 * minted it (a dedicated `client` session created alongside the key), so
 * every Graph call under it is still delegated and still bounded by what that
 * user can reach. On top of the user's and the tenant's policy, the key
 * carries its own: which services it may use, which of those are read-only
 * for it, and optionally which SharePoint sites it may see. Each of those can
 * only narrow. A Global Administrator sets a tenant-wide policy for keys
 * (`agentKeyPolicy` in `serviceSettings`) that no key can loosen, and both
 * policies are read on every request, so an admin's tightening applies at
 * once.
 *
 * Storage: table `agentKeys`, PartitionKey `key`, RowKey = HMAC(key)[:32],
 * the same shape as `mcpSessions`, so a lookup is one point read. The row
 * holds the full HMAC (compared on read), never the key itself. The key is
 * shown once, at creation.
 *
 * Revoking a key marks its row and deletes its session row, so a replica that
 * cached the session drops it within SESSION_REVALIDATE_MS, the same bound as
 * logout. Deleting the user's sessions (logout, purge) also ends every key.
 */

export const AGENT_KEY_PREFIX = 'm365k_';
const TABLE = 'agentKeys';
const PARTITION = 'key';
export const AGENT_KEY_LABEL_MAX = 64;
export const AGENT_KEY_MAX_LIFETIME_DAYS = 365;
/** How often lastUsedAt is written back. */
export const AGENT_KEY_TOUCH_INTERVAL_MS = 5 * 60 * 1000;

export interface AgentKeySitePick { id: string; name: string }

/** The policy a key is created with. Stored as JSON on the key row. */
export interface AgentKeyPolicy {
  /** Services the key may use. An allow-list; never wider than the user's own. */
  services: string[];
  /** Services whose write tools the key may not call, on top of the tenant's. */
  readOnlyServices: string[];
  /** SharePoint sites the key may see. Empty means the tenant's allow-list applies. */
  allowedSites: AgentKeySitePick[];
}

/** Tenant-wide constraints on every key, set by a Global Administrator. */
export interface TenantAgentKeyPolicy {
  /** When false no key authenticates and none can be created. */
  allowKeys: boolean;
  /** Services no key may use, whatever the key says. */
  disabledServices: string[];
  /** Services every key is read-only on, whatever the key says. */
  readOnlyServices: string[];
  /** Longest lifetime a user may give a key, in days. At most AGENT_KEY_MAX_LIFETIME_DAYS. */
  maxKeyLifetimeDays: number;
  /**
   * Whether an owner may widen one of their own keys after minting it. On by
   * default: a key names a policy row, not a baked-in grant, and editing in
   * place keeps the secret where it is. Off makes every edit narrow-only, for
   * tenants that want a wider key to mean a new key. Administrators can only
   * narrow another user's key whatever this says.
   */
  ownersMayWiden: boolean;
}

export const DEFAULT_TENANT_AGENT_KEY_POLICY: TenantAgentKeyPolicy = {
  allowKeys: true,
  disabledServices: [],
  readOnlyServices: [],
  maxKeyLifetimeDays: AGENT_KEY_MAX_LIFETIME_DAYS,
  ownersMayWiden: true,
};

/** What a request authenticated by a key carries on its session. */
export interface AgentKeyContext {
  keyId: string;
  label: string;
  /** The key's policy after the tenant policy has been applied. */
  policy: AgentKeyPolicy;
}

export interface AgentKeyRecord {
  keyId: string;
  tenantId: string;
  userId: string;
  userEmail: string;
  label: string;
  /** RowKey of the key's own session in `mcpSessions`. */
  sessionStorageKey: string;
  policy: AgentKeyPolicy;
  createdAt: number;
  /**
   * When the key stops working. Always set: the key's session carries this as
   * its own absolute cap in place of the 30-day one a sign-in gets, so a key
   * lives exactly as long as its owner chose, within the tenant's maximum.
   */
  expiresAt: number;
  lastUsedAt: number;
  /** 0 while the key is live. */
  revokedAt: number;
  revokedBy?: string;
  /** Graph id of the expiry-reminder event on the owner's calendar, when one was created. */
  reminderEventId?: string;
  /** When and by whom the policy or label was last edited after minting. */
  policyUpdatedAt?: number;
  policyUpdatedBy?: string;
  /** RowKey of the key's row. Set when loaded; not a column. */
  _storageKey?: string;
}

// ── Key material ─────────────────────────────────────────────────────────────

/** Mint a new key token. 32 random bytes, base64url, behind a fixed prefix. */
export function generateAgentKeyToken(): string {
  return AGENT_KEY_PREFIX + randomBytes(32).toString('base64url');
}

/** A short public identifier for a key; shown in lists and audit rows. */
export function generateAgentKeyId(): string {
  return randomBytes(6).toString('hex');
}

/** True when a bearer token is shaped like an agent key rather than a session token. */
export function looksLikeAgentKey(token: string): boolean {
  return token.startsWith(AGENT_KEY_PREFIX) && token.length > AGENT_KEY_PREFIX.length;
}

/** The device label the key's session carries, so audit rows name the key. */
export function agentKeyDeviceLabel(label: string, keyId: string): string {
  return `agent-key:${label}#${keyId}`;
}

// ── Policy arithmetic (pure) ─────────────────────────────────────────────────

function uniqueKnownServices(list: readonly string[]): string[] {
  return [...new Set(list)].filter((s) => ALL_SERVICE_KEYS.includes(s));
}

/**
 * Validate a policy a user asks for against what they may grant: the services
 * they themselves can use (tenant-enabled minus their own opt-outs minus the
 * tenant's key policy). Returns the normalised policy or a reason.
 */
export function validateRequestedPolicy(
  requested: Partial<AgentKeyPolicy> | undefined,
  grantable: readonly string[],
): { ok: true; policy: AgentKeyPolicy } | { ok: false; error: string } {
  if (!requested || !Array.isArray(requested.services) || requested.services.length === 0) {
    return { ok: false, error: 'services must be a non-empty array' };
  }
  if (requested.services.some((s) => typeof s !== 'string')) {
    return { ok: false, error: 'services must contain only strings' };
  }
  const unknown = requested.services.filter((s) => !ALL_SERVICE_KEYS.includes(s));
  if (unknown.length > 0) return { ok: false, error: `Unknown service(s): ${unknown.join(', ')}` };
  const notGrantable = requested.services.filter((s) => !grantable.includes(s));
  if (notGrantable.length > 0) {
    return { ok: false, error: `Service(s) not available to you, so not grantable to a key: ${notGrantable.join(', ')}` };
  }
  const services = uniqueKnownServices(requested.services);

  const ro = requested.readOnlyServices ?? [];
  if (!Array.isArray(ro) || ro.some((s) => typeof s !== 'string')) {
    return { ok: false, error: 'readOnlyServices must be an array of strings' };
  }
  const roOutside = ro.filter((s) => !services.includes(s));
  if (roOutside.length > 0) {
    return { ok: false, error: `readOnlyServices must be a subset of services: ${roOutside.join(', ')}` };
  }

  const sites = requested.allowedSites ?? [];
  if (!Array.isArray(sites)) return { ok: false, error: 'allowedSites must be an array' };
  const pickedSites: AgentKeySitePick[] = [];
  const seen = new Set<string>();
  for (const s of sites) {
    if (!s || typeof s !== 'object' || typeof (s as AgentKeySitePick).id !== 'string' || !(s as AgentKeySitePick).id) {
      return { ok: false, error: 'allowedSites entries must be { id, name }' };
    }
    const id = (s as AgentKeySitePick).id;
    if (seen.has(id)) continue;
    seen.add(id);
    const name = typeof (s as AgentKeySitePick).name === 'string' ? (s as AgentKeySitePick).name.slice(0, 200) : id;
    pickedSites.push({ id, name });
  }
  if (pickedSites.length > 0 && !services.includes('sharepoint')) {
    return { ok: false, error: 'allowedSites only applies when sharepoint is among the services' };
  }
  return { ok: true, policy: { services, readOnlyServices: uniqueKnownServices(ro), allowedSites: pickedSites } };
}

/**
 * Whether `next` grants no more than `prev`: fewer or equal services, every
 * read-only service still read-only, and a site list that is a subset
 * (an empty list means "the tenant's whole allow-list", the widest setting,
 * so going from empty to picks narrows and from picks to empty widens).
 */
export function isNarrowingOf(next: AgentKeyPolicy, prev: AgentKeyPolicy): boolean {
  if (next.services.some((s) => !prev.services.includes(s))) return false;
  for (const s of next.services) {
    if (prev.readOnlyServices.includes(s) && !next.readOnlyServices.includes(s)) return false;
  }
  if (prev.allowedSites.length === 0) return true;
  if (next.allowedSites.length === 0) return next.services.includes('sharepoint') ? false : true;
  const prevIds = new Set(prev.allowedSites.map((s) => s.id));
  return next.allowedSites.every((s) => prevIds.has(s.id));
}

/**
 * The policy in force for a key once the tenant's key policy is applied. The
 * result is never wider than the key's own policy.
 */
export function effectiveAgentKeyPolicy(key: AgentKeyPolicy, tenant: TenantAgentKeyPolicy): AgentKeyPolicy {
  const services = key.services.filter((s) => !tenant.disabledServices.includes(s));
  const readOnly = new Set<string>([...key.readOnlyServices, ...tenant.readOnlyServices]);
  return {
    services,
    readOnlyServices: services.filter((s) => readOnly.has(s)),
    allowedSites: key.allowedSites.map((s) => ({ ...s })),
  };
}

/** Whether a key-authenticated request may use `service` at all. */
export function keyAllowsService(ctx: AgentKeyContext | undefined, service: string): boolean {
  return !ctx || ctx.policy.services.includes(service);
}

/** Whether `service` is read-only for a key-authenticated request. */
export function keyIsReadOnly(ctx: AgentKeyContext | undefined, service: string): boolean {
  return !!ctx && ctx.policy.readOnlyServices.includes(service);
}

/**
 * The SharePoint site allow-list a key-authenticated request sees, given the
 * tenant's. An empty tenant list means the tenant allows every site, so the
 * key's picks are the whole list; otherwise it is the intersection by id,
 * and a key whose picks are all outside the tenant's list sees no site.
 * Without a key, or for a key with no picks, the tenant list stands.
 */
export function narrowAllowedSites(
  tenantAllowed: AgentKeySitePick[],
  ctx: AgentKeyContext | undefined,
): AgentKeySitePick[] {
  if (!ctx || ctx.policy.allowedSites.length === 0) return tenantAllowed;
  if (tenantAllowed.length === 0) return ctx.policy.allowedSites;
  const tenantIds = new Set(tenantAllowed.map((s) => s.id));
  const picked = ctx.policy.allowedSites.filter((s) => tenantIds.has(s.id));
  // A key restricted to sites the tenant no longer allows must not fall back to
  // the tenant's wider list. Return a sentinel that matches no site.
  return picked.length > 0 ? picked : [{ id: '\u0000no-site', name: 'No permitted site' }];
}

/**
 * The SharePoint allow-list in force for a request: the tenant's, narrowed by
 * the agent key's picks when the request carries one. Every place that reads
 * `getAllowedSites` for enforcement or filtering goes through this instead,
 * so a key cannot see a site through one surface that another hides.
 */
export async function effectiveAllowedSites(
  tenantId: string,
  session?: { agentKey?: AgentKeyContext } | null,
): Promise<AgentKeySitePick[]> {
  const tenantAllowed = await getAllowedSites(tenantId);
  return narrowAllowedSites(tenantAllowed, session?.agentKey);
}

// ── Tenant key policy (serviceSettings table) ────────────────────────────────

const SETTINGS_TABLE = 'serviceSettings';
const POLICY_ROW_KEY = 'agentKeyPolicy';

function settingsClient(): TableClient {
  return getStorageTableClient(SETTINGS_TABLE);
}

async function ensureTable(client: TableClient): Promise<void> {
  try {
    await client.createTable();
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status !== 409) throw err;
  }
}

export function normalizeTenantAgentKeyPolicy(raw: Partial<TenantAgentKeyPolicy> | null | undefined): TenantAgentKeyPolicy {
  const days = Number(raw?.maxKeyLifetimeDays);
  return {
    allowKeys: raw?.allowKeys !== false,
    ownersMayWiden: raw?.ownersMayWiden !== false,
    disabledServices: uniqueKnownServices(Array.isArray(raw?.disabledServices) ? raw.disabledServices : []),
    readOnlyServices: uniqueKnownServices(Array.isArray(raw?.readOnlyServices) ? raw.readOnlyServices : []),
    maxKeyLifetimeDays: Number.isInteger(days) && days >= 1 && days <= AGENT_KEY_MAX_LIFETIME_DAYS
      ? days
      : AGENT_KEY_MAX_LIFETIME_DAYS,
  };
}

export async function getTenantAgentKeyPolicy(tenantId: string): Promise<TenantAgentKeyPolicy> {
  const client = settingsClient();
  await ensureTable(client);
  try {
    const entity = await client.getEntity<{ policy: string }>(tenantId, POLICY_ROW_KEY);
    return normalizeTenantAgentKeyPolicy(JSON.parse(entity.policy) as Partial<TenantAgentKeyPolicy>);
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (status === 404) return normalizeTenantAgentKeyPolicy(DEFAULT_TENANT_AGENT_KEY_POLICY);
    throw err;
  }
}

export async function setTenantAgentKeyPolicy(tenantId: string, policy: TenantAgentKeyPolicy): Promise<void> {
  const client = settingsClient();
  await ensureTable(client);
  const entity: TableEntity = {
    partitionKey: tenantId,
    rowKey: POLICY_ROW_KEY,
    policy: JSON.stringify(normalizeTenantAgentKeyPolicy(policy)),
  };
  await client.upsertEntity(entity, 'Replace');
}

// ── Key rows ─────────────────────────────────────────────────────────────────

function keysClient(): TableClient {
  return getStorageTableClient(TABLE);
}

function isNotFound(err: unknown): boolean {
  return (err as { statusCode?: number } | null)?.statusCode === 404;
}

function rowKeyFor(token: string): { rowKey: string; hash: string } {
  const hash = hashSessionToken(token);
  return { rowKey: hash.slice(0, 32), hash };
}

function entityToRecord(entity: Record<string, unknown>): AgentKeyRecord {
  let policy: AgentKeyPolicy = { services: [], readOnlyServices: [], allowedSites: [] };
  try {
    const parsed = JSON.parse(String(entity.policy ?? '{}')) as Partial<AgentKeyPolicy>;
    policy = {
      services: Array.isArray(parsed.services) ? parsed.services : [],
      readOnlyServices: Array.isArray(parsed.readOnlyServices) ? parsed.readOnlyServices : [],
      allowedSites: Array.isArray(parsed.allowedSites) ? parsed.allowedSites : [],
    };
  } catch { /* an unreadable policy grants nothing */ }
  return {
    keyId: String(entity.keyId ?? ''),
    tenantId: String(entity.tenantId ?? ''),
    userId: String(entity.userId ?? ''),
    userEmail: String(entity.userEmail ?? ''),
    label: String(entity.label ?? ''),
    sessionStorageKey: String(entity.sessionStorageKey ?? ''),
    policy,
    createdAt: Number(entity.createdAt ?? 0),
    expiresAt: Number(entity.expiresAt ?? 0),
    lastUsedAt: Number(entity.lastUsedAt ?? 0),
    revokedAt: Number(entity.revokedAt ?? 0),
    revokedBy: entity.revokedBy ? String(entity.revokedBy) : undefined,
    reminderEventId: entity.reminderEventId ? String(entity.reminderEventId) : undefined,
    policyUpdatedAt: entity.policyUpdatedAt ? Number(entity.policyUpdatedAt) : undefined,
    policyUpdatedBy: entity.policyUpdatedBy ? String(entity.policyUpdatedBy) : undefined,
    _storageKey: String(entity.rowKey),
  };
}

export interface CreateAgentKeyInput {
  tenantId: string;
  userId: string;
  userEmail: string;
  label: string;
  sessionStorageKey: string;
  policy: AgentKeyPolicy;
  expiresAt: number;
  now?: number;
}

/**
 * Write a key row for a token the caller generated with generateAgentKeyToken.
 * The caller mints the key's session first and passes its RowKey.
 */
export async function createAgentKey(token: string, keyId: string, input: CreateAgentKeyInput): Promise<AgentKeyRecord> {
  const client = keysClient();
  await ensureTable(client);
  const { rowKey, hash } = rowKeyFor(token);
  const now = input.now ?? Date.now();
  const entity: TableEntity = {
    partitionKey: PARTITION,
    rowKey,
    keyHash: hash,
    keyId,
    tenantId: input.tenantId,
    userId: input.userId,
    userEmail: input.userEmail,
    label: input.label,
    sessionStorageKey: input.sessionStorageKey,
    policy: JSON.stringify(input.policy),
    createdAt: now,
    expiresAt: input.expiresAt,
    lastUsedAt: 0,
    revokedAt: 0,
  };
  // Create, never upsert: a collision on 32 hex chars of an HMAC is not a
  // thing that happens, and if it did the right answer is an error.
  await client.createEntity(entity);
  return entityToRecord({ ...entity });
}

/**
 * Resolve a presented key to its row. One point read; the stored HMAC is
 * compared in constant time. Null when no row has this key. Any other storage
 * failure propagates so the caller can refuse with a 503 rather than a 401.
 */
export async function loadAgentKeyByToken(token: string): Promise<AgentKeyRecord | null> {
  const client = keysClient();
  await ensureTable(client);
  const { rowKey, hash } = rowKeyFor(token);
  let entity: Record<string, unknown>;
  try {
    entity = await client.getEntity(PARTITION, rowKey);
  } catch (err) {
    if (isNotFound(err)) return null;
    throw err;
  }
  const stored = Buffer.from(String(entity.keyHash ?? ''), 'hex');
  const given = Buffer.from(hash, 'hex');
  if (stored.length !== given.length || !timingSafeEqual(stored, given)) return null;
  return entityToRecord(entity);
}

/** Every key in a tenant, or one user's. Revoked keys are included; callers filter. */
export async function listAgentKeys(tenantId: string, userId?: string): Promise<AgentKeyRecord[]> {
  const client = keysClient();
  await ensureTable(client);
  const filter = userId
    ? `PartitionKey eq '${PARTITION}' and tenantId eq '${escapeOData(tenantId)}' and userId eq '${escapeOData(userId)}'`
    : `PartitionKey eq '${PARTITION}' and tenantId eq '${escapeOData(tenantId)}'`;
  const out: AgentKeyRecord[] = [];
  for await (const entity of client.listEntities({ queryOptions: { filter } })) {
    out.push(entityToRecord(entity as Record<string, unknown>));
  }
  out.sort((a, b) => b.createdAt - a.createdAt);
  return out;
}

function escapeOData(v: string): string {
  return v.replace(/'/g, "''");
}

/** Find a key by its public id within a tenant. */
export async function findAgentKey(tenantId: string, keyId: string): Promise<AgentKeyRecord | null> {
  const client = keysClient();
  await ensureTable(client);
  const filter = `PartitionKey eq '${PARTITION}' and tenantId eq '${escapeOData(tenantId)}' and keyId eq '${escapeOData(keyId)}'`;
  for await (const entity of client.listEntities({ queryOptions: { filter } })) {
    return entityToRecord(entity as Record<string, unknown>);
  }
  return null;
}

/** Mark a key revoked. The caller deletes its session row. Idempotent. */
export async function markAgentKeyRevoked(record: AgentKeyRecord, byUserId: string, now = Date.now()): Promise<void> {
  if (!record._storageKey) throw new Error('agent key record has no storage key');
  const client = keysClient();
  await client.updateEntity({
    partitionKey: PARTITION,
    rowKey: record._storageKey,
    revokedAt: record.revokedAt || now,
    revokedBy: record.revokedBy ?? byUserId,
  }, 'Merge');
}

/** Replace a key's label and policy in place. Takes effect on the key's next request. */
export async function updateAgentKey(
  record: AgentKeyRecord,
  next: { label: string; policy: AgentKeyPolicy },
  byUserId: string,
  now = Date.now(),
): Promise<AgentKeyRecord> {
  if (!record._storageKey) throw new Error('agent key record has no storage key');
  const client = keysClient();
  await client.updateEntity({
    partitionKey: PARTITION,
    rowKey: record._storageKey,
    label: next.label,
    policy: JSON.stringify(next.policy),
    policyUpdatedAt: now,
    policyUpdatedBy: byUserId,
  }, 'Merge');
  return { ...record, label: next.label, policy: next.policy, policyUpdatedAt: now, policyUpdatedBy: byUserId };
}

/** Remember the calendar reminder created for a key, so revocation can remove it. */
export async function setAgentKeyReminder(record: AgentKeyRecord, eventId: string): Promise<void> {
  if (!record._storageKey) throw new Error('agent key record has no storage key');
  const client = keysClient();
  await client.updateEntity({ partitionKey: PARTITION, rowKey: record._storageKey, reminderEventId: eventId }, 'Merge');
  record.reminderEventId = eventId;
}

/** Record use. Best effort, fire and forget, rate-limited by the caller. */
export async function touchAgentKey(record: AgentKeyRecord, now = Date.now()): Promise<void> {
  if (!record._storageKey) return;
  const client = keysClient();
  await client.updateEntity({ partitionKey: PARTITION, rowKey: record._storageKey, lastUsedAt: now }, 'Merge');
}

/** The fields safe to show a user or an admin. Never the hash. */
export function publicAgentKey(record: AgentKeyRecord) {
  return {
    keyId: record.keyId,
    label: record.label,
    userId: record.userId,
    userEmail: record.userEmail,
    policy: record.policy,
    createdAt: record.createdAt,
    expiresAt: record.expiresAt,
    lastUsedAt: record.lastUsedAt || null,
    revokedAt: record.revokedAt || null,
    revokedBy: record.revokedBy ?? null,
    reminderSet: Boolean(record.reminderEventId),
    policyUpdatedAt: record.policyUpdatedAt ?? null,
    policyUpdatedBy: record.policyUpdatedBy ?? null,
  };
}
