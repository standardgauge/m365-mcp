import { TableClient, TableServiceClient } from '@azure/data-tables';
import { createHash } from 'crypto';
import {
  hashSessionToken,
  encryptWithDek,
  decryptWithDek,
} from './credentialCrypto.js';

const connectionString = process.env.AZURE_STORAGE_CONNECTION_STRING;

let sessionsTable: TableClient | null = null;
let msalCacheTable: TableClient | null = null;
let installNoncesTable: TableClient | null = null;
let initialized = false;

async function ensureTables(): Promise<void> {
  if (initialized || !connectionString) return;

  const serviceClient = TableServiceClient.fromConnectionString(connectionString);

  // Create tables if they don't exist
  try { await serviceClient.createTable('mcpSessions'); } catch { /* already exists */ }
  try { await serviceClient.createTable('mcpMsalCache'); } catch { /* already exists */ }
  try { await serviceClient.createTable('mcpInstallNonces'); } catch { /* already exists */ }

  sessionsTable = TableClient.fromConnectionString(connectionString, 'mcpSessions');
  msalCacheTable = TableClient.fromConnectionString(connectionString, 'mcpMsalCache');
  installNoncesTable = TableClient.fromConnectionString(connectionString, 'mcpInstallNonces');
  initialized = true;
}

function getSessionsTable(): TableClient {
  if (!sessionsTable) throw new Error('Table storage not initialized');
  return sessionsTable;
}

function getMsalCacheTable(): TableClient {
  if (!msalCacheTable) throw new Error('Table storage not initialized');
  return msalCacheTable;
}

function getInstallNoncesTable(): TableClient {
  if (!installNoncesTable) throw new Error('Table storage not initialized');
  return installNoncesTable;
}

// ── Session persistence ──
//
// Storage schema (mcpSessions table) — multi-session support:
//   PartitionKey       = 'session'
//   RowKey             = tokenHash[:32] (first 32 chars of HMAC-SHA256 hash)
//                        This allows multiple sessions per user (different
//                        devices, browser vs CLI, etc.) to coexist.
//   userId             = Microsoft Graph user object id (column, not RowKey)
//   homeAccountId      = MSAL home account id (plaintext, not a credential)
//   displayName, email = profile fields (plaintext, not credentials)
//   tenantId           = Entra tenant id (plaintext)
//   expiresAt          = unix ms (plaintext)
//
// Hardened columns (added in, replace earlier plaintext columns):
//   sessionTokenHash         = HMAC-SHA256(MCP_SESSION_HMAC_KEY, sessionToken)
//                              The lookup key for Bearer-token auth. Stolen
//                              storage data is useless without the HMAC key.
//   accessTokenCiphertext    = AES-256-GCM ciphertext of the Graph access token
//   accessTokenIv            = 12-byte IV (base64)
//   accessTokenAuthTag       = 16-byte GCM auth tag (base64)
//
// Backward compatibility: legacy rows use RowKey = userId. These are read
// transparently but new sessions always use the tokenHash RowKey format.
//
// We do NOT store the original sessionToken anywhere — it lives only in the
// client config files after install. Loss of the HMAC key means every user
// must re-OAuth, but that's the desired property: the secret-at-rest material
// is recoverable only via Key Vault, not via storage exfiltration.

export interface StoredSession {
  userId: string;
  homeAccountId: string;
  displayName: string;
  email: string;
  tenantId: string;
  /** Plaintext Microsoft Graph access token. Encrypted at rest. */
  accessToken: string;
  expiresAt: number;
  /**
   * Plaintext crypto-random session token — the Bearer credential clients send.
   * Hashed (HMAC) at rest; never stored as plaintext.
   */
  sessionToken: string;
  /** Unix timestamp (ms) when the session was created. Used for server-side expiry. */
  sessionCreatedAt: number;
  /** Immutable Unix timestamp (ms) when this session was first created via OAuth. Never reset. */
  sessionAbsoluteCreatedAt?: number;
  /** Human-readable label for the device/client that created this session. */
  deviceLabel?: string;
  /**
   * Table Storage RowKey for this session. Used by the refresh path to update
   * the correct row when the original sessionToken is not available.
   * Not persisted — set on load from storage.
   */
  _storageKey?: string;
}

/**
 * Internal helper: pull the encrypted access token off a Table Storage row
 * and decrypt it. Returns empty string if the columns are missing (which
 * indicates a pre- plaintext row that needs to be migrated).
 */
function decryptAccessTokenFromEntity(entity: Record<string, unknown>): string {
  const ct = entity.accessTokenCiphertext as string | undefined;
  const iv = entity.accessTokenIv as string | undefined;
  const tag = entity.accessTokenAuthTag as string | undefined;
  if (!ct || !iv || !tag) return '';
  return decryptWithDek({ ciphertext: ct, iv, authTag: tag });
}

/**
 * Helper: convert a Table Storage entity to a StoredSession.
 * Handles both legacy (RowKey = userId) and new (RowKey = tokenHash[:32],
 * userId as column) formats transparently.
 */
function entityToSession(
  entity: Record<string, unknown>,
  sessionToken?: string,
): StoredSession {
  // New format stores userId as a column; legacy uses RowKey as userId.
  const userId = (entity.userId as string) || (entity.rowKey as string);
  return {
    userId,
    homeAccountId: entity.homeAccountId as string,
    displayName: entity.displayName as string,
    email: entity.email as string,
    tenantId: entity.tenantId as string,
    accessToken: decryptAccessTokenFromEntity(entity),
    expiresAt: entity.expiresAt as number,
    sessionToken: sessionToken || '',
    sessionCreatedAt: (entity.sessionCreatedAt as number) || 0,
    sessionAbsoluteCreatedAt: (entity.sessionAbsoluteCreatedAt as number) || undefined,
    deviceLabel: entity.deviceLabel as string | undefined,
    _storageKey: entity.rowKey as string,
  };
}

export async function saveSession(session: StoredSession): Promise<void> {
  await ensureTables();
  const accessEnvelope = encryptWithDek(session.accessToken);

  if (session.sessionToken) {
    // New session or re-auth — RowKey = tokenHash[:32] for multi-device support.
    // Each device/login gets its own row; no more overwriting.
    const tokenHash = hashSessionToken(session.sessionToken);
    const rowKey = tokenHash.slice(0, 32);
    await getSessionsTable().upsertEntity({
      partitionKey: 'session',
      rowKey,
      userId: session.userId,
      homeAccountId: session.homeAccountId,
      displayName: session.displayName,
      email: session.email,
      tenantId: session.tenantId,
      expiresAt: session.expiresAt,
      sessionCreatedAt: session.sessionCreatedAt || Date.now(),
      sessionAbsoluteCreatedAt: session.sessionAbsoluteCreatedAt ?? null,
      deviceLabel: session.deviceLabel ?? null,
      accessTokenCiphertext: accessEnvelope.ciphertext,
      accessTokenIv: accessEnvelope.iv,
      accessTokenAuthTag: accessEnvelope.authTag,
      sessionTokenHash: tokenHash,
    }, 'Replace');
  } else if (session._storageKey) {
    // Refresh path — update the specific row identified by _storageKey.
    // Merge preserves the existing sessionTokenHash.
    await getSessionsTable().upsertEntity({
      partitionKey: 'session',
      rowKey: session._storageKey,
      userId: session.userId,
      homeAccountId: session.homeAccountId,
      displayName: session.displayName,
      email: session.email,
      tenantId: session.tenantId,
      expiresAt: session.expiresAt,
      sessionCreatedAt: session.sessionCreatedAt || Date.now(),
      sessionAbsoluteCreatedAt: session.sessionAbsoluteCreatedAt ?? null,
      deviceLabel: session.deviceLabel ?? null,
      accessTokenCiphertext: accessEnvelope.ciphertext,
      accessTokenIv: accessEnvelope.iv,
      accessTokenAuthTag: accessEnvelope.authTag,
    }, 'Merge');
  } else {
    console.warn('[tableStorage] saveSession called without sessionToken or _storageKey — cannot persist');
  }
}

/**
 * Load the most recent session for a user. Checks both legacy (RowKey = userId)
 * and new (RowKey = tokenHash[:32]) row formats.
 */
export async function loadSession(userId: string): Promise<StoredSession | null> {
  await ensureTables();

  // Try legacy format first (direct lookup, O(1))
  try {
    const entity = await getSessionsTable().getEntity('session', userId);
    return entityToSession(entity as Record<string, unknown>);
  } catch { /* not found — try new format */ }

  // Scan for new-format rows with matching userId column, return most recent
  let best: StoredSession | null = null;
  const entities = getSessionsTable().listEntities({
    queryOptions: { filter: "PartitionKey eq 'session'" },
  });
  for await (const entity of entities) {
    if (entity.userId === userId) {
      const session = entityToSession(entity as Record<string, unknown>);
      if (!best || session.sessionCreatedAt > best.sessionCreatedAt) {
        best = session;
      }
    }
  }
  return best;
}

/**
 * Look up a session by Bearer token. First tries direct RowKey lookup (O(1))
 * using the tokenHash prefix, then falls back to a partition scan for legacy rows.
 */
export async function loadSessionByToken(token: string): Promise<StoredSession | null> {
  await ensureTables();
  const targetHash = hashSessionToken(token);
  const rowKey = targetHash.slice(0, 32);

  // Fast path: direct RowKey lookup (new format)
  try {
    const entity = await getSessionsTable().getEntity('session', rowKey);
    if (entity.sessionTokenHash === targetHash) {
      return entityToSession(entity as Record<string, unknown>, token);
    }
  } catch { /* not found — fall through to scan */ }

  // Slow path: scan for legacy rows (RowKey = userId)
  const entities = getSessionsTable().listEntities({
    queryOptions: { filter: "PartitionKey eq 'session'" },
  });
  for await (const entity of entities) {
    if (entity.sessionTokenHash === targetHash) {
      return entityToSession(entity as Record<string, unknown>, token);
    }
  }
  return null;
}

/**
 * Remove a single session by its storage key (RowKey).
 */
export async function removeSessionByKey(storageKey: string): Promise<void> {
  await ensureTables();
  try {
    await getSessionsTable().deleteEntity('session', storageKey);
  } catch { /* not found is fine */ }
}

/**
 * Remove all sessions for a user (both legacy and new format rows).
 */
export async function removeSession(userId: string): Promise<void> {
  await ensureTables();
  // Delete legacy row
  try {
    await getSessionsTable().deleteEntity('session', userId);
  } catch { /* not found is fine */ }

  // Scan and delete new-format rows
  const entities = getSessionsTable().listEntities({
    queryOptions: { filter: "PartitionKey eq 'session'" },
  });
  for await (const entity of entities) {
    if (entity.userId === userId) {
      try {
        await getSessionsTable().deleteEntity('session', entity.rowKey as string);
      } catch { /* ignore */ }
    }
  }
}

export async function listAllSessions(): Promise<StoredSession[]> {
  await ensureTables();
  const sessions: StoredSession[] = [];
  const entities = getSessionsTable().listEntities({
    queryOptions: { filter: "PartitionKey eq 'session'" },
  });
  for await (const entity of entities) {
    sessions.push(entityToSession(entity as Record<string, unknown>));
  }
  return sessions;
}

// ── Install nonce persistence ──
//
// Implements the install-time hands-free auth flow. install-mcp.sh generates
// a one-time nonce, embeds it in the OAuth login URL, and polls install-poll
// while the user completes browser OAuth. The callback writes the resulting
// session here keyed by sha256(nonce); the poll endpoint reads and atomically
// deletes (one-time-use, race-safe via ETag). Records have a 5-minute TTL.
//
// Key design point: a "not found" lookup means the user's browser hasn't
// finished OAuth yet (or never will). The poll endpoint treats not-found as
// 'pending' and lets the client's own timeout handle the never-completes case.
// We do NOT pre-reserve the nonce on /api/auth/login because that creates a
// race with the polling client (the script can poll before the browser has
// navigated to the login URL). Instead, the entry is upserted at OAuth
// callback time only.

export interface InstallNonceRecord {
  sessionToken: string;
  userId: string;
  email: string;
  displayName: string;
  deviceLabel?: string;
  /** Unix timestamp in milliseconds when this record expires. */
  expiresAt: number;
}

function hashNonce(nonce: string): string {
  return createHash('sha256').update(nonce).digest('hex');
}

/**
 * Upsert a completed session under the given nonce. Called from the OAuth
 * callback after the session has been created. Always returns true on
 * successful write — there is no pre-reservation to fail against.
 */
export async function attachSessionToInstallNonce(
  nonce: string,
  record: InstallNonceRecord
): Promise<boolean> {
  await ensureTables();
  const rowKey = hashNonce(nonce);
  try {
    await getInstallNoncesTable().upsertEntity(
      {
        partitionKey: 'nonce',
        rowKey,
        sessionToken: record.sessionToken,
        userId: record.userId,
        email: record.email,
        displayName: record.displayName,
        deviceLabel: record.deviceLabel ?? null,
        expiresAt: record.expiresAt,
      },
      'Replace'
    );
    return true;
  } catch {
    return false;
  }
}

/**
 * Atomically read-and-delete an attached install nonce.
 *
 * Returns:
 *   - InstallNonceRecord — found, attached, not expired, and we won the
 *     ETag delete race. Caller should treat this as the one-shot consumption.
 *   - 'pending'          — not found in storage. The user's browser may not
 *     have completed OAuth yet, or this poll lost the consumption race to
 *     another concurrent client. The caller's own polling timeout decides
 *     when to give up.
 *   - null               — found but expired (record cleaned up best-effort).
 *
 * The poll endpoint maps:
 *   InstallNonceRecord → 200
 *   'pending'          → 202
 *   null               → 410
 */
export async function consumeInstallNonce(
  nonce: string
): Promise<InstallNonceRecord | 'pending' | null> {
  await ensureTables();
  const rowKey = hashNonce(nonce);
  let entity;
  try {
    entity = await getInstallNoncesTable().getEntity('nonce', rowKey);
  } catch {
    // Not in storage — either the OAuth callback hasn't fired yet, or this
    // nonce was consumed by another concurrent poll. Either way, the safe
    // signal to the client is "still waiting".
    return 'pending';
  }

  const expiresAt = entity.expiresAt as number;
  if (expiresAt < Date.now()) {
    try {
      await getInstallNoncesTable().deleteEntity('nonce', rowKey);
    } catch { /* race ok */ }
    return null;
  }

  const record: InstallNonceRecord = {
    sessionToken: entity.sessionToken as string,
    userId: entity.userId as string,
    email: entity.email as string,
    displayName: entity.displayName as string,
    deviceLabel: entity.deviceLabel as string | undefined,
    expiresAt,
  };

  // Race-safe one-time-use: only the caller that wins the ETag check returns
  // the record. Losers fall through to 'pending' so they don't return a stale
  // session to a second client (which then can't actually authenticate, since
  // the session is bound to this one nonce flow).
  try {
    await getInstallNoncesTable().deleteEntity('nonce', rowKey, {
      etag: entity.etag,
    });
  } catch {
    return 'pending';
  }
  return record;
}

// ── MSAL cache persistence ──
//
// The MSAL token cache holds refresh tokens for every signed-in user. A
// stolen refresh token grants perpetual delegated Graph access until an
// admin revokes it in Entra. This is the highest-value credential we hold,
// so it's encrypted at rest with AES-256-GCM via the same DEK as access
// tokens. The serialized MSAL cache blob (JSON) goes through the envelope
// before being written to Table Storage.
//
// Schema (mcpMsalCache table) after hardening:
//   PartitionKey = 'cache'
//   RowKey       = 'msal-token-cache' (single fixed key — there's only one cache)
//   ciphertext   = AES-256-GCM ciphertext of the MSAL cache JSON (base64)
//   iv           = 12-byte IV (base64)
//   authTag      = 16-byte GCM auth tag (base64)
//
// Pre- rows had a `data` column with the plaintext JSON. The migration
// script (infra/scripts/purge-credentials.sh) wipes the table to force MSAL
// to re-issue from a clean state on first request.

const MSAL_CACHE_KEY = 'msal-token-cache';

export async function saveMsalCache(cacheData: string): Promise<void> {
  await ensureTables();
  const envelope = encryptWithDek(cacheData);
  await getMsalCacheTable().upsertEntity({
    partitionKey: 'cache',
    rowKey: MSAL_CACHE_KEY,
    ciphertext: envelope.ciphertext,
    iv: envelope.iv,
    authTag: envelope.authTag,
  }, 'Replace');
}

export async function loadMsalCache(): Promise<string | null> {
  await ensureTables();
  try {
    const entity = await getMsalCacheTable().getEntity('cache', MSAL_CACHE_KEY);
    const ct = entity.ciphertext as string | undefined;
    const iv = entity.iv as string | undefined;
    const tag = entity.authTag as string | undefined;
    if (!ct || !iv || !tag) {
      // Pre- plaintext row (or partial row from a botched migration).
      // Treat as empty so MSAL re-issues from a clean state.
      console.warn('[tableStorage] MSAL cache row exists but missing encrypted columns — treating as empty');
      return null;
    }
    return decryptWithDek({ ciphertext: ct, iv, authTag: tag });
  } catch (err) {
    console.error(
      '[tableStorage] Failed to load MSAL cache:',
      err instanceof Error ? err.message : err
    );
    return null;
  }
}
