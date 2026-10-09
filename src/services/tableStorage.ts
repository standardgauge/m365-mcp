import { TableClient } from '@azure/data-tables';
import { getTableClient, getTableServiceClient, isStorageConfigured } from './storageClient.js';
import { createHash } from 'crypto';
import {
  hashSessionToken,
  encryptWithDek,
  decryptWithDek,
  decryptWithDekMigrating,
  envelopeAad,
  boundColumnsAad,
  isLegacyIdentityBindingAllowed,
} from './credentialCrypto.js';
import { splitMsalCacheByAccount } from './msalCacheSplit.js';

const storageConfigured = isStorageConfigured();
const INSTALL_NONCES_TABLE = 'mcpInstallNonces';

let sessionsTable: TableClient | null = null;
let msalCacheTable: TableClient | null = null;
let installNoncesTable: TableClient | null = null;
let initPromise: Promise<void> | null = null;

function ensureTables(): Promise<void> {
  if (!storageConfigured) return Promise.resolve();
  if (!initPromise) {
    initPromise = initTables().catch((err) => {
      initPromise = null;
      throw err;
    });
  }
  return initPromise;
}

async function initTables(): Promise<void> {
  const serviceClient = getTableServiceClient();

  // Create tables if they don't exist
  try { await serviceClient.createTable('mcpSessions'); } catch { /* already exists */ }
  try { await serviceClient.createTable('mcpMsalCache'); } catch { /* already exists */ }
  try { await serviceClient.createTable(INSTALL_NONCES_TABLE); } catch { /* already exists */ }

  sessionsTable = getTableClient('mcpSessions');
  msalCacheTable = getTableClient('mcpMsalCache');
  installNoncesTable = getTableClient(INSTALL_NONCES_TABLE);

  // Before any session lookup is served: token lookups are a single point
  // read, so a row not keyed by its token hash would be unreachable.
  await migrateLegacySessionRows();
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
//   The envelope is bound via GCM AAD (sessionAccessTokenAad) to its row
//   and to the row's userId, homeAccountId and tenantId columns: copying it
//   into another row, or editing any of those three columns, makes it fail to
//   decrypt, and the session fails authentication. homeAccountId selects the
//   MSAL account a silent refresh uses, so an unprotected one would let a
//   storage writer repoint a session at another user's refresh token.
//   Older envelopes are bound to the row only, or carry no AAD at all; they
//   are still readable while MCP_SESSION_REQUIRE_IDENTITY_BINDING (and, for
//   unbound ones, MCP_ENVELOPE_REQUIRE_AAD) is unset, and are rewritten fully
//   bound on first read.
//
// Legacy rows (RowKey = userId, from before multi-session support) are
// re-keyed to the tokenHash format once per process start, before the first
// lookup (migrateLegacySessionRows). Rows no Bearer token can reach are
// dropped. An unknown token therefore costs one point read, never a scan.
//
// We do NOT store the original sessionToken in this table. After install it
// lives only on the client, in the OS credential store; during the install
// handoff it sits briefly, encrypted, in mcpInstallNonces (below). Loss of the HMAC key means every user
// must re-OAuth, but that's the desired property: the secret-at-rest material
// lives only in Key Vault (infra/), never in storage, so exfiltrating the
// tables does not yield it.

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

/** The session row columns that say whose session it is. */
interface SessionIdentity {
  userId: string;
  homeAccountId: string;
  tenantId: string;
}

/** Row-only binding, as written before the identity columns were bound. */
function sessionAccessTokenRowAad(rowKey: string): string {
  return envelopeAad('mcpSessions', 'session', rowKey, 'accessToken');
}

function sessionAccessTokenAad(rowKey: string, identity: SessionIdentity): string {
  return boundColumnsAad(sessionAccessTokenRowAad(rowKey), [
    ['userId', identity.userId ?? ''],
    ['homeAccountId', identity.homeAccountId ?? ''],
    ['tenantId', identity.tenantId ?? ''],
  ]);
}

/**
 * Internal helper: pull the encrypted access token off a Table Storage row
 * and decrypt it, authenticating the row's identity columns on the way.
 * Throws if the envelope or any of those columns was tampered with.
 *
 * Returns empty string if the envelope columns are missing (a row from before
 * envelopes existed). Such a row authenticates nothing, so it is refused once
 * MCP_SESSION_REQUIRE_IDENTITY_BINDING=true.
 */
function decryptAccessTokenFromEntity(
  entity: Record<string, unknown>,
  identity: SessionIdentity,
): string {
  const rowKey = entity.rowKey as string;
  const { plaintext, legacy } = decryptAccessTokenEnvelope(entity, rowKey, identity);
  if (legacy) void rebindLegacyAccessToken(rowKey, identity, plaintext, entity.etag as string | undefined);
  return plaintext;
}

/**
 * Decrypt a row's access-token envelope against the AAD of `rowKey` and
 * `identity`, with no side effects. `legacy` is true when the envelope is
 * row-bound or unbound and should be rewritten fully bound. Empty plaintext
 * when the row has no envelope (refused once identity binding is required);
 * throws when the envelope does not decrypt.
 */
function decryptAccessTokenEnvelope(
  entity: Record<string, unknown>,
  rowKey: string,
  identity: SessionIdentity,
): { plaintext: string; legacy: boolean } {
  const ct = entity.accessTokenCiphertext as string | undefined;
  const iv = entity.accessTokenIv as string | undefined;
  const tag = entity.accessTokenAuthTag as string | undefined;
  if (!ct || !iv || !tag) {
    if (!isLegacyIdentityBindingAllowed()) {
      throw new Error('Session row has no access token envelope to authenticate it');
    }
    return { plaintext: '', legacy: false };
  }
  const envelope = { ciphertext: ct, iv, authTag: tag };
  try {
    return { plaintext: decryptWithDek(envelope, sessionAccessTokenAad(rowKey, identity)), legacy: false };
  } catch (err) {
    if (!isLegacyIdentityBindingAllowed()) throw err;
    try {
      // Row-bound or unbound envelope; the latter only while
      // MCP_ENVELOPE_REQUIRE_AAD is unset.
      const { plaintext } = decryptWithDekMigrating(envelope, sessionAccessTokenRowAad(rowKey));
      return { plaintext, legacy: true };
    } catch {
      throw err;
    }
  }
}

/**
 * Re-encrypt an older access token envelope bound to the row and its identity
 * columns as currently stored. Merges only the three envelope columns,
 * conditional on the ETag we read, so a concurrent refresh write wins and this
 * becomes a no-op. Best-effort: a failure leaves the legacy envelope for the
 * next read to retry.
 */
async function rebindLegacyAccessToken(
  rowKey: string,
  identity: SessionIdentity,
  plaintext: string,
  etag: string | undefined,
): Promise<void> {
  try {
    const envelope = encryptWithDek(plaintext, sessionAccessTokenAad(rowKey, identity));
    await getSessionsTable().updateEntity(
      {
        partitionKey: 'session',
        rowKey,
        accessTokenCiphertext: envelope.ciphertext,
        accessTokenIv: envelope.iv,
        accessTokenAuthTag: envelope.authTag,
      },
      'Merge',
      etag ? { etag } : undefined,
    );
  } catch (err) {
    console.warn(
      '[tableStorage] Could not rebind legacy access token envelope:',
      err instanceof Error ? err.message : err,
    );
  }
}

/** The identity columns of a session row, as bound into its envelope AAD. */
function entityIdentity(entity: Record<string, unknown>): SessionIdentity {
  // Current rows store userId as a column; legacy rows use RowKey as userId.
  return {
    userId: (entity.userId as string) || (entity.rowKey as string),
    homeAccountId: entity.homeAccountId as string,
    tenantId: entity.tenantId as string,
  };
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
  const identity = entityIdentity(entity);
  return {
    ...identity,
    displayName: entity.displayName as string,
    email: entity.email as string,
    accessToken: decryptAccessTokenFromEntity(entity, identity),
    expiresAt: entity.expiresAt as number,
    sessionToken: sessionToken || '',
    sessionCreatedAt: (entity.sessionCreatedAt as number) || 0,
    sessionAbsoluteCreatedAt: (entity.sessionAbsoluteCreatedAt as number) || undefined,
    deviceLabel: entity.deviceLabel as string | undefined,
    _storageKey: entity.rowKey as string,
  };
}

/**
 * Raised by saveSession in 'update' mode when the session's row no longer
 * exists. The row was deleted (logout, credential purge, manual deletion) on
 * some replica after this one cached the session, so the session is revoked.
 */
export class SessionRowMissingError extends Error {
  constructor(storageKey: string) {
    super(`Session row ${storageKey.slice(0, 8)}… no longer exists`);
    this.name = 'SessionRowMissingError';
  }
}

function isNotFound(err: unknown): boolean {
  return (err as { statusCode?: number } | null)?.statusCode === 404;
}

/**
 * Persist a session.
 *
 * 'create' writes a new row keyed by tokenHash[:32] (one row per device or
 * login). It upserts, so it must only be used for a session that has just
 * been minted.
 *
 * 'update' writes back a session that was already persisted (TTL refresh,
 * sliding-window touch, access-token refresh). It is a conditional Merge on
 * the existing row (`If-Match: *`), so it can never recreate a row that was
 * deleted in the meantime. A missing row raises SessionRowMissingError.
 */
export async function saveSession(
  session: StoredSession,
  mode: 'create' | 'update' = 'create',
): Promise<void> {
  await ensureTables();
  maybePurgeExpiredInstallNonces();

  if (mode === 'update' && session._storageKey) {
    const accessEnvelope = encryptWithDek(
      session.accessToken,
      sessionAccessTokenAad(session._storageKey, session),
    );
    try {
      await getSessionsTable().updateEntity({
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
      }, 'Merge', { etag: '*' });
    } catch (err) {
      if (isNotFound(err)) throw new SessionRowMissingError(session._storageKey);
      throw err;
    }
  } else if (session.sessionToken) {
    // New session — RowKey = tokenHash[:32] for multi-device support.
    // Each device/login gets its own row; no more overwriting.
    const tokenHash = hashSessionToken(session.sessionToken);
    const rowKey = tokenHash.slice(0, 32);
    const accessEnvelope = encryptWithDek(
      session.accessToken,
      sessionAccessTokenAad(rowKey, session),
    );
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
  } else {
    console.warn('[tableStorage] saveSession called without sessionToken or _storageKey — cannot persist');
  }
}

/**
 * Whether a session's row still exists. A cheap point read used to revalidate
 * sessions held in a replica's memory cache. Throws on any error other than
 * not-found so the caller can tell "deleted" from "storage unreachable".
 * Returns true when storage is not configured: there is no row to have been
 * deleted, and the memory cache is the only store.
 */
export async function sessionRowExists(storageKey: string): Promise<boolean> {
  if (!storageConfigured) return true;
  await ensureTables();
  try {
    await getSessionsTable().getEntity('session', storageKey, { queryOptions: { select: ['RowKey'] } });
    return true;
  } catch (err) {
    if (isNotFound(err)) return false;
    throw err;
  }
}

/**
 * Load the most recent session for a user.
 */
export async function loadSession(userId: string): Promise<StoredSession | null> {
  await ensureTables();

  // Scan for rows with matching userId column, return most recent
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
 * Look up a session by Bearer token: one point read on RowKey =
 * tokenHash[:32]. A miss is a miss; there is no scan fallback, so an
 * unknown token costs the same as a known one.
 */
export async function loadSessionByToken(token: string): Promise<StoredSession | null> {
  await ensureTables();
  maybePurgeExpiredInstallNonces();
  const targetHash = hashSessionToken(token);
  const rowKey = targetHash.slice(0, 32);

  let entity: Record<string, unknown>;
  try {
    entity = await getSessionsTable().getEntity('session', rowKey);
  } catch {
    return null;
  }
  if (entity.sessionTokenHash !== targetHash) return null;
  return entityToSession(entity, token);
}

/**
 * Re-key every session row that is not stored under RowKey =
 * sessionTokenHash[:32]. Those are legacy rows (RowKey = userId) written
 * before multi-session support; the token lookup no longer scans for them.
 *
 * A row carrying a sessionTokenHash is copied to its tokenHash RowKey, with
 * the access-token envelope re-bound to the new row and its identity
 * columns, and the old row is deleted. A row without one can never
 * authenticate a Bearer token and is dropped. So is a row whose envelope no
 * longer decrypts, or that has no envelope once identity binding is
 * required: its user signs in again.
 *
 * Safe to run on several replicas at once: the copy is an insert, so a row
 * already at the new key (another replica, or a newer sign-in) wins, and the
 * delete is conditional on the ETag read. Best-effort per row: a failure is
 * logged and retried on the next process start.
 */
export async function migrateLegacySessionRows(): Promise<{ migrated: number; dropped: number }> {
  let migrated = 0;
  let dropped = 0;
  const table = getSessionsTable();
  const entities = table.listEntities({ queryOptions: { filter: "PartitionKey eq 'session'" } });
  try {
    for await (const raw of entities) {
      const entity = raw as Record<string, unknown>;
      const oldKey = entity.rowKey as string;
      const hash = entity.sessionTokenHash as string | undefined;
      if (hash && oldKey === hash.slice(0, 32)) continue;
      const etag = entity.etag as string | undefined;

      try {
        const newKey = hash ? hash.slice(0, 32) : null;
        let copied = false;
        if (newKey) {
          const identity = entityIdentity(entity);
          let accessToken: string | null = null;
          try {
            accessToken = decryptAccessTokenEnvelope(entity, oldKey, identity).plaintext;
          } catch {
            accessToken = null;
          }
          if (accessToken !== null) {
            const envelope = accessToken
              ? encryptWithDek(accessToken, sessionAccessTokenAad(newKey, identity))
              : null;
            try {
              await table.createEntity({
                partitionKey: 'session',
                rowKey: newKey,
                userId: identity.userId,
                homeAccountId: entity.homeAccountId ?? null,
                displayName: entity.displayName ?? null,
                email: entity.email ?? null,
                tenantId: entity.tenantId ?? null,
                expiresAt: entity.expiresAt ?? null,
                sessionCreatedAt: entity.sessionCreatedAt ?? null,
                sessionAbsoluteCreatedAt: entity.sessionAbsoluteCreatedAt ?? null,
                deviceLabel: entity.deviceLabel ?? null,
                accessTokenCiphertext: envelope?.ciphertext ?? null,
                accessTokenIv: envelope?.iv ?? null,
                accessTokenAuthTag: envelope?.authTag ?? null,
                sessionTokenHash: hash,
              });
            } catch (err) {
              // 409: a row already holds this key. It is at least as current.
              if ((err as { statusCode?: number }).statusCode !== 409) throw err;
            }
            copied = true;
          }
        }
        await table.deleteEntity('session', oldKey, etag ? { etag } : undefined);
        if (copied) migrated++;
        else dropped++;
      } catch (err) {
        console.warn(
          '[tableStorage] Could not migrate legacy session row; retried on next start:',
          err instanceof Error ? err.message : err,
        );
      }
    }
  } catch (err) {
    console.error('[tableStorage] Legacy session migration stopped early:', err);
  }
  if (migrated || dropped) {
    console.log(`[tableStorage] Legacy session rows: ${migrated} re-keyed, ${dropped} dropped`);
  }
  return { migrated, dropped };
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
 * Remove all sessions for a user.
 */
export async function removeSession(userId: string): Promise<void> {
  await ensureTables();
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
// while the user completes browser OAuth. Once the browser has confirmed the
// handoff (install-confirm, below), the resulting session is written here keyed
// by sha256(nonce); the poll endpoint reads and atomically
// deletes (one-time-use, race-safe via ETag). Records have a 5-minute TTL.
//
// The session token is the one value in this table that authenticates, and it
// is the only place the server ever holds it in a recoverable form. It is
// stored as an AES-256-GCM envelope bound to its row (sessionToken* columns,
// AAD names this table, partition, row and column), the same as access tokens
// in mcpSessions, so a copy of the table yields nothing usable without the
// data encryption key. Rows written before encryption carried a plaintext
// `sessionToken` column; those are never handed out, only deleted.
//
// A row nobody polls (the installer was closed mid-flow) would otherwise sit
// in the table forever. Expired rows are deleted by purgeExpiredInstallNonces,
// which session saves, token lookups, attach and consume start at most once per
// NONCE_PURGE_INTERVAL_MS per process. Like the audit purge, this is not a
// timer trigger: the Container App runs without AzureWebJobsStorage (see
// auditLog.ts), so rows are purged while the server is in use, which is also
// the only time new ones are written.
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

const INSTALL_NONCE_PARTITION = 'nonce';
/** Pending install handoffs (install-confirm) share the table under their own partition. */
const HANDOFF_PARTITION = 'handoff';
const NONCE_PURGE_INTERVAL_MS = 10 * 60 * 1000;
// Bounds one run. A backlog larger than this is finished by later runs.
const NONCE_PURGE_MAX_ROWS_PER_RUN = 5_000;

let lastNoncePurgeStartedAt = 0;
let noncePurgeInFlight: Promise<number> | null = null;

function hashNonce(nonce: string): string {
  return createHash('sha256').update(nonce).digest('hex');
}

function installNonceTokenAad(rowKey: string): string {
  return envelopeAad(INSTALL_NONCES_TABLE, INSTALL_NONCE_PARTITION, rowKey, 'sessionToken');
}

/**
 * Upsert a completed session under the given nonce. Called from install-confirm
 * once the signed-in browser has entered the installer's confirmation code.
 * Returns false if the write (or encrypting the token) fails — there is no
 * pre-reservation to fail against. The plaintext token never reaches storage.
 */
export async function attachSessionToInstallNonce(
  nonce: string,
  record: InstallNonceRecord
): Promise<boolean> {
  await ensureTables();
  maybePurgeExpiredInstallNonces();
  const rowKey = hashNonce(nonce);
  try {
    const envelope = encryptWithDek(record.sessionToken, installNonceTokenAad(rowKey));
    await getInstallNoncesTable().upsertEntity(
      {
        partitionKey: INSTALL_NONCE_PARTITION,
        rowKey,
        sessionTokenCiphertext: envelope.ciphertext,
        sessionTokenIv: envelope.iv,
        sessionTokenAuthTag: envelope.authTag,
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
 *   - null               — found but expired, or its token envelope is
 *     missing or does not decrypt for this row (a pre-encryption plaintext
 *     row, or one tampered with). The row is deleted best-effort.
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
  maybePurgeExpiredInstallNonces();
  const rowKey = hashNonce(nonce);
  let entity;
  try {
    entity = await getInstallNoncesTable().getEntity(INSTALL_NONCE_PARTITION, rowKey);
  } catch {
    // Not in storage — either the OAuth callback hasn't fired yet, or this
    // nonce was consumed by another concurrent poll. Either way, the safe
    // signal to the client is "still waiting".
    return 'pending';
  }

  const discard = async (): Promise<null> => {
    try {
      await getInstallNoncesTable().deleteEntity(INSTALL_NONCE_PARTITION, rowKey);
    } catch { /* race ok */ }
    return null;
  };

  const expiresAt = entity.expiresAt as number;
  if (!(expiresAt >= Date.now())) return discard();

  let sessionToken: string;
  try {
    sessionToken = decryptWithDek(
      {
        ciphertext: entity.sessionTokenCiphertext as string,
        iv: entity.sessionTokenIv as string,
        authTag: entity.sessionTokenAuthTag as string,
      },
      installNonceTokenAad(rowKey),
    );
  } catch {
    // No envelope (a row written before encryption) or one that does not
    // belong to this row. The installer is told to start again.
    return discard();
  }

  const record: InstallNonceRecord = {
    sessionToken,
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
    await getInstallNoncesTable().deleteEntity(INSTALL_NONCE_PARTITION, rowKey, {
      etag: entity.etag,
    });
  } catch {
    return 'pending';
  }
  return record;
}

/**
 * Delete install-nonce and pending-handoff rows whose expiresAt has passed,
 * and rows with no expiresAt at all (nothing could ever consume them). Returns the number of
 * rows deleted. Each delete carries the row's ETag, so a row re-attached
 * between the scan and the delete is left alone.
 */
export async function purgeExpiredInstallNonces(nowMs: number = Date.now()): Promise<number> {
  await ensureTables();
  const table = getInstallNoncesTable();
  const iterator = table.listEntities<{ expiresAt?: number }>({
    queryOptions: {
      filter: `PartitionKey eq '${INSTALL_NONCE_PARTITION}' or PartitionKey eq '${HANDOFF_PARTITION}'`,
      select: ['PartitionKey', 'RowKey', 'expiresAt'],
    },
  });

  let deleted = 0;
  let seen = 0;
  for await (const entity of iterator) {
    if (++seen > NONCE_PURGE_MAX_ROWS_PER_RUN) break;
    const expiresAt = entity.expiresAt;
    if (typeof expiresAt === 'number' && expiresAt >= nowMs) continue;
    try {
      await table.deleteEntity(entity.partitionKey as string, entity.rowKey as string, { etag: entity.etag });
      deleted++;
    } catch {
      /* consumed, re-attached, or deleted by another replica meanwhile */
    }
  }
  if (deleted > 0) {
    console.log(`[installNonce] Purged ${deleted} expired install-nonce row(s)`);
  }
  return deleted;
}

/**
 * Start a purge of expired install-nonce rows if none is running in this
 * process and the last one started more than NONCE_PURGE_INTERVAL_MS ago.
 * Fire-and-forget; never throws.
 */
export function maybePurgeExpiredInstallNonces(nowMs: number = Date.now()): void {
  if (!storageConfigured || noncePurgeInFlight) return;
  if (nowMs - lastNoncePurgeStartedAt < NONCE_PURGE_INTERVAL_MS) return;
  lastNoncePurgeStartedAt = nowMs;
  noncePurgeInFlight = purgeExpiredInstallNonces(nowMs)
    .catch((err: unknown) => {
      console.error('[installNonce] Expired-row purge failed:', err instanceof Error ? err.message : err);
      return 0;
    })
    .finally(() => {
      noncePurgeInFlight = null;
    });
}

/** Test hook: wait for an in-flight purge, then forget it so the next call runs. */
export async function resetInstallNoncePurgeStateForTests(): Promise<void> {
  if (noncePurgeInFlight) await noncePurgeInFlight;
  lastNoncePurgeStartedAt = 0;
  noncePurgeInFlight = null;
}

// ── Install handoff confirmation ──
//
// The callback no longer attaches a session to an install nonce directly. It
// records a pending handoff here, keyed by a random id held in an HttpOnly
// cookie on the browser that signed in, and the install-confirm page attaches
// the session only after that browser enters the installer's confirmation code
// (see installConfirm.ts). Rows live in the install-nonce table under their own
// partition and carry no session token: the token is read from the confirming
// request's own mcp_session cookie and checked against the keyed hash stored
// here, so the handoff can only attach the session this sign-in created.

export const INSTALL_HANDOFF_MAX_ATTEMPTS = 5;

export interface InstallHandoffRecord {
  /** SHA256(verifier) the installer put in the login URL. */
  challenge: string;
  /** Keyed HMAC of the session token the callback issued to this browser. */
  sessionTokenHash: string;
  userId: string;
  email: string;
  displayName: string;
  deviceLabel?: string;
  /** Wrong codes entered so far. */
  attempts: number;
  /** Unix timestamp in milliseconds when this record expires. */
  expiresAt: number;
}

export interface StoredInstallHandoff {
  record: InstallHandoffRecord;
  etag: string;
}


export async function createInstallHandoff(
  handoffId: string,
  record: InstallHandoffRecord
): Promise<void> {
  await ensureTables();
  await getInstallNoncesTable().createEntity({
    partitionKey: HANDOFF_PARTITION,
    rowKey: hashNonce(handoffId),
    challenge: record.challenge,
    sessionTokenHash: record.sessionTokenHash,
    userId: record.userId,
    email: record.email,
    displayName: record.displayName,
    deviceLabel: record.deviceLabel ?? null,
    attempts: record.attempts,
    expiresAt: record.expiresAt,
  });
}

/** The pending handoff, or null if it is unknown, used, or expired. */
export async function getInstallHandoff(handoffId: string): Promise<StoredInstallHandoff | null> {
  await ensureTables();
  const rowKey = hashNonce(handoffId);
  let entity;
  try {
    entity = await getInstallNoncesTable().getEntity(HANDOFF_PARTITION, rowKey);
  } catch {
    return null;
  }
  const expiresAt = entity.expiresAt as number;
  if (expiresAt < Date.now()) {
    try {
      await getInstallNoncesTable().deleteEntity(HANDOFF_PARTITION, rowKey);
    } catch { /* race ok */ }
    return null;
  }
  return {
    etag: entity.etag as string,
    record: {
      challenge: entity.challenge as string,
      sessionTokenHash: entity.sessionTokenHash as string,
      userId: entity.userId as string,
      email: entity.email as string,
      displayName: entity.displayName as string,
      deviceLabel: (entity.deviceLabel as string | null) ?? undefined,
      attempts: (entity.attempts as number) ?? 0,
      expiresAt,
    },
  };
}

/**
 * Count a wrong code against the handoff, deleting it once the limit is
 * reached. Returns the attempts remaining, 0 when the handoff is gone, or null
 * if another request changed the row first (the caller asks the user to retry).
 */
export async function recordInstallHandoffFailure(
  handoffId: string,
  stored: StoredInstallHandoff
): Promise<number | null> {
  const rowKey = hashNonce(handoffId);
  const attempts = stored.record.attempts + 1;
  try {
    if (attempts >= INSTALL_HANDOFF_MAX_ATTEMPTS) {
      await getInstallNoncesTable().deleteEntity(HANDOFF_PARTITION, rowKey, { etag: stored.etag });
      return 0;
    }
    await getInstallNoncesTable().updateEntity(
      { partitionKey: HANDOFF_PARTITION, rowKey, attempts },
      'Merge',
      { etag: stored.etag }
    );
    return INSTALL_HANDOFF_MAX_ATTEMPTS - attempts;
  } catch {
    return null;
  }
}

/**
 * Remove the handoff. With an etag, only succeeds for the request that read
 * that version, which makes a successful confirmation one-time.
 */
export async function deleteInstallHandoff(handoffId: string, etag?: string): Promise<boolean> {
  await ensureTables();
  try {
    await getInstallNoncesTable().deleteEntity(
      HANDOFF_PARTITION,
      hashNonce(handoffId),
      etag ? { etag } : undefined
    );
    return true;
  } catch {
    return false;
  }
}

// ── MSAL cache persistence ──
//
// The MSAL token cache holds refresh tokens. A stolen refresh token grants
// delegated Graph access until an admin revokes it in Entra, so this is the
// highest-value credential we hold. It is encrypted at rest with AES-256-GCM
// under the same DEK as access tokens.
//
// Each account's cache is its own row (threat model section 2, row 2.4):
//   PartitionKey = 'account'
//   RowKey       = MSAL home account id (`<object id>.<tenant id>`; plaintext,
//                  as it already is on mcpSessions rows)
//   ciphertext   = AES-256-GCM ciphertext of that account's MSAL cache JSON
//   iv           = 12-byte IV (base64)
//   authTag      = 16-byte GCM auth tag (base64)
//   The envelope is bound by GCM AAD to its table, partition and row, so one
//   account's cache copied into another account's row does not decrypt.
//
// Every write is conditional on the ETag read before it (or on the row not
// existing yet), so a replica never overwrites a cache another replica wrote
// after it read. A losing write raises MsalCacheConflictError and the caller
// keeps the winner's cache.
//
// Before this, every account shared one row ('cache' / 'msal-token-cache'),
// written with an unconditional replace. splitLegacyMsalCache moves that row's
// accounts into their own rows once per process and deletes it.

const MSAL_PARTITION = 'account';
const LEGACY_MSAL_PARTITION = 'cache';
const LEGACY_MSAL_CACHE_KEY = 'msal-token-cache';
const LEGACY_MSAL_CACHE_AAD = envelopeAad(
  'mcpMsalCache', LEGACY_MSAL_PARTITION, LEGACY_MSAL_CACHE_KEY, 'msalCache',
);

// Entra home account ids are two GUIDs joined by a dot. Anything outside this
// set is refused rather than used as a RowKey.
const HOME_ACCOUNT_ID_PATTERN = /^[A-Za-z0-9._-]{1,256}$/;

function msalRowKey(homeAccountId: string): string {
  if (!HOME_ACCOUNT_ID_PATTERN.test(homeAccountId)) {
    throw new Error('Refusing to use a malformed home account id as an MSAL cache row key');
  }
  return homeAccountId;
}

function msalCacheAad(rowKey: string): string {
  return envelopeAad('mcpMsalCache', MSAL_PARTITION, rowKey, 'msalCache');
}

function statusOf(err: unknown): number | undefined {
  return (err as { statusCode?: number } | null)?.statusCode;
}

/** Another writer changed (or created) the row since it was read. */
export class MsalCacheConflictError extends Error {
  constructor(message = 'MSAL cache row changed since it was read') {
    super(message);
    this.name = 'MsalCacheConflictError';
  }
}

export interface MsalCachePartition {
  /** The account's serialized MSAL cache, or null when there is none to use. */
  data: string | null;
  /** ETag of the row as read; undefined when the row does not exist. */
  etag: string | undefined;
}

/**
 * Reads one account's MSAL cache. A missing row is `{ data: null, etag:
 * undefined }`. A row that will not decrypt is logged and returned as
 * `data: null` with its ETag, so the next sign-in can replace it. Storage
 * errors other than not-found are thrown.
 */
export async function loadMsalCachePartition(homeAccountId: string): Promise<MsalCachePartition> {
  await ensureTables();
  await legacySplitOnce();
  const rowKey = msalRowKey(homeAccountId);
  let entity;
  try {
    entity = await getMsalCacheTable().getEntity(MSAL_PARTITION, rowKey);
  } catch (err) {
    if (statusOf(err) === 404) return { data: null, etag: undefined };
    throw err;
  }
  const ct = entity.ciphertext as string | undefined;
  const iv = entity.iv as string | undefined;
  const tag = entity.authTag as string | undefined;
  if (!ct || !iv || !tag) {
    console.warn('[tableStorage] MSAL cache row is missing encrypted columns — treating as empty');
    return { data: null, etag: entity.etag };
  }
  try {
    // These rows were only ever written AAD-bound, so there is no unbound
    // form to fall back to.
    return { data: decryptWithDek({ ciphertext: ct, iv, authTag: tag }, msalCacheAad(rowKey)), etag: entity.etag };
  } catch (err) {
    console.error(
      '[tableStorage] MSAL cache row did not decrypt — treating as empty:',
      err instanceof Error ? err.message : err,
    );
    return { data: null, etag: entity.etag };
  }
}

/**
 * Writes one account's MSAL cache, conditional on `etag`: the row must still
 * carry that ETag, or, when `etag` is undefined, must not exist yet. Returns
 * the new ETag. Throws MsalCacheConflictError when another writer got there
 * first.
 */
export async function saveMsalCachePartition(
  homeAccountId: string,
  cacheData: string,
  etag: string | undefined,
): Promise<string | undefined> {
  await ensureTables();
  const rowKey = msalRowKey(homeAccountId);
  const envelope = encryptWithDek(cacheData, msalCacheAad(rowKey));
  const entity = {
    partitionKey: MSAL_PARTITION,
    rowKey,
    ciphertext: envelope.ciphertext,
    iv: envelope.iv,
    authTag: envelope.authTag,
  };
  try {
    const res = etag
      ? await getMsalCacheTable().updateEntity(entity, 'Replace', { etag })
      : await getMsalCacheTable().createEntity(entity);
    return res?.etag;
  } catch (err) {
    const status = statusOf(err);
    // 412: the row changed since we read it. 409: it was created since we
    // found it missing. 404: it was deleted since we read it.
    if (status === 412 || status === 409 || status === 404) {
      throw new MsalCacheConflictError();
    }
    throw err;
  }
}

let legacySplit: Promise<void> | null = null;

function legacySplitOnce(): Promise<void> {
  if (legacySplit) return legacySplit;
  const run: Promise<void> = splitLegacyMsalCache().then(
    () => undefined,
    (err) => {
      console.error(
        '[tableStorage] Could not split the shared MSAL cache row:',
        err instanceof Error ? err.message : err,
      );
      // Try again on the next access rather than never.
      if (legacySplit === run) legacySplit = null;
    },
  );
  legacySplit = run;
  return run;
}

/**
 * Moves every account in the old shared cache row into its own row, then
 * deletes the shared row. Idempotent, and safe to run on several replicas at
 * once:
 *   - an account that already has its own row keeps it (create-only write),
 *     because that row was written by this code and is newer;
 *   - the shared row is deleted only if it still carries the ETag read here.
 *     If a replica still running the previous release wrote to it meanwhile,
 *     it stays, and the next process to start splits it again.
 * A shared row that does not decrypt is left in place and logged; it holds
 * nothing this code can use, and deleting credentials it cannot read is the
 * operator's call (infra/scripts/purge-credentials.sh).
 * Returns the number of accounts given a row of their own.
 */
export async function splitLegacyMsalCache(): Promise<number> {
  await ensureTables();
  const table = getMsalCacheTable();
  let entity;
  try {
    entity = await table.getEntity(LEGACY_MSAL_PARTITION, LEGACY_MSAL_CACHE_KEY);
  } catch (err) {
    if (statusOf(err) === 404) return 0;
    throw err;
  }

  const ct = entity.ciphertext as string | undefined;
  const iv = entity.iv as string | undefined;
  const tag = entity.authTag as string | undefined;
  let plaintext: string | null = null;
  if (ct && iv && tag) {
    try {
      plaintext = decryptWithDekMigrating({ ciphertext: ct, iv, authTag: tag }, LEGACY_MSAL_CACHE_AAD).plaintext;
    } catch (err) {
      console.error(
        '[tableStorage] Shared MSAL cache row did not decrypt; leaving it in place:',
        err instanceof Error ? err.message : err,
      );
      return 0;
    }
  }

  let created = 0;
  if (plaintext) {
    for (const [homeAccountId, data] of splitMsalCacheByAccount(plaintext)) {
      if (!HOME_ACCOUNT_ID_PATTERN.test(homeAccountId)) {
        console.warn('[tableStorage] Skipping an MSAL account with a malformed home account id');
        continue;
      }
      try {
        await saveMsalCachePartition(homeAccountId, data, undefined);
        created++;
      } catch (err) {
        if (!(err instanceof MsalCacheConflictError)) throw err;
      }
    }
  }

  try {
    await table.deleteEntity(LEGACY_MSAL_PARTITION, LEGACY_MSAL_CACHE_KEY, { etag: entity.etag });
  } catch (err) {
    if (statusOf(err) !== 404 && statusOf(err) !== 412) throw err;
  }
  console.log(`[tableStorage] Split the shared MSAL cache row into ${created} account row(s)`);
  return created;
}

/** Test hook: forget that the legacy split already ran in this process. */
export function resetLegacyMsalSplitForTests(): void {
  legacySplit = null;
}
