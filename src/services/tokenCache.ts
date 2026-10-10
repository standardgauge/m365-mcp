import { acquireTokenSilent } from './graphClient.js';
import {
  saveSession,
  loadSession,
  loadSessionByToken,
  removeSession,
  removeSessionByKey,
  listAllSessions,
  sessionRowExists,
  SessionRowMissingError,
} from './tableStorage.js';
import { hashSessionToken } from './credentialCrypto.js';
import { SessionStoreUnavailableError } from './sessionStoreError.js';

/** Session inactivity TTL (in milliseconds). Active sessions auto-extend
 *  via a sliding window in the auth middleware. When the TTL fires after
 *  inactivity, MSAL silent refresh extends the session transparently.
 *  7 days is the industry standard for headless API integrations (e.g.
 *  Auth0 default idle timeout). MSAL refresh tokens are valid for 90 days,
 *  well beyond this window. */
export const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000;

/** Maximum absolute session lifetime (30 days). Sessions past this bound
 *  must re-authenticate via full OAuth regardless of MSAL refresh token
 *  validity. Enforced on every authenticated request path so a stolen token
 *  cannot be refreshed indefinitely while the MSAL refresh token remains
 *  valid (up to 90 days). */
export const SESSION_MAX_LIFETIME_MS = 30 * 24 * 60 * 60 * 1000;

/** How long a replica trusts a session in its memory cache before checking
 *  that the session's row still exists. Deleting rows (logout, the credential
 *  purge script, manual deletion) reaches other replicas only through this
 *  check, so it is the bound on how long a deleted session keeps working. */
export const SESSION_REVALIDATE_MS = 30 * 1000;

/** While storage is unreachable a cached session keeps working, so a storage
 *  blip does not log every client out. This caps that grace: past it, the
 *  session is refused until storage answers. */
export const SESSION_REVALIDATE_MAX_STALE_MS = 5 * 60 * 1000;

/** Thrown by storeSession when the session's row was deleted after this
 *  replica cached it. The session is revoked and has been evicted. */
export { SessionRowMissingError };
/** Thrown by getSessionByToken when storage cannot answer. */
export { SessionStoreUnavailableError };

export interface UserSession {
  userId: string;
  homeAccountId: string;
  displayName: string;
  email: string;
  tenantId: string;
  accessToken: string;
  /** Unix timestamp in milliseconds — Graph access token expiry */
  expiresAt: number;
  /**
   * Crypto-random session token — the Bearer credential clients send. Only
   * present in the in-memory representation when the original token is
   * available (i.e. just after OAuth callback or just after a successful
   * Bearer-auth lookup). Sessions reloaded from Table Storage by userId
   * will have this empty, since storage holds only the HMAC hash.
   */
  sessionToken: string;
  /**
   * Unix timestamp (ms) when the session was created. Used to enforce
   * server-side session expiry independent of the Graph access token TTL.
   * Reset by the sliding-window refresh to extend the inactivity timeout.
   */
  sessionCreatedAt: number;
  /**
   * Immutable Unix timestamp (ms) when this session was first created via
   * OAuth. Never reset by any refresh operation. Used by /refresh to enforce
   * an absolute session lifetime so a stolen token cannot be renewed
   * indefinitely while the MSAL refresh token remains valid.
   * Zero (or absent) for legacy sessions created before this field was added.
   */
  sessionAbsoluteCreatedAt?: number;
  /**
   * Human-readable label for the device/client that created this session.
   * Set at install time via the device_label query parameter.
   */
  deviceLabel?: string;
  /**
   * Table Storage RowKey for this session. Used by the refresh path to
   * update the correct row without needing the original sessionToken.
   * Set when loaded from storage; not persisted as a column.
   */
  _storageKey?: string;
}

/**
 * Returns true when a session has exceeded its absolute lifetime bound
 * (SESSION_MAX_LIFETIME_MS). Uses sessionAbsoluteCreatedAt when present and
 * non-zero, falls back to sessionCreatedAt for legacy sessions. Returns false
 * when both fields are zero/absent — no anchor means no absolute check (these
 * sessions re-OAuth when the inactivity TTL fires).
 */
export function isAbsoluteLifetimeExceeded(session: UserSession, now = Date.now()): boolean {
  const anchor = session.sessionAbsoluteCreatedAt || session.sessionCreatedAt;
  if (!anchor) return false;
  return (now - anchor) > SESSION_MAX_LIFETIME_MS;
}

/**
 * Materialize an immutable absolute-lifetime anchor for legacy sessions that
 * were created before sessionAbsoluteCreatedAt was added. Sets
 * sessionAbsoluteCreatedAt = sessionCreatedAt in-place so the anchor is fixed
 * before any sliding-window touch or TTL refresh can rewrite sessionCreatedAt.
 *
 * Must be called immediately after loading a session from storage or cache,
 * before returning it to any caller that might mutate it. Returns true when
 * materialization occurred (caller should fire-and-forget a storeSession to
 * persist the anchor).
 */
function materializeLegacyAnchor(session: UserSession): boolean {
  if (!session.sessionAbsoluteCreatedAt && session.sessionCreatedAt) {
    session.sessionAbsoluteCreatedAt = session.sessionCreatedAt;
    return true;
  }
  return false;
}

// ── In-memory cache ──
// Session-scoped: each session is cached independently so multiple
// sessions for the same user (different devices) don't collide.
//
// Primary index:   tokenHash → UserSession  (used by getSessionByToken)
// Secondary index: userId    → tokenHash    (used by getSession for token refresh)
//
// The userId index stores the MOST RECENTLY AUTHENTICATED session's hash.
// This is correct for getValidAccessToken because each request goes through
// authenticateRequest first, which updates the userId index to point to the
// session that's actually being used for this request.
const sessionCache = new Map<string, UserSession>();   // tokenHash → session
const userIndex = new Map<string, string>();             // userId → tokenHash of most recent auth
const verifiedAt = new Map<string, number>();            // cache key → when storage last confirmed the row

function cachePut(key: string, session: UserSession): void {
  sessionCache.set(key, session);
  verifiedAt.set(key, Date.now());
}

function cacheEvict(key: string): void {
  const session = sessionCache.get(key);
  sessionCache.delete(key);
  verifiedAt.delete(key);
  if (session && userIndex.get(session.userId) === key) userIndex.delete(session.userId);
}

/**
 * Confirm a cached session's row still exists, at most once per
 * SESSION_REVALIDATE_MS. Returns 'gone' (and evicts) when the row is gone, and
 * 'unavailable' when storage has been unreachable for longer than
 * SESSION_REVALIDATE_MAX_STALE_MS since the last confirmation. Either way the
 * cached session must not be served.
 */
async function revalidateCached(
  key: string,
  session: UserSession,
): Promise<'ok' | 'gone' | { unavailable: unknown }> {
  const last = verifiedAt.get(key) ?? 0;
  const now = Date.now();
  if (now - last < SESSION_REVALIDATE_MS || !session._storageKey) return 'ok';
  try {
    if (await sessionRowExists(session._storageKey)) {
      verifiedAt.set(key, now);
      return 'ok';
    }
    console.warn(`[tokenCache] Session row for user ${session.userId} was deleted — evicting cached session`);
  } catch (err) {
    if (now - last < SESSION_REVALIDATE_MAX_STALE_MS) {
      console.warn('[tokenCache] Could not revalidate cached session; serving cached copy:', err);
      return 'ok';
    }
    console.error('[tokenCache] Could not revalidate cached session past the stale limit; refusing it:', err);
    // Refuse but keep the entry: it is still valid if storage comes back.
    return { unavailable: err };
  }
  cacheEvict(key);
  return 'gone';
}

/**
 * Compute or retrieve the cache key for a session.
 * For sessions with a token, this is the HMAC hash.
 * For sessions loaded by userId (no token), we use the _storageKey.
 */
function cacheKey(session: UserSession): string {
  if (session.sessionToken) return hashSessionToken(session.sessionToken);
  if (session._storageKey) return `key:${session._storageKey}`;
  return `user:${session.userId}`;
}

/**
 * Cache and persist a session.
 *
 * A session without a _storageKey has just been minted (OAuth callback,
 * device login) and gets a new row. Anything carrying a _storageKey came from
 * storage or from an earlier storeSession, so the write is an update that
 * fails rather than recreating the row if it was deleted meanwhile. That case
 * evicts the session and throws SessionRowMissingError: a refresh must not undo
 * a logout or a purge.
 */
export async function storeSession(session: UserSession): Promise<void> {
  const mode = session._storageKey ? 'update' : 'create';
  // Ensure _storageKey is set so refresh/delete paths know which row to target
  if (session.sessionToken && !session._storageKey) {
    session._storageKey = hashSessionToken(session.sessionToken).slice(0, 32);
  }
  const key = cacheKey(session);
  cachePut(key, session);
  userIndex.set(session.userId, key);
  try {
    await saveSession(session, mode);
  } catch (err) {
    if (err instanceof SessionRowMissingError) {
      console.warn(`[tokenCache] Session row for user ${session.userId} was deleted — not writing it back`);
      cacheEvict(key);
      throw err;
    }
    console.error('[tokenCache] Failed to persist session to Table Storage:', err);
  }
}

export async function getSession(userId: string): Promise<UserSession | undefined> {
  // Check memory — use the userId index to find the right session
  const cachedKey = userIndex.get(userId);
  if (cachedKey) {
    const cached = sessionCache.get(cachedKey);
    if (cached && (await revalidateCached(cachedKey, cached)) !== 'ok') return undefined;
    if (cached) {
      if (materializeLegacyAnchor(cached)) {
        storeSession(cached).catch((err) =>
          console.error('[tokenCache] Failed to persist legacy absolute anchor (getSession):', err)
        );
      }
      return cached;
    }
    // Stale index — drop and fall through
    userIndex.delete(userId);
  }

  // Fall back to Table Storage (e.g. after container restart)
  try {
    const stored = await loadSession(userId);
    if (stored) {
      if (materializeLegacyAnchor(stored)) {
        storeSession(stored).catch((err) =>
          console.error('[tokenCache] Failed to persist legacy absolute anchor (getSession storage):', err)
        );
      }
      const key = cacheKey(stored);
      cachePut(key, stored);
      userIndex.set(userId, key);
      return stored;
    }
  } catch (err) {
    console.error('[tokenCache] Failed to load session from Table Storage:', err);
  }
  return undefined;
}

/**
 * Look up a session by its Bearer token. Session-scoped: returns exactly
 * the session matching this token, even if the same user has other sessions.
 *
 * Two-tier lookup:
 *   1. sessionCache (in-memory, keyed by tokenHash) — fast path
 *   2. loadSessionByToken (storage lookup) — cold path on container restart
 *
 * Returns undefined only when storage says no session has this token. When
 * storage cannot answer (throttling, 5xx, auth, network), or a cached session
 * is past the revalidation stale limit, it throws SessionStoreUnavailableError
 * instead, so an outage surfaces as a server error rather than a 401.
 */
export async function getSessionByToken(token: string): Promise<UserSession | undefined> {
  const hash = hashSessionToken(token);

  // Fast path: direct session lookup by token hash
  const cached = sessionCache.get(hash);
  if (cached) {
    const state = await revalidateCached(hash, cached);
    if (state === 'gone') return undefined;
    if (state !== 'ok') throw new SessionStoreUnavailableError(state.unavailable);
    if (!cached.sessionToken) cached.sessionToken = token;
    if (materializeLegacyAnchor(cached)) {
      storeSession(cached).catch((err) =>
        console.error('[tokenCache] Failed to persist legacy absolute anchor (getSessionByToken):', err)
      );
    }
    return cached;
  }

  // Cold path: storage lookup (O(1) with new RowKey format)
  let stored: UserSession | null;
  try {
    stored = await loadSessionByToken(token);
  } catch (err) {
    console.error('[tokenCache] Failed to load session by token from Table Storage:', err);
    throw new SessionStoreUnavailableError(err);
  }
  if (!stored) return undefined;
  if (materializeLegacyAnchor(stored)) {
    storeSession(stored).catch((err) =>
      console.error('[tokenCache] Failed to persist legacy absolute anchor (getSessionByToken storage):', err)
    );
  }
  cachePut(hash, stored);
  // Update userId index so subsequent getSession/getValidAccessToken
  // calls in this request context use THIS session
  userIndex.set(stored.userId, hash);
  return stored;
}

/**
 * Delete a specific session by its storage key. Used when a single session's
 * token refresh fails — only that session is removed, not all sessions for
 * the user.
 */
export async function deleteSessionByKey(storageKey: string, userId: string): Promise<void> {
  // Remove the targeted session from the in-memory cache, regardless of which
  // session the userId index currently points to. Concurrent sessions for the
  // same user are cached under distinct token-hash keys, so we locate the entry
  // by its _storageKey rather than assuming it is the indexed one. (Previously
  // this only deleted when userIndex happened to point at the targeted session,
  // leaving a non-indexed session — e.g. one whose refresh just failed — live
  // in the cache.)
  for (const [cacheEntryKey, cached] of sessionCache) {
    if (cached.userId === userId && cached._storageKey === storageKey) {
      sessionCache.delete(cacheEntryKey);
      verifiedAt.delete(cacheEntryKey);
      // Only drop the userId index if it pointed at THIS session; a different
      // concurrent session for the same user must stay indexed.
      if (userIndex.get(userId) === cacheEntryKey) {
        userIndex.delete(userId);
      }
      break;
    }
  }
  try {
    await removeSessionByKey(storageKey);
  } catch (err) {
    console.error('[tokenCache] Failed to delete session from Table Storage:', err);
  }
}

/**
 * Delete ALL sessions for a user. Used for logout and admin "kick user"
 * operations. Also exported as deleteSession for backward compatibility.
 *
 * Evicts every cached session for the user on this replica, not only the
 * indexed one. Other replicas drop theirs on the next revalidation
 * (SESSION_REVALIDATE_MS), and cannot write a deleted row back meanwhile.
 */
export async function deleteAllUserSessions(userId: string): Promise<void> {
  for (const [key, cached] of sessionCache) {
    if (cached.userId === userId) cacheEvict(key);
  }
  userIndex.delete(userId);
  try {
    await removeSession(userId);
  } catch (err) {
    console.error('[tokenCache] Failed to delete sessions from Table Storage:', err);
  }
}

/**
 * Returns a valid access token for a specific session, refreshing silently
 * if the token is within 5 minutes of expiry. This is the request-scoped
 * version — callers that have the authenticated session should use this
 * to avoid the global userIndex.
 */
export async function getValidAccessTokenForSession(session: UserSession): Promise<string> {
  const bufferMs = 5 * 60 * 1000;
  if (Date.now() + bufferMs < session.expiresAt) {
    return session.accessToken;
  }

  // Token is expired or expiring soon — refresh silently via MSAL
  try {
    const { accessToken, expiresOn } = await acquireTokenSilent(session.homeAccountId);

    const updated: UserSession = {
      ...session,
      accessToken,
      expiresAt: expiresOn ? expiresOn.getTime() : Date.now() + 3_600_000,
      sessionCreatedAt: Date.now(),
    };
    await storeSession(updated);
    console.log(`[tokenCache] Access token refreshed for user ${session.userId}, session TTL extended`);
    return updated.accessToken;
  } catch (err) {
    if (err instanceof SessionRowMissingError) throw err;
    // MSAL silent refresh failed. Do NOT delete the session here.
    //
    // Deleting it permanently unbinds the client's long-lived session token: the
    // MCP client keeps sending the same Bearer token, while a browser re-login at
    // /api/auth/login mints a BRAND-NEW session token the client never receives.
    // Once the row is gone, no successful web re-auth can re-bind the client's
    // existing token, so the tool path stays "Session expired" forever even though
    // the web session is valid — the exact desync was hit on your-mcp-host.example.com.
    //
    // Retaining the session lets it self-heal: a later web re-auth refreshes the
    // shared MSAL account (keyed by the same homeAccountId) in Table Storage, after
    // which the next call here silently repairs THIS session in place and returns a
    // fresh token — no client reconfiguration required. This mirrors the already
    // non-destructive /api/auth/refresh path. The absolute-lifetime bound
    // still forces a full re-OAuth via authenticateRequest, so a stale/stolen token
    // cannot be renewed indefinitely — retention is not a security regression.
    //
    // Log the underlying MSAL cause: a refresh that fails silently and destroys
    // recoverable state is a write path with no consumer.
    console.error(
      `[tokenCache] Silent token refresh FAILED for user ${session.userId} ` +
      `(homeAccountId ${session.homeAccountId}): ` +
      `${err instanceof Error ? err.message : String(err)} — session retained for ` +
      `recovery after re-auth.`
    );
    throw new Error(`Token refresh failed for user ${session.userId}. Re-authentication required.`);
  }
}

/**
 * Returns a valid access token for the given user. Falls back to the
 * userIndex to find a session — use getValidAccessTokenForSession when
 * you have the authenticated session available.
 */
export async function getValidAccessToken(userId: string): Promise<string> {
  const session = await getSession(userId);
  if (!session) {
    throw new Error(`No session found for user ${userId}. Re-authentication required.`);
  }
  return getValidAccessTokenForSession(session);
}

/**
 * Get tenantId directly from a session (request-scoped, no userIndex lookup).
 */
export function getTenantIdFromSession(session: UserSession): string {
  if (!session.tenantId) throw new Error(`No tenantId in session for user ${session.userId}. Please re-authenticate.`);
  return session.tenantId;
}

/**
 * Get tenantId by userId. Falls back to userIndex — prefer getTenantIdFromSession.
 */
export async function getTenantId(userId: string): Promise<string> {
  const session = await getSession(userId);
  if (!session) throw new Error(`No session for user ${userId}`);
  return getTenantIdFromSession(session);
}

export async function listActiveSessions(): Promise<Array<{
  userId: string;
  email: string;
  displayName: string;
  expiresAt: number;
  sessionCreatedAt: number;
  sessionAbsoluteCreatedAt?: number;
}>> {
  try {
    const sessions = await listAllSessions();
    // Refresh memory cache with all sessions
    for (const s of sessions) {
      const key = cacheKey(s);
      cachePut(key, s);
      // Only update userIndex if no entry exists (don't override active session)
      if (!userIndex.has(s.userId)) {
        userIndex.set(s.userId, key);
      }
    }
    return sessions.map((s) => ({
      userId: s.userId,
      email: s.email,
      displayName: s.displayName,
      expiresAt: s.expiresAt,
      sessionCreatedAt: s.sessionCreatedAt || 0,
      sessionAbsoluteCreatedAt: s.sessionAbsoluteCreatedAt,
    }));
  } catch (err) {
    console.error('[tokenCache] Failed to list sessions from Table Storage:', err);
    return Array.from(sessionCache.values()).map((s) => ({
      userId: s.userId,
      email: s.email,
      displayName: s.displayName,
      expiresAt: s.expiresAt,
      sessionCreatedAt: s.sessionCreatedAt || 0,
      sessionAbsoluteCreatedAt: s.sessionAbsoluteCreatedAt,
    }));
  }
}

/** @deprecated Use deleteAllUserSessions or deleteSessionByKey instead. */
export const deleteSession = deleteAllUserSessions;
