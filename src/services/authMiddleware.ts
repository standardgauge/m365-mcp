import { HttpRequest } from '@azure/functions';
import { getSessionByToken, SESSION_TTL_MS, storeSession, isAbsoluteLifetimeExceeded } from './tokenCache.js';
import type { UserSession } from './tokenCache.js';
import { createGraphClient, acquireTokenSilent } from './graphClient.js';
import { getValidAccessToken } from './tokenCache.js';
import { checkBrowserOrigin, readCookie, verifyConsoleToken, CONSOLE_COOKIE } from './consoleSession.js';
import type { ConsoleClaims } from './consoleSession.js';
import { auditActor, logAccess } from './auditLog.js';

// Well-known roleTemplateId for the Global Administrator built-in role
const GLOBAL_ADMIN_ROLE_TEMPLATE_ID = '62e90394-69f5-4237-9190-012177145e10';

export interface AuthResult {
  userId: string;
  session: UserSession;
}

type TokenSource = 'bearer' | 'header' | 'cookie';

interface TokenCandidate {
  token: string;
  source: TokenSource;
}

/**
 * Extract all candidate session tokens from the request. Returns them in
 * priority order so authenticateRequest can try each until one matches.
 *
 * Multiple candidates are needed because the admin UI may send an
 * Authorization: Bearer header containing a Graph access token (not a session
 * token) alongside the mcp_session cookie that IS the session token.
 *
 * The cookie is a candidate only when the request passes the browser Origin
 * check. A browser attaches the cookie on its own, including to a request a
 * page on a same-site sibling subdomain makes, and SameSite=Lax does not stop
 * that; the Origin check does (threat model 3.4). The headers need no such
 * check: a page on another origin cannot set them without a CORS preflight,
 * and no route answers one.
 */
function extractSessionTokenCandidates(request: HttpRequest): TokenCandidate[] {
  const candidates: TokenCandidate[] = [];

  // 1. Authorization header
  const authHeader = request.headers.get('authorization') ?? '';
  if (authHeader.startsWith('Bearer ')) {
    candidates.push({ token: authHeader.slice(7), source: 'bearer' });
  }

  // 2. x-session-token header
  const headerToken = request.headers.get('x-session-token');
  if (headerToken) candidates.push({ token: headerToken, source: 'header' });

  // 3. mcp_session cookie, from this instance's own pages only
  const cookieToken = readCookie(request, 'mcp_session');
  if (cookieToken) {
    const origin = checkBrowserOrigin(request);
    if (origin.ok) {
      candidates.push({ token: cookieToken, source: 'cookie' });
    } else {
      console.warn(`[auth] mcp_session cookie ignored: ${origin.reason}`);
    }
  }

  return candidates;
}

/**
 * Authenticate an incoming request by validating the per-user session token
 * issued at OAuth callback time. The token is unguessable (32 random bytes),
 * stored server-side in Azure Table Storage, and looked up by token value —
 * never by client-supplied userId.
 *
 * Tokens are accepted from (in priority order):
 *   1. Authorization: Bearer <token>   — for headless / supergateway clients
 *   2. x-session-token: <token>        — alternate header for tools that
 *                                        reserve Authorization for Graph
 *   3. mcp_session=<token> cookie      — for browser clients (admin SPA),
 *                                        and only past the Origin check
 *
 * The admin API does not use this: /api/manage/* goes through
 * authenticateConsoleRequest, which accepts the cookie only.
 *
 * No shared secrets, no x-user-id trust, no fallback paths.
 *
 * Throws SessionStoreUnavailableError when storage cannot say whether a
 * candidate is a session. That stops the loop: the remaining candidates are
 * not tried, because the unanswered one may have been a session that should
 * have been refused. Callers turn it into a 503, never into an auth result.
 */
export async function authenticateRequest(request: HttpRequest): Promise<AuthResult | null> {
  const candidates = extractSessionTokenCandidates(request);
  for (const { token } of candidates) {
    const result = await authenticateToken(token);
    if (result === 'unknown') continue;
    return result;
  }
  console.warn(
    `[auth] No matching session found (${candidates.length} token candidate(s) tried)`
  );
  return null;
}

/**
 * Validate one session token: absolute lifetime, renewal window with silent
 * refresh, sliding-window touch. 'unknown' means no session has this token, so
 * the caller may try the next candidate; null means the session exists but may
 * not be used. Storage failures propagate as SessionStoreUnavailableError.
 */
async function authenticateToken(token: string): Promise<AuthResult | null | 'unknown'> {
  const session = await getSessionByToken(token);
  if (!session) return 'unknown';

  // Reject sessions past absolute lifetime before any refresh/touch —
  // prevents silent-renewal of a stolen token through normal request paths.
  if (isAbsoluteLifetimeExceeded(session)) {
    console.warn(
      `[auth] Session for user ${session.userId} exceeded absolute lifetime — forcing re-OAuth`
    );
    return null;
  }

  const now = Date.now();
  const elapsed = now - (session.sessionCreatedAt || 0);

  // ── Renewal window passed — renew through MSAL to continue ──
  // Not an idle timeout: an idle session is renewed here, not ended. Only
  // a failed silent refresh (refresh token revoked or expired) ends it.
  // The absolute cap above is what bounds a stolen token.
  if (session.sessionCreatedAt && elapsed > SESSION_TTL_MS) {
    try {
      const { accessToken, expiresOn } = await acquireTokenSilent(session.homeAccountId);
      const refreshed: UserSession = {
        ...session,
        accessToken,
        expiresAt: expiresOn ? expiresOn.getTime() : now + 3_600_000,
        sessionCreatedAt: now,
      };
      await storeSession(refreshed);
      logAccess({
        ...auditActor(session),
        operation: 'auth.session_renew',
        result: 'allowed',
        reason: `idle ${Math.round(elapsed / 60_000)}m`,
        source: 'http',
      });
      console.log(
        `[auth] Session renewal succeeded for user ${session.userId} (unused ${Math.round(elapsed / 60_000)}m)`
      );
      return { userId: refreshed.userId, session: refreshed };
    } catch (err) {
      logAccess({
        ...auditActor(session),
        operation: 'auth.session_renew',
        result: 'denied',
        reason: `silent token acquisition failed after idle ${Math.round(elapsed / 60_000)}m`,
        source: 'http',
      });
      console.error(
        `[auth] Session renewal FAILED for user ${session.userId} ` +
        `(unused ${Math.round(elapsed / 60_000)}m):`,
        err instanceof Error ? err.message : err
      );
      return null; // refresh token truly expired — caller gets 401
    }
  }

  // ── Sliding window: touch sessionCreatedAt past 50% TTL ──
  // Keeps active sessions alive without waiting for the TTL boundary.
  // Fire-and-forget to avoid adding latency to the request path.
  if (elapsed > SESSION_TTL_MS / 2) {
    const touched: UserSession = { ...session, sessionCreatedAt: now };
    storeSession(touched).catch((err) =>
      console.error('[auth] Failed to touch session TTL:', err)
    );
  }

  return { userId: session.userId, session };
}

export interface ConsoleAuthResult extends AuthResult {
  /** The verified console token's claims, for renewal by /api/auth/me. */
  console: ConsoleClaims;
  /** The cookie session token the console token is bound to. */
  sessionToken: string;
}

/**
 * Authenticate a request to the admin API (/api/manage/*). Browser only.
 *
 * Unlike authenticateRequest, this ignores the bearer and x-session-token
 * headers entirely: the session token comes from the `mcp_session` cookie,
 * and the request must also carry an `mcp_console` cookie bound to that token
 * (see services/consoleSession.ts) and pass the Origin check, and the session
 * must be a browser session. The install flow gives an MCP client a separate
 * client session, so the token it holds is not an admin-API credential, however
 * it is presented, and not even with a stolen console cookie beside it.
 * Threat model section 7, rows 7.3 and 3.4.
 *
 * Throws SessionStoreUnavailableError when storage cannot answer, as
 * authenticateRequest does.
 */
export async function authenticateConsoleRequest(request: HttpRequest): Promise<ConsoleAuthResult | null> {
  const origin = checkBrowserOrigin(request);
  if (!origin.ok) {
    console.warn(`[auth] console request refused: ${origin.reason}`);
    return null;
  }

  const sessionToken = readCookie(request, 'mcp_session');
  const claims = verifyConsoleToken(readCookie(request, CONSOLE_COOKIE), sessionToken);
  if (!sessionToken || !claims) {
    console.warn('[auth] console request refused: no valid console session');
    return null;
  }

  const result = await authenticateToken(sessionToken);
  if (!result || result === 'unknown') return null;
  // Only a browser session backs a console session. A row from before session
  // kinds existed may share its token with an installed client, so it is
  // refused too; the SPA sends the user through sign-in, which mints a
  // browser session.
  if (result.session.kind !== 'browser') {
    console.warn('[auth] console request refused: session is not a browser session');
    return null;
  }
  return { ...result, console: claims, sessionToken };
}

/**
 * Authenticate a request without the renewal-window check. The session must
 * exist and have a valid token hash; a session past SESSION_TTL_MS is
 * accepted so clients can renew it without a full re-login.
 *
 * The absolute session lifetime (SESSION_MAX_LIFETIME_MS) IS enforced here.
 * A session older than that bound returns null regardless of MSAL token
 * validity — forcing the client to start a fresh OAuth flow. This prevents
 * a stolen session token from being refreshed indefinitely while the MSAL
 * refresh token remains valid.
 *
 * The absolute anchor is sessionAbsoluteCreatedAt (immutable, set once at
 * OAuth time). For legacy sessions that predate this field, sessionCreatedAt
 * is used as a best-effort fallback; sessions with neither field set (both
 * zero/absent) are treated as past the cap and rejected.
 *
 * Throws SessionStoreUnavailableError when storage cannot answer, as
 * authenticateRequest does.
 */
export async function authenticateRequestAllowExpired(request: HttpRequest): Promise<AuthResult | null> {
  const candidates = extractSessionTokenCandidates(request);
  for (const { token } of candidates) {
    const session = await getSessionByToken(token);
    if (session) {
      if (isAbsoluteLifetimeExceeded(session)) {
        const anchor = session.sessionAbsoluteCreatedAt || session.sessionCreatedAt || 0;
        const age = anchor ? `${Math.round((Date.now() - anchor) / 86_400_000)}d old` : 'no anchor';
        console.warn(
          `[auth] Session for user ${session.userId} exceeded absolute lifetime ` +
          `(${age}) — forcing re-OAuth`
        );
        return null;
      }
      return { userId: session.userId, session };
    }
  }
  return null;
}

/**
 * Check whether the authenticated user holds the Global Administrator role
 * in their Entra ID tenant.
 */
export async function checkGlobalAdmin(userId: string, bearerToken?: string): Promise<boolean> {
  try {
    const accessToken = bearerToken ?? (await getValidAccessToken(userId));
    const graph = createGraphClient(accessToken);
    const result: { value?: Array<{ roleTemplateId?: string }> } = await graph
      .api('/me/transitiveMemberOf')
      .select('id,roleTemplateId')
      .get();
    return (result.value ?? []).some(
      (m) => m.roleTemplateId === GLOBAL_ADMIN_ROLE_TEMPLATE_ID
    );
  } catch {
    return false; // fail closed — deny if we cannot verify
  }
}

export const ADMIN_REQUIRED = 'Global Administrator role required';

/**
 * Record that a signed-in caller was refused an operation that needs Global
 * Administrator. `operation` is the name the same change is audited under when
 * it succeeds, so a refusal and a success sit side by side in the log.
 */
export function auditAdminRefusal(auth: AuthResult, operation: string, resource?: string): void {
  logAccess({
    ...auditActor(auth.session),
    operation,
    resource,
    result: 'denied',
    reason: ADMIN_REQUIRED,
    source: 'http',
  });
}

/** checkGlobalAdmin for an admin route, with the refusal audited. */
export async function authorizeAdmin(auth: AuthResult, operation: string, resource?: string): Promise<boolean> {
  const isAdmin = await checkGlobalAdmin(auth.userId);
  if (!isAdmin) auditAdminRefusal(auth, operation, resource);
  return isAdmin;
}

/**
 * Convenience: authenticate a console (admin SPA) request + require Global
 * Admin. Returns the auth result or null if the caller is not authenticated or
 * not an admin.
 */
export async function requireGlobalAdmin(request: HttpRequest): Promise<AuthResult | null> {
  const auth = await authenticateConsoleRequest(request);
  if (!auth) return null;

  const isAdmin = await checkGlobalAdmin(auth.userId);
  if (!isAdmin) return null;

  return auth;
}
