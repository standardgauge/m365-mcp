import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { authenticateRequestAllowExpired } from '../../services/authMiddleware.js';
import { acquireTokenSilent } from '../../services/graphClient.js';
import { storeSession } from '../../services/tokenCache.js';
import type { UserSession } from '../../services/tokenCache.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { auditActor, logAccess } from '../../services/auditLog.js';

/**
 * POST /api/auth/refresh
 *
 * Silently refreshes an expired (or about-to-expire) session using MSAL's
 * cached refresh token. The caller must present a valid session token (via
 * the same mechanisms as any other authenticated request), but the session's
 * TTL is not enforced — an expired session is accepted as long as the token
 * hash matches.
 *
 * On success the session's `sessionCreatedAt` is reset to `Date.now()`,
 * extending the inactivity TTL, and fresh Graph tokens are stored.
 * The immutable `sessionAbsoluteCreatedAt` is never reset — it anchors the
 * absolute session lifetime enforced by authenticateRequestAllowExpired.
 * Once SESSION_MAX_LIFETIME_MS is exceeded even this endpoint returns 401.
 *
 * If the MSAL silent acquisition fails (e.g. refresh token revoked or truly
 * expired), a 401 is returned with a `loginUrl` the client can redirect to.
 */
async function refresh(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateRequestAllowExpired(request);
    if (!auth) {
      return {
        status: 401,
        jsonBody: { refreshed: false, loginUrl: '/api/auth/login' },
      };
    }

    const { session } = auth;

    // Attempt silent token acquisition via MSAL cached refresh token
    try {
      const { accessToken, expiresOn } = await acquireTokenSilent(session.homeAccountId);

      const updated: UserSession = {
        ...session,
        accessToken,
        expiresAt: expiresOn ? expiresOn.getTime() : Date.now() + 3_600_000,
        sessionCreatedAt: Date.now(), // Reset TTL
      };

      await storeSession(updated);
      logAccess({ ...auditActor(session), operation: 'auth.refresh', result: 'allowed', source: 'http' });
      context.log(`[refresh] Session refreshed for user ${session.userId}`);

      return {
        status: 200,
        jsonBody: { refreshed: true },
      };
    } catch (msalErr: unknown) {
      const message = msalErr instanceof Error ? msalErr.message : 'Unknown MSAL error';
      context.warn(`[refresh] Silent token acquisition failed for user ${session.userId}: ${message}`);
      logAccess({
        ...auditActor(session),
        operation: 'auth.refresh',
        result: 'denied',
        reason: 'silent token acquisition failed',
        source: 'http',
      });

      return {
        status: 401,
        jsonBody: { refreshed: false, loginUrl: '/api/auth/login' },
      };
    }
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Unknown error';
    context.error('[refresh] Unexpected error:', message);
    return { status: 500, jsonBody: { error: 'Session refresh failed' } };
  }
}

app.http('refresh', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/auth/refresh',
  handler: withSecurity(refresh),
});
