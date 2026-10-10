import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { authenticateRequest } from '../../services/authMiddleware.js';
import { deleteAllUserSessions } from '../../services/tokenCache.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { checkBrowserOrigin, expiredConsoleCookie } from '../../services/consoleSession.js';
import { resolveFrontendUrl } from '../../services/frontendUrl.js';
import { auditActor, logAccess } from '../../services/auditLog.js';
import { SessionStoreUnavailableError, SESSION_STORE_RETRY_AFTER_S } from '../../services/sessionStoreError.js';

/**
 * POST /api/auth/logout
 *
 * Ends the server session: deletes the stored session(s) for the user and
 * expires the browser cookies (`mcp_session`, `mcp_console`, `user_id`,
 * `user_name`), then
 * redirects to the Microsoft identity-platform logout so the upstream SSO
 * session is cleared too. Post-logout the user lands back on the admin app,
 * at the site root (never /admin, which the Functions host reserves), where
 * /api/auth/me will 401 and trigger a fresh sign-in.
 *
 * Replaces the SPA's old MSAL `logoutRedirect()` — sign-out now clears the
 * server session that actually authorizes API calls, not just client cache.
 *
 * If session storage cannot be reached the session cannot be identified or
 * ended, so the response is a 503 with the cookies untouched and the user can
 * try again.
 *
 * POST with an Origin check, not GET: a GET logout let any page, or a link,
 * sign the user out cross-site (threat model 3.4). The SPA submits a form.
 */
async function logout(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  const origin = checkBrowserOrigin(request);
  if (!origin.ok) {
    context.warn(`[logout] refused: ${origin.reason}`);
    return { status: 403, jsonBody: { error: 'Forbidden' } };
  }

  // Best-effort: drop the server-side session if we can identify it. Only a
  // logout that found a session is audited; the route is anonymous, so
  // recording the rest would let anyone write rows.
  let auth: Awaited<ReturnType<typeof authenticateRequest>> = null;
  try {
    auth = await authenticateRequest(request);
    if (auth) {
      await deleteAllUserSessions(auth.userId);
      logAccess({ ...auditActor(auth.session), operation: 'auth.logout', result: 'allowed', source: 'http' });
    }
  } catch (err) {
    // Storage could not say whose session this is, so it cannot be ended.
    // Answer 503 and leave the cookies alone: redirecting as if signed out
    // would leave a live server session behind and drop the cookie the user
    // needs to retry.
    if (err instanceof SessionStoreUnavailableError) {
      context.error('[logout] session store unavailable; session not ended:', err.cause);
      return { status: 503, headers: { 'Retry-After': String(SESSION_STORE_RETRY_AFTER_S) }, jsonBody: { error: 'Session store unavailable, retry shortly' } };
    }
    context.warn('[logout] session cleanup failed:', err instanceof Error ? err.message : err);
    if (auth) {
      logAccess({
        ...auditActor(auth.session),
        operation: 'auth.logout',
        result: 'denied',
        reason: 'session cleanup failed; the session may still be valid',
        source: 'http',
      });
    }
  }

  const expired = {
    httpOnly: true,
    secure: true,
    sameSite: 'Lax' as const,
    path: '/',
    maxAge: 0,
  };

  // Build the Microsoft logout URL with a post-logout redirect back to the app.
  const tenantId = process.env.AZURE_TENANT_ID ?? 'common';
  // Same target the callback redirects to after sign-in: the SPA lives at the
  // site root, and resolveFrontendUrl refuses a path the Functions host
  // reserves (/admin, /runtime). Microsoft needs an absolute URI, so a
  // site-relative value is anchored on this request's origin.
  const frontendUrl = resolveFrontendUrl(process.env.FRONTEND_URL, (m) => context.error(m));
  const proto = request.headers.get('x-forwarded-proto') ?? 'https';
  const host = request.headers.get('x-forwarded-host') ?? request.headers.get('host') ?? '';
  const postLogoutRedirect = /^https?:\/\//i.test(frontendUrl) || !host
    ? frontendUrl
    : new URL(frontendUrl, `${proto}://${host}`).toString();
  const location = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/logout` +
    `?post_logout_redirect_uri=${encodeURIComponent(postLogoutRedirect)}`;

  // 303: the browser follows a POST's redirect with a GET.
  return {
    status: 303,
    headers: { Location: location },
    cookies: [
      { name: 'mcp_session', value: '', ...expired },
      expiredConsoleCookie(),
      // user_id / user_name are not HttpOnly (the SPA could read them), so expire them without httpOnly
      { name: 'user_id', value: '', secure: true, sameSite: 'Lax' as const, path: '/', maxAge: 0 },
      { name: 'user_name', value: '', secure: true, sameSite: 'Lax' as const, path: '/', maxAge: 0 },
    ],
  };
}

app.http('authLogout', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/auth/logout',
  handler: withSecurity(logout),
});
