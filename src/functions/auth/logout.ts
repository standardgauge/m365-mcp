import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { authenticateRequest } from '../../services/authMiddleware.js';
import { deleteAllUserSessions } from '../../services/tokenCache.js';
import { withSecurity } from '../../services/securityHeaders.js';

/**
 * GET /api/auth/logout
 *
 * Ends the server session: deletes the stored session(s) for the user and
 * expires the browser cookies (`mcp_session`, `user_id`, `user_name`), then
 * redirects to the Microsoft identity-platform logout so the upstream SSO
 * session is cleared too. Post-logout the user lands back on the admin app,
 * where /api/auth/me will 401 and trigger a fresh sign-in.
 *
 * Replaces the SPA's old MSAL `logoutRedirect()` — sign-out now clears the
 * server session that actually authorizes API calls, not just client cache.
 */
async function logout(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  // Best-effort: drop the server-side session if we can identify it.
  try {
    const auth = await authenticateRequest(request);
    if (auth) await deleteAllUserSessions(auth.userId);
  } catch (err) {
    context.warn('[logout] session cleanup failed:', err instanceof Error ? err.message : err);
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
  const proto = request.headers.get('x-forwarded-proto') ?? 'https';
  const host = request.headers.get('x-forwarded-host') ?? request.headers.get('host') ?? '';
  const postLogoutRedirect = host ? `${proto}://${host}/admin` : '/admin';
  const location = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/logout` +
    `?post_logout_redirect_uri=${encodeURIComponent(postLogoutRedirect)}`;

  return {
    status: 302,
    headers: { Location: location },
    cookies: [
      { name: 'mcp_session', value: '', ...expired },
      // user_id / user_name are not HttpOnly (the SPA could read them), so expire them without httpOnly
      { name: 'user_id', value: '', secure: true, sameSite: 'Lax' as const, path: '/', maxAge: 0 },
      { name: 'user_name', value: '', secure: true, sameSite: 'Lax' as const, path: '/', maxAge: 0 },
    ],
  };
}

app.http('authLogout', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/auth/logout',
  handler: withSecurity(logout),
});
