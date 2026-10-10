import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { authenticateConsoleRequest, authenticateRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { consoleCookie, renewConsoleToken } from '../../services/consoleSession.js';
import { withSecurity } from '../../services/securityHeaders.js';

/**
 * GET /api/auth/me
 *
 * Returns the identity of the currently authenticated user, resolved purely
 * from the server-side session (the `mcp_session` cookie / bearer token) —
 * never from a client-supplied userId.
 *
 * This is the entry point the admin SPA calls on load: a 200 means a valid
 * session cookie is present (so all other /api/* calls will authenticate via
 * the same cookie), and a 401 tells the SPA to start the server OAuth flow at
 * `loginUrl`. It replaces the old client-side MSAL flow, which only ever
 * obtained a Microsoft Graph token and never established the `mcp_session`
 * that every hardened endpoint requires.
 *
 * `isGlobalAdmin` is computed server-side using the session's stored Graph
 * token (via checkGlobalAdmin → /me/transitiveMemberOf), so the client no
 * longer needs to hold a Directory.Read.All token to unlock the admin view.
 *
 * `consoleSession` says whether the request also carries a valid console
 * session, the browser-only credential /api/manage/* requires
 * (services/consoleSession.ts). When it does, the console cookie is re-issued
 * with its idle expiry pushed out, so the SPA keeps it alive by calling this
 * endpoint. When it does not, the SPA sends the user back through sign-in,
 * which is the only place a console session is minted.
 */
async function me(
  request: HttpRequest,
  _context: InvocationContext,
): Promise<HttpResponseInit> {
  const consoleAuth = await authenticateConsoleRequest(request);
  const auth = consoleAuth ?? (await authenticateRequest(request));
  if (!auth) {
    return {
      status: 401,
      jsonBody: { authenticated: false, loginUrl: '/api/auth/login' },
    };
  }

  const isGlobalAdmin = await checkGlobalAdmin(auth.userId);

  const now = Date.now();
  const renewed = consoleAuth ? renewConsoleToken(consoleAuth.sessionToken, consoleAuth.console, now) : null;

  return {
    status: 200,
    headers: { 'Cache-Control': 'no-store' },
    jsonBody: {
      authenticated: true,
      userId: auth.userId,
      displayName: auth.session.displayName,
      email: auth.session.email,
      isGlobalAdmin,
      consoleSession: renewed !== null,
      ...(renewed ? { consoleExpiresAt: renewed.expiresAt } : {}),
    },
    ...(renewed ? { cookies: [consoleCookie(renewed, now)] } : {}),
  };
}

app.http('authMe', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/auth/me',
  handler: withSecurity(me),
});
