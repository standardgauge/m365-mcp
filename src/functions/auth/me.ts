import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { authenticateRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
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
 */
async function me(
  request: HttpRequest,
  _context: InvocationContext,
): Promise<HttpResponseInit> {
  const auth = await authenticateRequest(request);
  if (!auth) {
    return {
      status: 401,
      jsonBody: { authenticated: false, loginUrl: '/api/auth/login' },
    };
  }

  const isGlobalAdmin = await checkGlobalAdmin(auth.userId);

  return {
    status: 200,
    jsonBody: {
      authenticated: true,
      userId: auth.userId,
      displayName: auth.session.displayName,
      email: auth.session.email,
      isGlobalAdmin,
    },
  };
}

app.http('authMe', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/auth/me',
  handler: withSecurity(me),
});
