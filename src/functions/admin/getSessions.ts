import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { listActiveSessions } from '../../services/tokenCache.js';
import { authenticateConsoleRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function getSessions(
  request: HttpRequest,
  _context: InvocationContext
): Promise<HttpResponseInit> {
  // Require authentication
  const auth = await authenticateConsoleRequest(request);
  if (!auth) {
    return { status: 401, jsonBody: { error: 'Authentication required' } };
  }

  // Require Global Admin role
  const isAdmin = await checkGlobalAdmin(auth.userId);
  if (!isAdmin) {
    return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
  }

  const sessions = await listActiveSessions();
  // Strip sensitive fields — never expose access tokens or session tokens
  const safe = sessions.map((s) => ({
    userId: s.userId,
    email: s.email,
    displayName: s.displayName,
    expiresAt: s.expiresAt,
  }));
  return { status: 200, jsonBody: { sessions: safe } };
}

app.http('getSessions', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/manage/sessions',
  handler: withSecurity(getSessions),
});
