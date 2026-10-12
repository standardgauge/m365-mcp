import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { authenticateConsoleRequest, authorizeAdmin } from '../../services/authMiddleware.js';
import { collectReadinessChecks } from '../../services/readiness.js';
import { withSecurity } from '../../services/securityHeaders.js';

/**
 * GET /api/manage/readiness
 *
 * Reports the post-deploy operator checks (services/readiness.ts) so the admin
 * console can show that a migration window is still open instead of relying
 * on the operator remembering. Global Administrator only: the report describes
 * which protections are not yet enforced, which is not for every signed-in
 * user to read.
 */
async function getReadiness(
  request: HttpRequest,
  _context: InvocationContext
): Promise<HttpResponseInit> {
  const auth = await authenticateConsoleRequest(request);
  if (!auth) {
    return { status: 401, jsonBody: { error: 'Authentication required' } };
  }

  const isAdmin = await authorizeAdmin(auth, 'admin.readiness.read');
  if (!isAdmin) {
    return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
  }

  return {
    status: 200,
    jsonBody: collectReadinessChecks(),
    headers: { 'Cache-Control': 'no-store' },
  };
}

app.http('getReadiness', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/manage/readiness',
  handler: withSecurity(getReadiness),
});
