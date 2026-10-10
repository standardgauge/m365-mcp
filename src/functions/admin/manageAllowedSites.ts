import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getAllowedSites, setAllowedSites } from '../../services/serviceSettings.js';
import { getTenantId } from '../../services/tokenCache.js';
import { authenticateConsoleRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function manageAllowedSites(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateConsoleRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }
    const userId = auth.userId;
    const tenantId = await getTenantId(userId);

    if (request.method === 'GET') {
      const allowedSites = await getAllowedSites(tenantId);
      return { status: 200, jsonBody: { allowedSites } };
    }

    // POST requires Global Admin
    const isAdmin = await checkGlobalAdmin(userId);
    if (!isAdmin) {
      return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
    }

    const body = await request.json() as { allowedSites?: Array<{ id: string; name: string }> };
    if (!Array.isArray(body.allowedSites)) {
      return { status: 400, jsonBody: { error: 'Missing allowedSites array' } };
    }
    await setAllowedSites(tenantId, body.allowedSites);
    return { status: 200, jsonBody: { ok: true } };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('manageAllowedSites error:', message);
    return { status: 500, jsonBody: { error: message } };
  }
}

app.http('manageAllowedSites', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'api/manage/allowed-sites',
  handler: withSecurity(manageAllowedSites),
});
