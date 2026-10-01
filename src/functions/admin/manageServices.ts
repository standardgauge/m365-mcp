import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getEnabledServices, setEnabledServices, getReadOnlyServices, setReadOnlyServices } from '../../services/serviceSettings.js';
import { getTenantId } from '../../services/tokenCache.js';
import { authenticateRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function manageServices(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }
    const userId = auth.userId;
    const tenantId = await getTenantId(userId);

    if (request.method === 'GET') {
      const [enabledServices, readOnlyServices] = await Promise.all([
        getEnabledServices(tenantId),
        getReadOnlyServices(tenantId),
      ]);
      return { status: 200, jsonBody: { enabledServices, readOnlyServices } };
    }

    // POST requires Global Admin
    const isAdmin = await checkGlobalAdmin(userId);
    if (!isAdmin) {
      return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
    }

    const body = await request.json() as { enabledServices?: string[]; readOnlyServices?: string[] };
    // Both fields are optional so an admin can update either list independently,
    // but at least one must be present.
    if (body.enabledServices === undefined && body.readOnlyServices === undefined) {
      return { status: 400, jsonBody: { error: 'Provide enabledServices and/or readOnlyServices array' } };
    }
    if (body.enabledServices !== undefined) {
      if (!Array.isArray(body.enabledServices)) {
        return { status: 400, jsonBody: { error: 'enabledServices must be an array' } };
      }
      await setEnabledServices(tenantId, body.enabledServices);
    }
    if (body.readOnlyServices !== undefined) {
      if (!Array.isArray(body.readOnlyServices)) {
        return { status: 400, jsonBody: { error: 'readOnlyServices must be an array' } };
      }
      await setReadOnlyServices(tenantId, body.readOnlyServices);
    }
    return { status: 200, jsonBody: { ok: true } };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('manageServices error:', message);
    return { status: 500, jsonBody: { error: message } };
  }
}

app.http('manageServices', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'api/manage/services',
  handler: withSecurity(manageServices),
});
