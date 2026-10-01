import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getUserServiceOverrides, setUserServiceOverrides } from '../../services/userServiceOverrides.js';
import { ALL_SERVICE_KEYS } from '../../services/serviceSettings.js';
import { getTenantId } from '../../services/tokenCache.js';
import { authenticateRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

/**
 * Endpoint for managing per-user service overrides.
 *
 * GET  /api/manage/user-services?userId=<target>
 *   Returns the disabled services for the target user.
 *   - Global Admins can query any user.
 *   - Non-admins can only query themselves.
 *
 * POST /api/manage/user-services
 *   Body: { userId: string, disabledServices: string[] }
 *   Sets the disabled services for a user.
 *   - Global Admins can set overrides for any user.
 *   - Non-admins can only set their own overrides (self-service opt-out).
 *   Pass an empty array to clear all overrides.
 */
async function manageUserServices(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }

    const tenantId = await getTenantId(auth.userId);
    const isAdmin = await checkGlobalAdmin(auth.userId);

    if (request.method === 'GET') {
      const targetUserId = request.query.get('userId');
      if (!targetUserId) {
        return { status: 400, jsonBody: { error: 'Missing userId query parameter' } };
      }
      // Non-admins can only query their own overrides
      if (!isAdmin && targetUserId !== auth.userId) {
        return { status: 403, jsonBody: { error: 'You can only view your own service overrides' } };
      }
      const disabledServices = await getUserServiceOverrides(tenantId, targetUserId);
      return { status: 200, jsonBody: { userId: targetUserId, disabledServices } };
    }

    // POST — set disabled services for a user
    const body = (await request.json()) as { userId?: string; disabledServices?: string[] };
    if (!body.userId) {
      return { status: 400, jsonBody: { error: 'Missing userId in request body' } };
    }
    if (!Array.isArray(body.disabledServices)) {
      return { status: 400, jsonBody: { error: 'Missing disabledServices array in request body' } };
    }

    // Non-admins can only modify their own overrides
    if (!isAdmin && body.userId !== auth.userId) {
      return { status: 403, jsonBody: { error: 'You can only modify your own service overrides' } };
    }

    // Validate service keys
    const invalid = body.disabledServices.filter((s) => !ALL_SERVICE_KEYS.includes(s));
    if (invalid.length > 0) {
      return { status: 400, jsonBody: { error: `Invalid service key(s): ${invalid.join(', ')}` } };
    }

    await setUserServiceOverrides(tenantId, body.userId, body.disabledServices);
    return { status: 200, jsonBody: { ok: true, userId: body.userId, disabledServices: body.disabledServices } };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('manageUserServices error:', message);
    return { status: 500, jsonBody: { error: message } };
  }
}

app.http('manageUserServices', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'api/manage/user-services',
  handler: withSecurity(manageUserServices),
});
