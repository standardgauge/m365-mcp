import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getMailConfig, setMailConfig } from '../../services/userMailConfig.js';
import { getTenantId } from '../../services/tokenCache.js';
import { authenticateRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

/**
 * Admin endpoint for managing per-user mail config.
 *
 * GET  /api/manage/mail-config?userId=<target>
 *   Returns the mail config for the target user.
 *   Requires Global Administrator role.
 *
 * POST /api/manage/mail-config
 *   Body: { userId: string, disable_mail_indexing: boolean }
 *   Sets or clears the mailbox indexing disable flag for the user.
 *   Requires Global Administrator role.
 *
 * Unlike /api/manage/user-services, this endpoint is admin-only for both
 * reads and writes — no self-service. IR-adjacent roles cannot opt themselves
 * back in; only an admin can change the flag.
 *
 * All flag changes are written to audit log (console → App Insights).
 */
async function manageMailConfig(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateRequest(request);
    if (!auth) {
      return { status: 401, jsonBody: { error: 'Authentication required' } };
    }

    const isAdmin = await checkGlobalAdmin(auth.userId);
    if (!isAdmin) {
      return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
    }

    const tenantId = await getTenantId(auth.userId);

    if (request.method === 'GET') {
      const targetUserId = request.query.get('userId');
      if (!targetUserId) {
        return { status: 400, jsonBody: { error: 'Missing userId query parameter' } };
      }
      const config = await getMailConfig(tenantId, targetUserId);
      return { status: 200, jsonBody: { userId: targetUserId, ...config } };
    }

    // POST — set or clear the flag
    const body = (await request.json()) as {
      userId?: string;
      disable_mail_indexing?: boolean;
    };

    if (!body.userId) {
      return { status: 400, jsonBody: { error: 'Missing userId in request body' } };
    }
    if (typeof body.disable_mail_indexing !== 'boolean') {
      return {
        status: 400,
        jsonBody: { error: 'Missing or invalid disable_mail_indexing field (must be boolean)' },
      };
    }

    await setMailConfig(
      tenantId,
      body.userId,
      { disable_mail_indexing: body.disable_mail_indexing },
      auth.userId,
    );

    console.log(
      `[audit] mail-config-set admin=${auth.userId} target=${body.userId} disable_mail_indexing=${body.disable_mail_indexing} ts=${new Date().toISOString()}`,
    );

    context.log(
      `[admin] mail config updated: admin=${auth.userId} target=${body.userId} disable_mail_indexing=${body.disable_mail_indexing}`,
    );

    return {
      status: 200,
      jsonBody: { ok: true, userId: body.userId, disable_mail_indexing: body.disable_mail_indexing },
    };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('manageMailConfig error:', message);
    return { status: 500, jsonBody: { error: message } };
  }
}

app.http('manageMailConfig', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'api/manage/mail-config',
  handler: withSecurity(manageMailConfig),
});
