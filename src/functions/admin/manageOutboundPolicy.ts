import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import {
  getOutboundEnforcement,
  getOutboundPolicy,
  setOutboundPolicy,
  OUTBOUND_CHANNELS,
  OUTBOUND_MODES,
  type OutboundMode,
  type OutboundModes,
} from '../../services/outboundPolicy.js';
import { getTenantId } from '../../services/tokenCache.js';
import { authenticateRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { logAccess } from '../../services/auditLog.js';

/**
 * Admin endpoint for the outbound policy: calendar invitations, comments on
 * meeting responses and Teams messages, the channels enforced draft mode does
 * not hold. See services/outboundPolicy.ts for what each mode does.
 *
 * GET  /api/manage/outbound-policy[?userId=<target>]
 *   Returns the tenant-wide policy. With ?userId, also returns that user's
 *   row and the effective mode per channel (the stricter of the two).
 *   Requires Global Administrator role.
 *
 * POST /api/manage/outbound-policy
 *   Body: { scope: 'tenant', calendarInvites?, eventResponses?, teamsMessages? }
 *      or { scope: 'user', userId: string, ...same channels }
 *   Each channel is 'allow' | 'internal' | 'block'; channels left out keep
 *   their stored value, and at least one must be given.
 *   Requires Global Administrator role.
 *
 * Admin-only for reads and writes, like the draft-mode policy: a user cannot
 * see or lift a restriction that applies to them. Every change is written to
 * the audit log table and echoed to the console audit stream.
 */
async function manageOutboundPolicy(
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
        const tenant = await getOutboundPolicy(tenantId, { scope: 'tenant' });
        return { status: 200, jsonBody: { tenant } };
      }
      const enforcement = await getOutboundEnforcement(tenantId, targetUserId);
      return { status: 200, jsonBody: { userId: targetUserId, ...enforcement } };
    }

    const body = (await request.json()) as Record<string, unknown>;

    if (body.scope !== 'tenant' && body.scope !== 'user') {
      return { status: 400, jsonBody: { error: 'scope must be "tenant" or "user"' } };
    }
    if (body.scope === 'user' && (typeof body.userId !== 'string' || body.userId.length === 0)) {
      return { status: 400, jsonBody: { error: 'Missing userId in request body for scope "user"' } };
    }
    const allowedKeys = new Set(['scope', 'userId', ...OUTBOUND_CHANNELS]);
    const unknown = Object.keys(body).filter((k) => !allowedKeys.has(k));
    if (unknown.length > 0) {
      return { status: 400, jsonBody: { error: `Unknown field(s): ${unknown.join(', ')}` } };
    }
    const modes: Partial<OutboundModes> = {};
    for (const channel of OUTBOUND_CHANNELS) {
      const value = body[channel];
      if (value === undefined) continue;
      if (typeof value !== 'string' || !(OUTBOUND_MODES as readonly string[]).includes(value)) {
        return { status: 400, jsonBody: { error: `${channel} must be one of ${OUTBOUND_MODES.join(', ')}` } };
      }
      modes[channel] = value as OutboundMode;
    }
    if (Object.keys(modes).length === 0) {
      return { status: 400, jsonBody: { error: `Give at least one of ${OUTBOUND_CHANNELS.join(', ')}` } };
    }

    const target = body.scope === 'tenant'
      ? { scope: 'tenant' as const }
      : { scope: 'user' as const, userId: body.userId as string };
    const policy = await setOutboundPolicy(tenantId, target, modes, auth.userId);

    const resource = target.scope === 'tenant' ? 'tenant' : `user:${target.userId}`;
    const change = Object.entries(modes).map(([k, v]) => `${k}=${v}`).join(' ');
    logAccess({
      tenantId,
      userId: auth.userId,
      userEmail: auth.session.email,
      deviceLabel: auth.session.deviceLabel,
      operation: 'set_outbound_policy',
      resource,
      result: 'allowed',
      reason: change,
      source: 'http',
    });
    console.log(`[audit] outbound-policy-set admin=${auth.userId} target=${resource} ${change} ts=${policy.updatedAt}`);
    context.log(`[admin] outbound policy updated: admin=${auth.userId} target=${resource} ${change}`);

    return {
      status: 200,
      jsonBody: {
        ok: true,
        scope: target.scope,
        ...(target.scope === 'user' ? { userId: target.userId } : {}),
        policy,
      },
    };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('manageOutboundPolicy error:', message);
    return { status: 500, jsonBody: { error: message } };
  }
}

app.http('manageOutboundPolicy', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'api/manage/outbound-policy',
  handler: withSecurity(manageOutboundPolicy),
});
