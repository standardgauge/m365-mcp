import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { randomBytes } from 'crypto';
import { authenticateConsoleRequest, checkGlobalAdmin, auditAdminRefusal } from '../../services/authMiddleware.js';
import { auditActor, auditSnapshot, logAccess } from '../../services/auditLog.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { getEnabledServices } from '../../services/serviceSettings.js';
import { getUserServiceOverrides } from '../../services/userServiceOverrides.js';
import { storeSession, deleteSessionByKey, getTenantIdFromSession } from '../../services/tokenCache.js';
import type { UserSession } from '../../services/tokenCache.js';
import {
  AGENT_KEY_LABEL_MAX,
  AGENT_KEY_MAX_LIFETIME_DAYS,
  agentKeyDeviceLabel,
  createAgentKey,
  findAgentKey,
  generateAgentKeyId,
  generateAgentKeyToken,
  getTenantAgentKeyPolicy,
  listAgentKeys,
  markAgentKeyRevoked,
  publicAgentKey,
  validateRequestedPolicy,
} from '../../services/agentKeys.js';
import type { AgentKeyPolicy } from '../../services/agentKeys.js';
import { SessionStoreUnavailableError, SESSION_STORE_RETRY_AFTER_S } from '../../services/sessionStoreError.js';

/**
 * Agent API keys (services/agentKeys.ts).
 *
 * GET    /api/manage/agent-keys                 the caller's keys
 * GET    /api/manage/agent-keys?all=true        every key in the tenant (Global Administrator)
 * GET    /api/manage/agent-keys?userId=<id>     one user's keys (Global Administrator, or self)
 * POST   /api/manage/agent-keys                 mint a key for the caller
 *        { label, services, readOnlyServices?, allowedSites?, expiresInDays? }
 *        → { key, ...publicAgentKey }   the key value appears in this response only
 * DELETE /api/manage/agent-keys?keyId=<id>      revoke; own keys, or any key as Global Administrator
 *
 * Browser-only, like every /api/manage route: a console session is required,
 * so neither an installed client's token nor an agent key can mint or list
 * keys (threat model 7.3). A key is bound to a fresh `client` session of the
 * caller that inherits their absolute-lifetime anchor, the same way the
 * install flow mints an installer's session.
 */
async function manageAgentKeys(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  try {
    const auth = await authenticateConsoleRequest(request);
    if (!auth) return { status: 401, jsonBody: { error: 'Authentication required' } };
    const tenantId = getTenantIdFromSession(auth.session);

    if (request.method === 'GET') {
      const all = request.query.get('all') === 'true';
      const targetUserId = request.query.get('userId');
      if (all || (targetUserId && targetUserId !== auth.userId)) {
        const isAdmin = await checkGlobalAdmin(auth.userId);
        if (!isAdmin) {
          auditAdminRefusal(auth, 'admin.agent_keys.read', all ? 'tenant' : `user:${targetUserId}`);
          return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
        }
      }
      const keys = await listAgentKeys(tenantId, all ? undefined : (targetUserId ?? auth.userId));
      return { status: 200, jsonBody: { keys: keys.map(publicAgentKey) } };
    }

    if (request.method === 'DELETE') {
      const keyId = request.query.get('keyId');
      if (!keyId) return { status: 400, jsonBody: { error: 'Missing keyId query parameter' } };
      const record = await findAgentKey(tenantId, keyId);
      if (!record) return { status: 404, jsonBody: { error: 'No such key' } };
      if (record.userId !== auth.userId) {
        const isAdmin = await checkGlobalAdmin(auth.userId);
        if (!isAdmin) {
          auditAdminRefusal(auth, 'policy.agent_key.revoke', `key:${keyId}`);
          return { status: 403, jsonBody: { error: 'You can only revoke your own keys' } };
        }
      }
      if (!record.revokedAt) {
        await markAgentKeyRevoked(record, auth.userId);
        await deleteSessionByKey(record.sessionStorageKey, record.userId);
        logAccess({
          ...auditActor(auth.session),
          operation: 'policy.agent_key.revoke',
          resource: `key:${keyId}`,
          result: 'allowed',
          source: 'http',
          before: auditSnapshot({ owner: record.userEmail, label: record.label, policy: record.policy }),
          after: auditSnapshot({ revoked: true }),
        });
      }
      return { status: 200, jsonBody: { ok: true, keyId } };
    }

    // POST — mint a key for the caller.
    const policy = await getTenantAgentKeyPolicy(tenantId);
    if (!policy.allowKeys) {
      logAccess({
        ...auditActor(auth.session),
        operation: 'policy.agent_key.create',
        result: 'denied',
        reason: 'agent keys are disabled for this tenant',
        source: 'http',
      });
      return { status: 403, jsonBody: { error: 'Agent keys are disabled for this organization' } };
    }

    const body = (await request.json().catch(() => null)) as {
      label?: unknown; services?: unknown; readOnlyServices?: unknown; allowedSites?: unknown; expiresInDays?: unknown;
    } | null;
    if (!body || typeof body !== 'object') return { status: 400, jsonBody: { error: 'JSON body required' } };

    const label = typeof body.label === 'string' ? body.label.trim() : '';
    if (!label) return { status: 400, jsonBody: { error: 'label is required' } };
    if (label.length > AGENT_KEY_LABEL_MAX) return { status: 400, jsonBody: { error: `label must be at most ${AGENT_KEY_LABEL_MAX} characters` } };
    if (/[#\r\n]/.test(label)) return { status: 400, jsonBody: { error: 'label may not contain "#" or line breaks' } };

    let expiresAt = 0;
    if (body.expiresInDays !== undefined && body.expiresInDays !== null) {
      const days = Number(body.expiresInDays);
      if (!Number.isInteger(days) || days < 1 || days > AGENT_KEY_MAX_LIFETIME_DAYS) {
        return { status: 400, jsonBody: { error: `expiresInDays must be an integer from 1 to ${AGENT_KEY_MAX_LIFETIME_DAYS}` } };
      }
      expiresAt = Date.now() + days * 86_400_000;
    }

    // What this user may grant: what they can reach themselves, less what the
    // tenant's key policy takes from every key.
    const [enabled, userDisabled] = await Promise.all([
      getEnabledServices(tenantId),
      getUserServiceOverrides(tenantId, auth.userId),
    ]);
    const grantable = enabled.filter((s) => !userDisabled.includes(s) && !policy.disabledServices.includes(s));
    const requested: Partial<AgentKeyPolicy> = {
      services: body.services as string[],
      readOnlyServices: body.readOnlyServices as string[] | undefined,
      allowedSites: body.allowedSites as AgentKeyPolicy['allowedSites'] | undefined,
    };
    const validated = validateRequestedPolicy(requested, grantable);
    if (!validated.ok) return { status: 400, jsonBody: { error: validated.error } };

    const keyId = generateAgentKeyId();
    const token = generateAgentKeyToken();
    const now = Date.now();
    const browser = auth.session;
    const keySession: UserSession = {
      userId: browser.userId,
      homeAccountId: browser.homeAccountId,
      displayName: browser.displayName,
      email: browser.email,
      tenantId: browser.tenantId,
      accessToken: browser.accessToken,
      expiresAt: browser.expiresAt,
      sessionToken: randomBytes(32).toString('hex'),
      sessionCreatedAt: now,
      sessionAbsoluteCreatedAt: browser.sessionAbsoluteCreatedAt || browser.sessionCreatedAt || now,
      deviceLabel: agentKeyDeviceLabel(label, keyId),
      kind: 'client',
    };
    await storeSession(keySession);
    if (!keySession._storageKey) throw new Error('key session was not persisted');

    const record = await createAgentKey(token, keyId, {
      tenantId,
      userId: auth.userId,
      userEmail: browser.email,
      label,
      sessionStorageKey: keySession._storageKey,
      policy: validated.policy,
      expiresAt,
      now,
    });
    logAccess({
      ...auditActor(auth.session),
      operation: 'policy.agent_key.create',
      resource: `key:${keyId}`,
      result: 'allowed',
      source: 'http',
      after: auditSnapshot({ label, policy: validated.policy, expiresAt: expiresAt || null }),
    });
    return {
      status: 201,
      headers: { 'Cache-Control': 'no-store' },
      jsonBody: { key: token, ...publicAgentKey(record) },
    };
  } catch (err: unknown) {
    if (err instanceof SessionStoreUnavailableError) {
      return { status: 503, headers: { 'Retry-After': String(SESSION_STORE_RETRY_AFTER_S) }, jsonBody: { error: 'Session store unavailable, retry shortly' } };
    }
    const message = err instanceof Error ? err.message : String(err);
    context.error('manageAgentKeys error:', message);
    return { status: 500, jsonBody: { error: 'Internal server error' } };
  }
}

app.http('manageAgentKeys', {
  methods: ['GET', 'POST', 'DELETE'],
  authLevel: 'anonymous',
  route: 'api/manage/agent-keys',
  handler: withSecurity(manageAgentKeys),
});
