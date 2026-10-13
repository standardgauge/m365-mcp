import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { randomBytes } from 'crypto';
import { authenticateConsoleRequest, checkGlobalAdmin, auditAdminRefusal } from '../../services/authMiddleware.js';
import { auditActor, auditSnapshot, logAccess } from '../../services/auditLog.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { getEnabledServices } from '../../services/serviceSettings.js';
import { getUserServiceOverrides } from '../../services/userServiceOverrides.js';
import { storeSession, deleteSessionByKey, getSessionByStorageKey, getTenantIdFromSession } from '../../services/tokenCache.js';
import { createKeyExpiryReminder, deleteKeyExpiryReminder } from '../../services/agentKeyReminder.js';
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
  setAgentKeyReminder,
  updateAgentKey,
  isNarrowingOf,
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
 *        { label, services, expiresInDays, readOnlyServices?, allowedSites?, reminder? }
 *        → { key, ...publicAgentKey, reminder }   the key value appears in this response only
 * PATCH  /api/manage/agent-keys?keyId=<id>      edit label and policy in place
 *        { label?, services?, readOnlyServices?, allowedSites? }
 *        An owner may edit either way within what they could grant today,
 *        unless the tenant policy's ownersMayWiden is off, in which case
 *        only a narrowing is accepted. A Global Administrator may only
 *        narrow another user's key. The key value never changes.
 * DELETE /api/manage/agent-keys?keyId=<id>      revoke; own keys, or any key as Global Administrator
 *
 * Browser-only, like every /api/manage route: a console session is required,
 * so neither an installed client's token nor an agent key can mint or list
 * keys (threat model 7.3). A key is bound to a fresh `client` session of the
 * caller, minted the way the install flow mints an installer's session but
 * anchored at the key's creation: its absolute cap is the key's own lifetime,
 * chosen by the owner within the tenant's maximum, not the 30 days a sign-in
 * gets. With `reminder` (the default) a short event goes on the owner's own
 * calendar a few days before expiry (services/agentKeyReminder.ts).
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

    if (request.method === 'PATCH') {
      const keyId = request.query.get('keyId');
      if (!keyId) return { status: 400, jsonBody: { error: 'Missing keyId query parameter' } };
      const record = await findAgentKey(tenantId, keyId);
      if (!record) return { status: 404, jsonBody: { error: 'No such key' } };
      if (record.revokedAt) return { status: 409, jsonBody: { error: 'The key is revoked; mint a new one' } };
      const isOwner = record.userId === auth.userId;
      let isAdmin = false;
      if (!isOwner) {
        isAdmin = await checkGlobalAdmin(auth.userId);
        if (!isAdmin) {
          auditAdminRefusal(auth, 'policy.agent_key.update', `key:${keyId}`);
          return { status: 403, jsonBody: { error: 'You can only edit your own keys' } };
        }
      }

      const body = (await request.json().catch(() => null)) as {
        label?: unknown; services?: unknown; readOnlyServices?: unknown; allowedSites?: unknown;
      } | null;
      if (!body || typeof body !== 'object') return { status: 400, jsonBody: { error: 'JSON body required' } };
      if (body.label === undefined && body.services === undefined && body.readOnlyServices === undefined && body.allowedSites === undefined) {
        return { status: 400, jsonBody: { error: 'Provide label, services, readOnlyServices and/or allowedSites' } };
      }

      let label = record.label;
      if (body.label !== undefined) {
        label = typeof body.label === 'string' ? body.label.trim() : '';
        if (!label) return { status: 400, jsonBody: { error: 'label must not be empty' } };
        if (label.length > AGENT_KEY_LABEL_MAX) return { status: 400, jsonBody: { error: `label must be at most ${AGENT_KEY_LABEL_MAX} characters` } };
        if (/[#\r\n]/.test(label)) return { status: 400, jsonBody: { error: 'label may not contain "#" or line breaks' } };
      }

      const tenantPolicy = await getTenantAgentKeyPolicy(tenantId);
      // The owner's grantable set, as it stands today. An admin narrowing is
      // validated against the key's own current policy instead, since an
      // admin cannot speak for what the owner may reach.
      // A field the caller leaves out keeps its current value, pruned to the
      // new service set: dropping SharePoint drops its read-only flag and its
      // site picks with it, rather than failing validation on a leftover.
      const services = (body.services ?? record.policy.services) as string[];
      const inheritedReadOnly = Array.isArray(services) ? record.policy.readOnlyServices.filter((s) => services.includes(s)) : record.policy.readOnlyServices;
      const inheritedSites = Array.isArray(services) && !services.includes('sharepoint') ? [] : record.policy.allowedSites;
      const requested: Partial<AgentKeyPolicy> = {
        services,
        readOnlyServices: (body.readOnlyServices ?? inheritedReadOnly) as string[],
        allowedSites: (body.allowedSites ?? inheritedSites) as AgentKeyPolicy['allowedSites'],
      };
      let grantable: string[];
      if (isOwner) {
        const [enabled, userDisabled] = await Promise.all([
          getEnabledServices(tenantId),
          getUserServiceOverrides(tenantId, auth.userId),
        ]);
        grantable = enabled.filter((s) => !userDisabled.includes(s) && !tenantPolicy.disabledServices.includes(s));
      } else {
        grantable = record.policy.services;
      }
      const validated = validateRequestedPolicy(requested, grantable);
      if (!validated.ok) return { status: 400, jsonBody: { error: validated.error } };

      const narrowing = isNarrowingOf(validated.policy, record.policy);
      if (!narrowing && !isOwner) {
        auditAdminRefusal(auth, 'policy.agent_key.update', `key:${keyId}`);
        return { status: 403, jsonBody: { error: "An administrator can only narrow another user's key" } };
      }
      if (!narrowing && isOwner && !tenantPolicy.ownersMayWiden) {
        logAccess({
          ...auditActor(auth.session),
          operation: 'policy.agent_key.update',
          resource: `key:${keyId}`,
          result: 'denied',
          reason: 'tenant policy allows owners to narrow keys only',
          source: 'http',
        });
        return { status: 403, jsonBody: { error: 'Your organization only allows a key to be narrowed; mint a new key for wider access' } };
      }

      const updated = await updateAgentKey(record, { label, policy: validated.policy }, auth.userId);
      logAccess({
        ...auditActor(auth.session),
        operation: 'policy.agent_key.update',
        resource: `key:${keyId}`,
        result: 'allowed',
        reason: narrowing ? 'narrowed' : 'widened',
        source: 'http',
        before: auditSnapshot({ label: record.label, policy: record.policy }),
        after: auditSnapshot({ label, policy: validated.policy }),
      });
      return { status: 200, jsonBody: { ok: true, widened: !narrowing, ...publicAgentKey(updated) } };
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
        if (record.reminderEventId) {
          // The key's own session is the owner's delegated session, so it can
          // remove the owner's reminder whoever is revoking. Best effort.
          try {
            const keySession = await getSessionByStorageKey(record.sessionStorageKey);
            if (keySession) await deleteKeyExpiryReminder(keySession, record.reminderEventId);
          } catch (err: unknown) {
            context.warn(`[agent-keys] reminder for key ${keyId} not removed:`, err instanceof Error ? err.message : err);
          }
        }
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
      label?: unknown; services?: unknown; readOnlyServices?: unknown; allowedSites?: unknown; expiresInDays?: unknown; reminder?: unknown;
    } | null;
    if (!body || typeof body !== 'object') return { status: 400, jsonBody: { error: 'JSON body required' } };

    const label = typeof body.label === 'string' ? body.label.trim() : '';
    if (!label) return { status: 400, jsonBody: { error: 'label is required' } };
    if (label.length > AGENT_KEY_LABEL_MAX) return { status: 400, jsonBody: { error: `label must be at most ${AGENT_KEY_LABEL_MAX} characters` } };
    if (/[#\r\n]/.test(label)) return { status: 400, jsonBody: { error: 'label may not contain "#" or line breaks' } };

    // Lifetime is required: the key's session is capped by it, so there is no
    // such thing as a key without one. The tenant policy bounds it.
    const maxDays = Math.min(AGENT_KEY_MAX_LIFETIME_DAYS, policy.maxKeyLifetimeDays);
    const days = Number(body.expiresInDays);
    if (body.expiresInDays === undefined || body.expiresInDays === null || !Number.isInteger(days) || days < 1 || days > maxDays) {
      return { status: 400, jsonBody: { error: `expiresInDays must be an integer from 1 to ${maxDays} (the organization's maximum)` } };
    }
    const wantReminder = body.reminder !== false;

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
    const expiresAt = now + days * 86_400_000;
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
      // Its own anchor: the key's lifetime is the cap, enforced in
      // authenticateAgentKey from the key row's createdAt and expiresAt.
      sessionAbsoluteCreatedAt: now,
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
      after: auditSnapshot({ label, policy: validated.policy, expiresAt }),
    });

    // Calendar reminder, as the owner, on the owner's own calendar. Never fatal.
    let reminder: { eventId: string } | { skipped: string } | { error: string } = { skipped: 'not requested' };
    if (wantReminder) {
      try {
        const eventId = await createKeyExpiryReminder(browser, record);
        if (eventId) {
          await setAgentKeyReminder(record, eventId);
          reminder = { eventId };
        } else {
          reminder = { skipped: 'key is too short-lived for a reminder' };
        }
      } catch (err: unknown) {
        const message = err instanceof Error ? err.message : String(err);
        context.warn(`[agent-keys] reminder for key ${keyId} not created:`, message);
        logAccess({
          ...auditActor(auth.session),
          operation: 'policy.agent_key.reminder',
          resource: `key:${keyId}`,
          result: 'denied',
          reason: `calendar reminder not created: ${message}`,
          source: 'http',
        });
        reminder = { error: 'The key was created, but the calendar reminder could not be.' };
      }
    }
    return {
      status: 201,
      headers: { 'Cache-Control': 'no-store' },
      jsonBody: { key: token, ...publicAgentKey(record), reminder },
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
  methods: ['GET', 'POST', 'PATCH', 'DELETE'],
  authLevel: 'anonymous',
  route: 'api/manage/agent-keys',
  handler: withSecurity(manageAgentKeys),
});
