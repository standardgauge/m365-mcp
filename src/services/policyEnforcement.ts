/**
 * Shared policy-enforcement helpers for HTTP function routes.
 *
 * The MCP endpoint (mcpEndpoint.ts) already enforces enabledServices,
 * allowedSites, and deny-list policies inline. These helpers replicate
 * the same checks so that the parallel HTTP routes cannot be used to
 * bypass tenant policy.
 */

import { HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getTenantId, getTenantIdFromSession } from './tokenCache.js';
import type { UserSession } from './tokenCache.js';
import { getEnabledServices, getAllowedSites, getReadOnlyServices } from './serviceSettings.js';
import { isServiceDisabledForUser } from './userServiceOverrides.js';
import { isPathDenied, type DenySubject } from './denyList.js';
import { isMailIndexingDisabled } from './userMailConfig.js';
import { authenticateRequest, AuthResult } from './authMiddleware.js';
import { logAccess } from './auditLog.js';
import { auditClientAddress } from './clientAddress.js';
import { ValidationError } from './opaqueId.js';
import { SessionStoreUnavailableError, SESSION_STORE_RETRY_AFTER_S } from './sessionStoreError.js';

export type ServiceCategory = 'mail' | 'sharepoint' | 'onedrive' | 'calendar' | 'onenote' | 'contacts' | 'teams';

export interface PolicyViolation {
  status: number;
  error: string;
}

/**
 * Verify that the given service is enabled for the user's tenant.
 * Uses the session directly — no global userIndex lookup.
 */
export async function checkServiceEnabled(
  userId: string,
  service: ServiceCategory,
  session?: UserSession,
): Promise<PolicyViolation | null> {
  const tenantId = session ? getTenantIdFromSession(session) : await getTenantId(userId);

  // 1. Tenant-level check — if disabled at tenant level, user overrides are irrelevant
  const enabled = await getEnabledServices(tenantId);
  if (!enabled.includes(service)) {
    return { status: 403, error: `Service "${service}" is not enabled for this tenant` };
  }

  // 2. Per-user override — admin can disable a tenant-enabled service for specific users
  if (await isServiceDisabledForUser(tenantId, userId, service)) {
    return { status: 403, error: `Service "${service}" is disabled for your account` };
  }

  return null;
}

/**
 * Verify that the given siteId is in the tenant's allowedSites list.
 * Uses the session directly — no global userIndex lookup.
 */
export async function checkAllowedSite(
  userId: string,
  siteId: string,
  session?: UserSession,
): Promise<PolicyViolation | null> {
  const tenantId = session ? getTenantIdFromSession(session) : await getTenantId(userId);
  const allowed = await getAllowedSites(tenantId);
  if (allowed.length > 0) {
    const siteAllowed = allowed.some((s) => s.id === siteId);
    if (!siteAllowed) {
      return { status: 403, error: `Site "${siteId}" is not in the allowed sites list` };
    }
  }
  return null;
}

/**
 * Check whether the service is in read-only mode for the user's tenant
 *. When true, mutating operations (create/update/delete) are
 * refused while reads continue to work. Callers pass this only for mutating
 * routes; read routes never invoke it.
 * Uses the session directly — no global userIndex lookup.
 */
export async function checkReadOnly(
  userId: string,
  service: ServiceCategory,
  session?: UserSession,
): Promise<PolicyViolation | null> {
  const tenantId = session ? getTenantIdFromSession(session) : await getTenantId(userId);
  const readOnly = await getReadOnlyServices(tenantId);
  if (readOnly.includes(service)) {
    return { status: 403, error: `Service "${service}" is in read-only mode; write operations are disabled` };
  }
  return null;
}

/**
 * Check whether mail indexing is disabled for the user.
 * When true, ALL mail operations are denied before folder-level enforcement.
 * Uses the session directly — no global userIndex lookup.
 */
export async function checkMailIndexing(
  userId: string,
  session?: UserSession,
): Promise<PolicyViolation | null> {
  const tenantId = session ? getTenantIdFromSession(session) : await getTenantId(userId);
  if (await isMailIndexingDisabled(tenantId, userId)) {
    return { status: 403, error: 'Mail access is disabled for your account' };
  }
  return null;
}

/**
 * Check a single path against the deny list.
 * Uses the session directly — no global userIndex lookup.
 * `subject` names whose per-user lists apply when it is not the caller alone
 * (delegated mailbox access — see resolveDenySubject in mailboxOwner.ts).
 */
export async function checkDenyList(
  userId: string,
  service: ServiceCategory,
  path: string,
  session?: UserSession,
  subject: DenySubject = userId,
): Promise<PolicyViolation | null> {
  const tenantId = session ? getTenantIdFromSession(session) : await getTenantId(userId);
  if (await isPathDenied(tenantId, subject, service, path)) {
    return { status: 403, error: 'Access restricted by deny list' };
  }
  return null;
}

/* ── Higher-order wrapper ──────────────────────────────────────────────────── */

/**
 * Options for the policy enforcement wrapper.
 * - getSiteId: extract a SharePoint siteId from the request for allowedSites
 * - getDenyListPaths: extract zero or more paths to check against the deny list
 * - getResource: extract the specific resource identifier for audit logging
 *   (e.g. siteId/itemId, messageId, file path). When omitted, a default extractor
 *   builds a resource string from common path params and query params so allowed
 *   log entries always name the object that was accessed.
 */
export interface PolicyOptions {
  getSiteId?: (req: HttpRequest) => string | undefined;
  getDenyListPaths?: (req: HttpRequest) => string[];
  getResource?: (req: HttpRequest) => string | undefined;
  /**
   * Marks the route as a mutating (write) operation. When set, the wrapper
   * refuses the request with 403 if the service is in read-only mode for the
   * tenant. Read routes omit this flag.
   */
  mutating?: boolean;
}

/** Extract a resource string from common HTTP route params and query strings. */
function defaultResource(req: HttpRequest, options?: PolicyOptions): string | undefined {
  const parts: string[] = [];

  // Path params — ordered from coarser to finer grain
  const pathParamNames = [
    'siteId', 'driveId', 'itemId', 'messageId', 'folderId',
    'eventId', 'notebookId', 'sectionId', 'contactId', 'attachmentId',
  ];
  for (const name of pathParamNames) {
    const v = req.params?.[name];
    if (v) parts.push(v);
  }

  // Query params that identify the target object
  const queryParamNames = [
    'siteId', 'driveId', 'itemId', 'mailboxId', 'folderId',
    'listId', 'path', 'messageId', 'contactId',
  ];
  for (const name of queryParamNames) {
    const v = req.query.get(name);
    if (v && !parts.includes(v)) parts.push(v);
  }

  // getSiteId extractor (e.g. parsed from body for POST routes)
  if (options?.getSiteId) {
    const sid = options.getSiteId(req);
    if (sid && !parts.includes(sid)) parts.unshift(sid);
  }

  return parts.length > 0 ? parts.join('/') : undefined;
}

/**
 * The inner handler signature after auth + policy checks have passed.
 * Receives the authenticated user info and the original request/context.
 */
export type PolicyHandler = (
  req: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
) => Promise<HttpResponseInit>;

/**
 * Wraps an HTTP function handler with automatic authentication and policy
 * enforcement. The wrapper:
 *   1. Authenticates the request (401 if invalid)
 *   2. Checks enabledServices for the given service category (403 if blocked)
 *   3. For mail: checks disable_mail_indexing flag (403 + audit log if set)
 *   4. Optionally checks allowedSites when getSiteId returns a value
 *   5. Optionally checks the deny list for each path returned by getDenyListPaths
 *   6. Calls the inner handler only if all checks pass
 *
 * Error handling mirrors the existing per-route pattern (Re-authentication
 * required -> 401, session store unavailable -> 503, everything else -> 500).
 */
export function withPolicyEnforcement(
  service: ServiceCategory,
  handler: PolicyHandler,
  options?: PolicyOptions & { adminBypass?: boolean },
): (req: HttpRequest, context: InvocationContext) => Promise<HttpResponseInit> {
  return async (req: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    try {
      // 1. Authenticate
      const auth = await authenticateRequest(req);
      if (!auth) {
        return { status: 401, jsonBody: { error: 'Authentication required' } };
      }

      // 2. Admin bypass — ONLY for routes that explicitly opt in via
      //    adminBypass: true (admin config endpoints like manageAllowedSites,
      //    manageServices, denyList). Data routes never skip enforcement.
      if (options?.adminBypass && req.query.get('admin') === 'true') {
        const { checkGlobalAdmin } = await import('./authMiddleware.js');
        const isAdmin = await checkGlobalAdmin(auth.userId);
        if (isAdmin) {
          return handler(req, context, auth);
        }
      }

      // 3. Service enabled? (session-scoped — no global userIndex lookup)
      const serviceViolation = await checkServiceEnabled(auth.userId, service, auth.session);
      if (serviceViolation) {
        logAccess({
          tenantId: auth.session.tenantId,
          userId: auth.userId,
          userEmail: auth.session.email,
          deviceLabel: auth.session.deviceLabel,
          operation: `${service}.${req.method?.toLowerCase() ?? 'unknown'}`,
          result: 'denied',
          reason: serviceViolation.error,
          source: 'http',
          ip: auditClientAddress(req),
        });
        return { status: serviceViolation.status, jsonBody: { error: serviceViolation.error } };
      }

      // 3b. Mail indexing disabled? Short-circuits before folder deny-list.
      //     Runs only after service-enabled confirms mail is on for the tenant.
      if (service === 'mail') {
        const mailViolation = await checkMailIndexing(auth.userId, auth.session);
        if (mailViolation) {
          logAccess({
            tenantId: auth.session.tenantId,
            userId: auth.userId,
            userEmail: auth.session.email,
            deviceLabel: auth.session.deviceLabel,
            operation: `${service}.mail_access`,
            result: 'denied',
            reason: mailViolation.error,
            source: 'http',
            ip: auditClientAddress(req),
          });
          console.log(
            `[audit] mail-indexing-blocked user=${auth.userId} op=${req.method} url=${req.url} ts=${new Date().toISOString()}`,
          );
          return { status: mailViolation.status, jsonBody: { error: mailViolation.error } };
        }
      }

      // 3c. Read-only mode? Block mutating routes when the service is
      //     configured read-only for the tenant.
      if (options?.mutating) {
        const readOnlyViolation = await checkReadOnly(auth.userId, service, auth.session);
        if (readOnlyViolation) {
          logAccess({
            tenantId: auth.session.tenantId,
            userId: auth.userId,
            userEmail: auth.session.email,
            deviceLabel: auth.session.deviceLabel,
            operation: `${service}.${req.method?.toLowerCase() ?? 'unknown'}`,
            result: 'denied',
            reason: readOnlyViolation.error,
            source: 'http',
            ip: auditClientAddress(req),
          });
          return { status: readOnlyViolation.status, jsonBody: { error: readOnlyViolation.error } };
        }
      }

      // 4. Allowed site? (SharePoint only, when a siteId is available)
      if (options?.getSiteId) {
        const siteId = options.getSiteId(req);
        if (siteId) {
          const siteViolation = await checkAllowedSite(auth.userId, siteId, auth.session);
          if (siteViolation) {
            logAccess({
              tenantId: auth.session.tenantId,
              userId: auth.userId,
              userEmail: auth.session.email,
              deviceLabel: auth.session.deviceLabel,
              operation: `${service}.${req.method?.toLowerCase() ?? 'unknown'}`,
              resource: siteId,
              result: 'denied',
              reason: siteViolation.error,
              source: 'http',
              ip: auditClientAddress(req),
            });
            return { status: siteViolation.status, jsonBody: { error: siteViolation.error } };
          }
        }
      }

      // 5. Deny list?
      if (options?.getDenyListPaths) {
        const paths = options.getDenyListPaths(req);
        for (const p of paths) {
          if (p) {
            const denyViolation = await checkDenyList(auth.userId, service, p, auth.session);
            if (denyViolation) {
              logAccess({
                tenantId: auth.session.tenantId,
                userId: auth.userId,
                userEmail: auth.session.email,
                deviceLabel: auth.session.deviceLabel,
                operation: `${service}.${req.method?.toLowerCase() ?? 'unknown'}`,
                resource: p,
                result: 'denied',
                reason: denyViolation.error,
                source: 'http',
                ip: auditClientAddress(req),
              });
              return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
            }
          }
        }
      }

      // 6. All clear — run the handler
      // Log allowed access (fire-and-forget) with the most specific resource available
      const resource = options?.getResource
        ? options.getResource(req)
        : defaultResource(req, options);
      logAccess({
        tenantId: auth.session.tenantId,
        userId: auth.userId,
        userEmail: auth.session.email,
        deviceLabel: auth.session.deviceLabel,
        operation: `${service}.${req.method?.toLowerCase() ?? 'unknown'}`,
        resource,
        result: 'allowed',
        source: 'http',
        ip: auditClientAddress(req),
      });
      return await handler(req, context, auth);
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : 'Unknown error';
      context.error(`[${service}] handler error:`, message);
      if (err instanceof ValidationError) {
        return { status: 400, jsonBody: { error: message } };
      }
      // Storage could not authenticate the request: a server error, never a 401.
      if (err instanceof SessionStoreUnavailableError) {
        return { status: 503, headers: { 'Retry-After': String(SESSION_STORE_RETRY_AFTER_S) }, jsonBody: { error: 'Session store unavailable, retry shortly' } };
      }
      if (message.includes('Re-authentication required')) {
        return { status: 401, jsonBody: { error: message } };
      }
      return { status: 500, jsonBody: { error: 'Internal server error' } };
    }
  };
}
