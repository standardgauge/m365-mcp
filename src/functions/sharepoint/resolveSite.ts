import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { checkGlobalAdmin } from '../../services/authMiddleware.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

/**
 * GET /api/sharepoint/resolve-site?url=<sharepoint-site-url>
 *
 * Resolves a SharePoint site by URL to its Graph site id + display name,
 * server-side, using the caller's session token. This replaces the admin UI's
 * former client-side Graph call in "Add by URL" — which depended on an MSAL
 * browser token that no longer exists now that the SPA authenticates purely
 * via the server session.
 *
 * Global-admin only: the sole consumer is the Allowed Sites admin panel, and
 * resolving arbitrary sites is an admin capability.
 */
async function resolveSiteHandler(
  request: HttpRequest,
  _context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const isAdmin = await checkGlobalAdmin(auth.userId, accessToken);
  if (!isAdmin) {
    return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
  }

  const rawUrl = request.query.get('url');
  if (!rawUrl) {
    return { status: 400, jsonBody: { error: 'Missing url query parameter' } };
  }

  let hostname: string;
  let serverRelPath: string;
  try {
    const parsed = new URL(rawUrl.startsWith('http') ? rawUrl : `https://${rawUrl}`);
    hostname = parsed.hostname;
    serverRelPath = parsed.pathname.replace(/\/+$/, '');
  } catch {
    return { status: 400, jsonBody: { error: 'Invalid site URL' } };
  }

  const graph = createGraphClient(accessToken);
  const apiPath = serverRelPath
    ? `/sites/${hostname}:${serverRelPath}`
    : `/sites/${hostname}`;

  try {
    const site = await graph.api(apiPath).select('id,displayName,webUrl').get();
    return {
      status: 200,
      jsonBody: { id: site.id, displayName: site.displayName, webUrl: site.webUrl },
    };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    return { status: 404, jsonBody: { error: `Could not resolve site: ${message}` } };
  }
}

app.http('resolveSite', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/resolve-site',
  handler: withSecurity(withPolicyEnforcement('sharepoint', resolveSiteHandler)),
});
