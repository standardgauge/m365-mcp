import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterAndDisambiguateSites } from '../../services/sharepointFilter.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { getAllowedSites } from '../../services/serviceSettings.js';
import { checkGlobalAdmin } from '../../services/authMiddleware.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface SharePointSite {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
  description: string | null;
  createdDateTime: string;
  lastModifiedDateTime: string;
}

async function listSitesHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const graph = createGraphClient(accessToken);

  const result = await graph
    .api('/sites?search=*')
    .select('id,name,displayName,webUrl,description,createdDateTime,lastModifiedDateTime')
    .get();

  const rawSites: SharePointSite[] = (result.value ?? []).map((s: SharePointSite) => ({
    id: s.id,
    name: s.name,
    displayName: s.displayName,
    webUrl: s.webUrl,
    description: s.description ?? null,
    createdDateTime: s.createdDateTime,
    lastModifiedDateTime: s.lastModifiedDateTime,
  }));

  //: drop system sites (contentTypeHub, appcatalog, etc.) and
  // disambiguate displayName collisions by appending URL path so the
  // admin UI dropdown doesn't pick the wrong site silently.
  let sites = filterAndDisambiguateSites(rawSites);

  // Admin UI needs to see ALL sites the admin could possibly add to the allowlist.
  // For Global Admins making admin=true requests, return Graph results UNION
  // existing allowedSites entries — covers the case where a site was added via
  // "Add by URL" but isn't in SharePoint's search index (so /sites?search=*
  // doesn't return it).
  //
  // For non-admin requests, enforce the allowedSites filter as usual.
  const isAdminRequest = request.query.get('admin') === 'true';
  const isAdmin = isAdminRequest ? await checkGlobalAdmin(userId, accessToken) : false;
  const tenantId = await getTenantId(userId);
  const allowed = await getAllowedSites(tenantId);

  if (isAdmin) {
    // Union: include allowedSites entries that Graph search didn't return,
    // so the admin can see them in the UI even if they're not search-indexed.
    const seenIds = new Set(sites.map((s) => s.id));
    for (const allowedSite of allowed) {
      if (!seenIds.has(allowedSite.id)) {
        sites.push({
          id: allowedSite.id,
          name: allowedSite.name,
          displayName: allowedSite.name,
          webUrl: '',
          description: null,
          createdDateTime: '',
          lastModifiedDateTime: '',
        });
      }
    }
  } else if (allowed.length > 0) {
    const allowedIds = new Set(allowed.map((s) => s.id));
    sites = sites.filter((s) => allowedIds.has(s.id));
  }

  return { status: 200, jsonBody: { sites, count: sites.length } };
}

app.http('listSites', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/sites',
  handler: withSecurity(withPolicyEnforcement('sharepoint', listSitesHandler)),
});
