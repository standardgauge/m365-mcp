import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedPaths } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface FolderItem {
  id: string;
  name: string;
  webUrl: string;
  path: string;
  childCount: number;
  createdDateTime: string;
  lastModifiedDateTime: string;
}

async function listFoldersHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const siteId = request.query.get('siteId');
  const driveId = request.query.get('driveId');
  const parentId = request.query.get('parentId');

  if (!siteId) {
    return { status: 400, jsonBody: { error: 'siteId query param is required' } };
  }
  assertOpaqueId(siteId, 'siteId');
  if (driveId) assertOpaqueId(driveId, 'driveId');
  if (parentId) assertOpaqueId(parentId, 'parentId');

  const graph = createGraphClient(accessToken);

  // Build the Graph API path depending on which params are provided
  let apiPath: string;
  if (driveId && parentId) {
    apiPath = `/sites/${siteId}/drives/${driveId}/items/${parentId}/children`;
  } else if (driveId) {
    apiPath = `/sites/${siteId}/drives/${driveId}/root/children`;
  } else if (parentId) {
    apiPath = `/sites/${siteId}/drive/items/${parentId}/children`;
  } else {
    apiPath = `/sites/${siteId}/drive/root/children`;
  }

  const result = await graph
    .api(apiPath)
    .select('id,name,webUrl,folder,parentReference,createdDateTime,lastModifiedDateTime')
    .filter('folder ne null')
    .get();

  const folders: FolderItem[] = (result.value ?? [])
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    .filter((item: any) => item.folder !== undefined)
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    .map((item: any) => {
      const parentPath: string = item.parentReference?.path ?? '';
      return {
        id: item.id,
        name: item.name,
        webUrl: item.webUrl,
        path: parentPath ? `${parentPath}/${item.name}` : `/${item.name}`,
        childCount: item.folder?.childCount ?? 0,
        createdDateTime: item.createdDateTime,
        lastModifiedDateTime: item.lastModifiedDateTime,
      };
    });

  const allowed = await filterDeniedPaths(await getTenantId(userId), userId, 'sharepoint', folders);

  return { status: 200, jsonBody: { folders: allowed, count: allowed.length } };
}

app.http('listFolders', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/sharepoint/folders',
  handler: withSecurity(withPolicyEnforcement('sharepoint', listFoldersHandler, {
    getSiteId: (req) => req.query.get('siteId') ?? undefined,
  })),
});
