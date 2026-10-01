import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedPaths } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface OneDriveItem {
  id: string;
  name: string;
  path: string;
  type: 'folder' | 'file';
  childCount?: number;
  size?: number;
  mimeType?: string;
  lastModifiedDateTime?: string;
}

async function listOneDriveItemsHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const parentId = request.query.get('parentId');
  if (parentId) assertOpaqueId(parentId, 'parentId');
  const foldersOnly = request.query.get('foldersOnly') === 'true';
  const graph = createGraphClient(accessToken);

  const apiPath = parentId
    ? `/me/drive/items/${parentId}/children`
    : '/me/drive/root/children';

  let query = graph
    .api(apiPath)
    .select('id,name,folder,file,size,parentReference,lastModifiedDateTime')
    .top(200);

  if (foldersOnly) {
    query = query.filter('folder ne null');
  }

  const result = await query.get();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const items: OneDriveItem[] = (result.value ?? []).map((item: any) => {
    const parentPath: string = item.parentReference?.path ?? '';
    const isFolder = item.folder !== undefined;
    return {
      id: item.id,
      name: item.name,
      path: parentPath ? `${parentPath}/${item.name}` : `/${item.name}`,
      type: isFolder ? 'folder' : 'file',
      ...(isFolder ? { childCount: item.folder?.childCount ?? 0 } : {}),
      ...(!isFolder ? {
        size: item.size,
        mimeType: item.file?.mimeType ?? null,
      } : {}),
      lastModifiedDateTime: item.lastModifiedDateTime,
    };
  });

  const allowed = await filterDeniedPaths(await getTenantId(userId), userId, 'onedrive', items);

  const folders = allowed.filter(i => i.type === 'folder');
  const files = allowed.filter(i => i.type === 'file');

  return {
    status: 200,
    jsonBody: {
      items: allowed,
      count: allowed.length,
      folders: folders.length,
      files: files.length,
    },
  };
}

app.http('listOneDriveItems', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/onedrive/folders',
  handler: withSecurity(withPolicyEnforcement('onedrive', listOneDriveItemsHandler)),
});
