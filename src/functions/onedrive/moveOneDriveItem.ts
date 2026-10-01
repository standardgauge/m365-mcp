import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface MoveRequest {
  itemId: string;
  destinationFolderId: string;
  newName?: string;
}

async function moveOneDriveItemHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as MoveRequest;
  if (!body.itemId || !body.destinationFolderId) {
    return { status: 400, jsonBody: { error: 'itemId and destinationFolderId are required' } };
  }
  assertOpaqueId(body.itemId, 'itemId');
  assertOpaqueId(body.destinationFolderId, 'destinationFolderId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Check deny list for source item
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const srcMeta: any = await graph.api(`/me/drive/items/${body.itemId}`).select('name,parentReference').get();
  const srcPath = `${srcMeta.parentReference?.path ?? ''}/${srcMeta.name}`;
  const srcDeny = await checkDenyList(userId, 'onedrive', srcPath);
  if (srcDeny) {
    return { status: srcDeny.status, jsonBody: { error: srcDeny.error } };
  }

  // Check deny list for destination folder
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const destMeta: any = await graph.api(`/me/drive/items/${body.destinationFolderId}`).select('name,parentReference').get();
  const destPath = `${destMeta.parentReference?.path ?? ''}/${destMeta.name}`;
  const destDeny = await checkDenyList(userId, 'onedrive', destPath);
  if (destDeny) {
    return { status: destDeny.status, jsonBody: { error: 'Destination folder restricted by deny list' } };
  }

  const patchBody: Record<string, unknown> = {
    parentReference: { id: body.destinationFolderId },
  };
  if (body.newName) {
    patchBody.name = body.newName;
  }

  const result = await graph
    .api(`/me/drive/items/${body.itemId}`)
    .patch(patchBody);

  return {
    status: 200,
    jsonBody: {
      id: result.id,
      name: result.name,
      webUrl: result.webUrl,
      parentFolderId: result.parentReference?.id,
      status: 'moved',
    },
  };
}

app.http('moveOneDriveItem', {
  methods: ['PATCH'],
  authLevel: 'anonymous',
  route: 'api/onedrive/move',
  handler: withSecurity(withPolicyEnforcement('onedrive', moveOneDriveItemHandler)),
});
