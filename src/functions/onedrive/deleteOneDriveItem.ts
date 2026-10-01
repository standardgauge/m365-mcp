import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function deleteOneDriveItemHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const itemId = request.params['itemId'];
  if (!itemId) {
    return { status: 400, jsonBody: { error: 'itemId path param is required' } };
  }
  assertOpaqueId(itemId, 'itemId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Check deny list via item path
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const meta: any = await graph.api(`/me/drive/items/${itemId}`).select('name,parentReference').get();
  const itemPath = `${meta.parentReference?.path ?? ''}/${meta.name}`;
  const denyViolation = await checkDenyList(userId, 'onedrive', itemPath);
  if (denyViolation) {
    return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
  }

  await graph.api(`/me/drive/items/${itemId}`).delete();

  return { status: 200, jsonBody: { status: 'deleted', itemId } };
}

app.http('deleteOneDriveItem', {
  methods: ['DELETE'],
  authLevel: 'anonymous',
  route: 'api/onedrive/files/{itemId}',
  handler: withSecurity(withPolicyEnforcement('onedrive', deleteOneDriveItemHandler)),
});
