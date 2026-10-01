import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface CreateFolderRequest {
  name: string;
  parentId?: string;
}

async function createFolderHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as CreateFolderRequest;
  if (!body.name) {
    return { status: 400, jsonBody: { error: 'name is required' } };
  }
  if (body.parentId) assertOpaqueId(body.parentId, 'parentId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Resolve the parent path to check deny list before creating
  if (body.parentId) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const parentMeta: any = await graph.api(`/me/drive/items/${body.parentId}`).select('name,parentReference').get();
    const parentPath = `${parentMeta.parentReference?.path ?? ''}/${parentMeta.name}`;
    const denyViolation = await checkDenyList(userId, 'onedrive', `${parentPath}/${body.name}`);
    if (denyViolation) {
      return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
    }
  }

  const apiPath = body.parentId
    ? `/me/drive/items/${body.parentId}/children`
    : '/me/drive/root/children';

  const result = await graph.api(apiPath).post({
    name: body.name,
    folder: {},
    '@microsoft.graph.conflictBehavior': 'fail',
  });

  return {
    status: 201,
    jsonBody: {
      id: result.id,
      name: result.name,
      webUrl: result.webUrl,
      status: 'created',
    },
  };
}

app.http('createFolder', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/onedrive/folders',
  handler: withSecurity(withPolicyEnforcement('onedrive', createFolderHandler)),
});
