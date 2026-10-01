import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface CreateContactFolderRequest {
  displayName: string;
  parentFolderId?: string;
}

async function createContactFolderHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as CreateContactFolderRequest;
  if (!body.displayName) {
    return { status: 400, jsonBody: { error: 'displayName is required' } };
  }
  if (body.parentFolderId && body.parentFolderId !== 'contacts-root') {
    assertOpaqueId(body.parentFolderId, 'parentFolderId');
  }

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Guard the parent: creating a child under a deny-listed folder is denied.
  if (body.parentFolderId && body.parentFolderId !== 'contacts-root') {
    const tenantId = await getTenantId(userId);
    if (await isPathDenied(tenantId, userId, 'contacts', body.parentFolderId)) {
      return { status: 403, jsonBody: { error: 'Access restricted by deny list' } };
    }
  }

  const apiPath = body.parentFolderId && body.parentFolderId !== 'contacts-root'
    ? `/me/contactFolders/${body.parentFolderId}/childFolders`
    : '/me/contactFolders';

  const result = await graph.api(apiPath).post({ displayName: body.displayName });

  return {
    status: 201,
    jsonBody: {
      id: result.id,
      displayName: result.displayName,
      parentFolderId: result.parentFolderId ?? null,
      status: 'created',
    },
  };
}

app.http('createContactFolder', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/contacts/folders',
  handler: withSecurity(withPolicyEnforcement('contacts', createContactFolderHandler, { mutating: true })),
});
