import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { resolveDefaultContactFolder } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { contactsApiPath, runContactBatch, type ContactInput } from '../../services/contactFields.js';

interface CreateContactsBatchRequest {
  contacts: ContactInput[];
  folderId?: string;
}

async function createContactsBatchHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as CreateContactsBatchRequest;
  if (!Array.isArray(body.contacts) || body.contacts.length === 0) {
    return { status: 400, jsonBody: { error: 'contacts must be a non-empty array' } };
  }
  if (body.folderId && body.folderId !== 'contacts-root') {
    assertOpaqueId(body.folderId, 'folderId');
  }

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Deny-list check once for the shared target folder.
  const tenantId = await getTenantId(userId);
  if (body.folderId && body.folderId !== 'contacts-root') {
    if (await isPathDenied(tenantId, userId, 'contacts', body.folderId)) {
      return { status: 403, jsonBody: { error: 'Access restricted by deny list' } };
    }
  } else {
    const defaultFolder = await resolveDefaultContactFolder(graph, userId);
    if (defaultFolder && await isPathDenied(tenantId, userId, 'contacts', defaultFolder)) {
      return { status: 403, jsonBody: { error: 'Access to the default contacts folder is restricted by the deny list' } };
    }
  }

  const results = await runContactBatch(graph, body.contacts, contactsApiPath(body.folderId));
  const created = results.filter((r) => r.status === 'created').length;

  return {
    status: 200,
    jsonBody: { created, failed: results.length - created, results },
  };
}

app.http('createContactsBatch', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/contacts/batch',
  handler: withSecurity(withPolicyEnforcement('contacts', createContactsBatchHandler, { mutating: true })),
});
