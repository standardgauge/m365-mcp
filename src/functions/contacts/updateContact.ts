import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { resolveContactParentFolder } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { buildContactBody, type ContactInput } from '../../services/contactFields.js';

type UpdateContactRequest = ContactInput;

async function updateContactHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const contactId = request.params['contactId'];
  if (!contactId) {
    return { status: 400, jsonBody: { error: 'contactId path param is required' } };
  }
  assertOpaqueId(contactId, 'contactId');

  const body = (await request.json()) as UpdateContactRequest;

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Resolve parent folder and check deny list for direct contactId operations
  const parentFolderId = await resolveContactParentFolder(graph, userId, contactId);
  if (parentFolderId) {
    const tenantId = await getTenantId(userId);
    if (await isPathDenied(tenantId, userId, 'contacts', parentFolderId)) {
      return { status: 403, jsonBody: { error: 'Access restricted by deny list' } };
    }
  }

  const patch = buildContactBody(body);

  const result = await graph.api(`/me/contacts/${contactId}`).patch(patch);

  return {
    status: 200,
    jsonBody: {
      id: result.id,
      displayName: result.displayName,
      status: 'updated',
    },
  };
}

app.http('updateContact', {
  methods: ['PATCH'],
  authLevel: 'anonymous',
  route: 'api/contacts/{contactId}',
  handler: withSecurity(withPolicyEnforcement('contacts', updateContactHandler)),
});
