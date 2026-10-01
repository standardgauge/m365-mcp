import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { resolveDefaultContactFolder } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { buildContactBody, contactsApiPath, type ContactInput } from '../../services/contactFields.js';

interface ContactRequest extends ContactInput {
  givenName: string;
  folderId?: string;
}

async function createContactHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as ContactRequest;
  if (!body.givenName) {
    return { status: 400, jsonBody: { error: 'givenName is required' } };
  }
  if (body.folderId) assertOpaqueId(body.folderId, 'folderId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Deny-list check on folderId (body-level, not available to wrapper)
  if (body.folderId) {
    const denyViolation = await checkDenyList(userId, 'contacts', body.folderId);
    if (denyViolation) {
      return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
    }
  } else {
    // No folderId — contact goes to default contacts folder; check that
    const defaultFolder = await resolveDefaultContactFolder(graph, userId);
    if (defaultFolder) {
      const tenantId = await getTenantId(userId);
      if (await isPathDenied(tenantId, userId, 'contacts', defaultFolder)) {
        return { status: 403, jsonBody: { error: 'Access to the default contacts folder is restricted by the deny list' } };
      }
    }
  }

  const contact = buildContactBody(body);

  const result = await graph.api(contactsApiPath(body.folderId)).post(contact);

  return {
    status: 201,
    jsonBody: {
      id: result.id,
      displayName: result.displayName,
      status: 'created',
    },
  };
}

app.http('createContact', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/contacts',
  handler: withSecurity(withPolicyEnforcement('contacts', createContactHandler)),
});
