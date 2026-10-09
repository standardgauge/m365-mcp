import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { resolveDenySubject } from '../../services/mailboxOwner.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface CreateMailFolderRequest {
  displayName: string;
  parentFolderId?: string;
  mailboxId?: string;
}

async function createMailFolderHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as CreateMailFolderRequest;
  if (!body.displayName) {
    return { status: 400, jsonBody: { error: 'displayName is required' } };
  }

  if (body.mailboxId && body.mailboxId !== 'me') assertOpaqueId(body.mailboxId, 'mailboxId');
  if (body.parentFolderId) assertOpaqueId(body.parentFolderId, 'parentFolderId');

  // Mail deny-list entries match on display name (see listFoldersMail / moveMessage).
  // Block creating a folder whose name is already denied, so it can't be used to
  // route mail past the deny list.
  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, body.mailboxId);
  const denyViolation = await checkDenyList(userId, 'mail', body.displayName, undefined, denySubject);
  if (denyViolation) {
    return { status: denyViolation.status, jsonBody: { error: denyViolation.error } };
  }

  const mailboxBase = body.mailboxId && body.mailboxId !== 'me'
    ? `/users/${body.mailboxId}`
    : '/me';

  const apiPath = body.parentFolderId
    ? `${mailboxBase}/mailFolders/${body.parentFolderId}/childFolders`
    : `${mailboxBase}/mailFolders`;

  const result = await graph.api(apiPath).post({
    displayName: body.displayName,
  });

  return {
    status: 201,
    jsonBody: {
      id: result.id,
      displayName: result.displayName,
      parentFolderId: result.parentFolderId,
      status: 'created',
    },
  };
}

app.http('createMailFolder', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/mail/folders',
  handler: withSecurity(withPolicyEnforcement('mail', createMailFolderHandler)),
});
