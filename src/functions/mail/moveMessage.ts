import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { resolveMailFolderName } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface MoveRequest {
  destinationFolderId: string;
}

async function moveMessageHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const messageId = request.params['messageId'];
  if (!messageId) {
    return { status: 400, jsonBody: { error: 'messageId path param is required' } };
  }
  assertOpaqueId(messageId, 'messageId');

  const body = (await request.json()) as MoveRequest;
  if (!body.destinationFolderId) {
    return { status: 400, jsonBody: { error: 'destinationFolderId is required' } };
  }
  assertOpaqueId(body.destinationFolderId, 'destinationFolderId');

  const mailboxId = request.query.get('mailboxId');
  if (mailboxId && mailboxId !== 'me') assertOpaqueId(mailboxId, 'mailboxId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  const baseMsgPath = mailboxId && mailboxId !== 'me'
    ? `/users/${mailboxId}/messages/${messageId}`
    : `/me/messages/${messageId}`;

  const mailboxBase = mailboxId && mailboxId !== 'me' ? `/users/${mailboxId}` : '/me';
  const tenantId = await getTenantId(userId);

  // Check deny list for source folder (resolve display name — deny list stores names, not IDs)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const msgMeta: any = await graph.api(baseMsgPath).select('parentFolderId').get();
  if (msgMeta.parentFolderId) {
    const srcFolderName = await resolveMailFolderName(graph, msgMeta.parentFolderId, mailboxBase);
    if (srcFolderName && await isPathDenied(tenantId, userId, 'mail', srcFolderName)) {
      return { status: 403, jsonBody: { error: 'Access restricted by deny list' } };
    }
  }

  // Check deny list for destination folder (resolve display name)
  const destFolderName = await resolveMailFolderName(graph, body.destinationFolderId, mailboxBase);
  if (destFolderName && await isPathDenied(tenantId, userId, 'mail', destFolderName)) {
    return { status: 403, jsonBody: { error: 'Destination folder restricted by deny list' } };
  }

  const apiPath = `${baseMsgPath}/move`;

  const result = await graph.api(apiPath).post({
    destinationId: body.destinationFolderId,
  });

  return {
    status: 200,
    jsonBody: {
      id: result.id,
      subject: result.subject,
      folderId: result.parentFolderId,
      status: 'moved',
    },
  };
}

app.http('moveMessage', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/mail/messages/{messageId}/move',
  handler: withSecurity(withPolicyEnforcement('mail', moveMessageHandler)),
});
