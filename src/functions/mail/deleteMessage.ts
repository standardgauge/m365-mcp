import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { resolveDenySubject } from '../../services/mailboxOwner.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { resolveMailFolderName } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function deleteMessageHandler(
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

  const mailboxId = request.query.get('mailboxId');
  if (mailboxId && mailboxId !== 'me') assertOpaqueId(mailboxId, 'mailboxId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, mailboxId);

  const apiPath = mailboxId && mailboxId !== 'me'
    ? `/users/${mailboxId}/messages/${messageId}`
    : `/me/messages/${messageId}`;

  // Check deny list via parent folder display name (deny list stores names, not IDs)
  const mailboxBase = mailboxId && mailboxId !== 'me' ? `/users/${mailboxId}` : '/me';
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const msgMeta: any = await graph.api(apiPath).select('parentFolderId').get();
  if (msgMeta.parentFolderId) {
    const folderName = await resolveMailFolderName(graph, msgMeta.parentFolderId, mailboxBase);
    if (folderName) {
      const tenantId = await getTenantId(userId);
      if (await isPathDenied(tenantId, denySubject, 'mail', folderName)) {
        return { status: 403, jsonBody: { error: 'Access restricted by deny list' } };
      }
    }
  }

  await graph.api(apiPath).delete();

  return { status: 200, jsonBody: { status: 'deleted', messageId } };
}

app.http('deleteMessage', {
  methods: ['DELETE'],
  authLevel: 'anonymous',
  route: 'api/mail/messages/{messageId}',
  handler: withSecurity(withPolicyEnforcement('mail', deleteMessageHandler)),
});
