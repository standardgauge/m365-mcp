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
import { readMailAttachment } from '../../services/mailAttachmentContent.js';

async function getAttachmentHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const messageId = request.params['messageId'];
  const attachmentId = request.query.get('attachmentId');
  const mailboxId = request.query.get('mailboxId');

  if (!messageId) {
    return { status: 400, jsonBody: { error: 'messageId path param is required' } };
  }
  assertOpaqueId(messageId, 'messageId');
  if (attachmentId) assertOpaqueId(attachmentId, 'attachmentId');
  if (mailboxId && mailboxId !== 'me') assertOpaqueId(mailboxId, 'mailboxId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, mailboxId);

  const basePath = mailboxId && mailboxId !== 'me'
    ? `/users/${mailboxId}/messages/${messageId}`
    : `/me/messages/${messageId}`;

  // Check deny list via parent folder display name (deny list stores names, not IDs)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const msgMeta: any = await graph.api(basePath).select('parentFolderId').get();
  if (msgMeta.parentFolderId) {
    const mailboxBase = mailboxId && mailboxId !== 'me' ? `/users/${mailboxId}` : '/me';
    const folderName = await resolveMailFolderName(graph, msgMeta.parentFolderId, mailboxBase);
    if (folderName) {
      const tenantId = await getTenantId(userId);
      if (await isPathDenied(tenantId, denySubject, 'mail', folderName)) {
        return { status: 403, jsonBody: { error: 'Access restricted by deny list' } };
      }
    }
  }

  if (attachmentId) {
    // Fetch a specific attachment (file bytes, or raw MIME for an attached email)
    const attachment = await readMailAttachment(graph, `${basePath}/attachments/${attachmentId}`);
    return { status: 200, jsonBody: attachment };
  } else {
    // List all attachments on the message
    const result = await graph
      .api(`${basePath}/attachments`)
      .select('id,name,contentType,size')
      .get();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const attachments = (result.value ?? []).map((a: any) => ({
      id: a.id,
      name: a.name,
      contentType: a.contentType,
      size: a.size,
    }));

    return { status: 200, jsonBody: { attachments, count: attachments.length } };
  }
}

app.http('getAttachment', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/mail/messages/{messageId}/attachments',
  handler: withSecurity(withPolicyEnforcement('mail', getAttachmentHandler)),
});
