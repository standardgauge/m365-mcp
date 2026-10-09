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

async function readMessageHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const messageId = request.params['messageId'];
  const mailboxId = request.query.get('mailboxId') ?? 'me';

  if (!messageId) {
    return { status: 400, jsonBody: { error: 'messageId path param is required' } };
  }
  assertOpaqueId(messageId, 'messageId');
  if (mailboxId !== 'me') assertOpaqueId(mailboxId, 'mailboxId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, mailboxId);

  const base = mailboxId === 'me' ? '/me' : `/users/${mailboxId}`;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const message: any = await graph
    .api(`${base}/messages/${messageId}`)
    .select(
      'id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,parentFolderId,isRead,importance,categories'
    )
    .get();

  // Deny-list check: resolve folder display name (deny list stores names, not IDs)
  if (message.parentFolderId) {
    const folderName = await resolveMailFolderName(graph, message.parentFolderId, base);
    if (folderName && await isPathDenied(await getTenantId(userId), denySubject, 'mail', folderName)) {
      return {
        status: 403,
        jsonBody: { error: 'Access to this mail folder is restricted by the deny list' },
      };
    }
  }

  return {
    status: 200,
    jsonBody: {
      id: message.id,
      subject: message.subject ?? null,
      from: message.from?.emailAddress ?? null,
      to: (message.toRecipients ?? []).map(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (r: any) => r.emailAddress
      ),
      cc: (message.ccRecipients ?? []).map(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (r: any) => r.emailAddress
      ),
      receivedDateTime: message.receivedDateTime,
      body: message.body?.content ?? null,
      bodyType: message.body?.contentType ?? null,
      hasAttachments: message.hasAttachments ?? false,
      isRead: message.isRead ?? false,
      importance: message.importance ?? 'normal',
      categories: message.categories ?? [],
      folderId: message.parentFolderId ?? null,
    },
  };
}

app.http('readMessage', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/mail/messages/{messageId}',
  handler: withSecurity(withPolicyEnforcement('mail', readMessageHandler)),
});
