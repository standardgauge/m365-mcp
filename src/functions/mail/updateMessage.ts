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
import { getMailboxAddresses, validateFromAddress, toGraphFrom, FromAddressError } from '../../services/mailFrom.js';

interface UpdateMessageRequest {
  subject?: string;
  body?: string;
  bodyType?: 'text' | 'html';
  to?: string[];
  cc?: string[];
  bcc?: string[];
  /** Replacement From — one of the mailbox's proxy addresses. */
  from?: string;
}

async function updateMessageHandler(
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

  const reqBody = (await request.json()) as UpdateMessageRequest;
  // Graph allows a much narrower set of patches (categories, isRead, flag, etc.)
  // on sent mail; the fields below are draft-only and we surface a friendly
  // error if they're sent against a non-draft.
  const requestedDraftFields: string[] = [];
  if (reqBody.subject !== undefined) requestedDraftFields.push('subject');
  if (reqBody.body !== undefined) requestedDraftFields.push('body');
  if (reqBody.to !== undefined) requestedDraftFields.push('to');
  if (reqBody.cc !== undefined) requestedDraftFields.push('cc');
  if (reqBody.bcc !== undefined) requestedDraftFields.push('bcc');
  if (reqBody.from !== undefined) requestedDraftFields.push('from');
  if (requestedDraftFields.length === 0) {
    return { status: 400, jsonBody: { error: 'at least one of subject, body, to, cc, bcc, from must be provided' } };
  }

  const mailboxId = request.query.get('mailboxId');
  if (mailboxId && mailboxId !== 'me') assertOpaqueId(mailboxId, 'mailboxId');
  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, mailboxId);

  const apiPath = mailboxId && mailboxId !== 'me'
    ? `/users/${mailboxId}/messages/${messageId}`
    : `/me/messages/${messageId}`;
  const mailboxBase = mailboxId && mailboxId !== 'me' ? `/users/${mailboxId}` : '/me';

  // Pull isDraft + parentFolderId in one round-trip so we can both gate on
  // draft-status and enforce the deny list on the message's folder.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const msgMeta: any = await graph.api(apiPath).select('isDraft,parentFolderId').get();

  if (msgMeta.parentFolderId) {
    const folderName = await resolveMailFolderName(graph, msgMeta.parentFolderId, mailboxBase);
    if (folderName) {
      const tenantId = await getTenantId(userId);
      if (await isPathDenied(tenantId, denySubject, 'mail', folderName)) {
        return { status: 403, jsonBody: { error: 'Access restricted by deny list' } };
      }
    }
  }

  if (msgMeta.isDraft === false) {
    return {
      status: 400,
      jsonBody: {
        error:
          'This message is not a draft. Graph only allows updates to categories/flag/isRead on sent messages — to rewrite the body or recipients, delete the original (delete_message) and create a new draft (create_draft).',
        fields: requestedDraftFields,
      },
    };
  }

  const patch: Record<string, unknown> = {};
  if (reqBody.subject !== undefined) patch.subject = reqBody.subject;
  if (reqBody.body !== undefined) {
    patch.body = {
      contentType: reqBody.bodyType === 'html' ? 'HTML' : 'Text',
      content: reqBody.body,
    };
  }
  if (reqBody.to !== undefined) {
    patch.toRecipients = reqBody.to.map((addr) => ({ emailAddress: { address: addr } }));
  }
  if (reqBody.cc !== undefined) {
    patch.ccRecipients = reqBody.cc.map((addr) => ({ emailAddress: { address: addr } }));
  }
  if (reqBody.bcc !== undefined) {
    patch.bccRecipients = reqBody.bcc.map((addr) => ({ emailAddress: { address: addr } }));
  }
  if (reqBody.from !== undefined) {
    try {
      patch.from = toGraphFrom(validateFromAddress(reqBody.from, await getMailboxAddresses(graph, mailboxBase)));
    } catch (err) {
      if (err instanceof FromAddressError) return { status: 400, jsonBody: { error: err.message } };
      throw err;
    }
  }

  const result = await graph.api(apiPath).patch(patch);

  return {
    status: 200,
    jsonBody: {
      id: result.id,
      subject: result.subject,
      webLink: result.webLink,
      lastModifiedDateTime: result.lastModifiedDateTime,
      status: 'updated',
    },
  };
}

app.http('updateMessage', {
  methods: ['PATCH'],
  authLevel: 'anonymous',
  route: 'api/mail/messages/{messageId}',
  handler: withSecurity(withPolicyEnforcement('mail', updateMessageHandler)),
});
