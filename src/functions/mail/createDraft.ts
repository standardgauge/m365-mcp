import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantIdFromSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import {
  parseAttachments,
  createDraftWithAttachments,
} from '../../services/mailAttachments.js';
import { resolveAttachments, AttachmentSourceError } from '../../services/driveAttachments.js';
import { getMailboxAddresses, validateFromAddress, toGraphFrom, FromAddressError } from '../../services/mailFrom.js';

interface DraftRequest {
  subject: string;
  body: string;
  bodyType?: 'text' | 'html';
  to: string[];
  cc?: string[];
  bcc?: string[];
  /** Inline `{ name, contentType, content }` items, or drive references. */
  attachments?: unknown[];
  /** Send as this proxy address of the mailbox. */
  from?: string;
}

async function createDraftHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as DraftRequest;
  if (!body.subject || !body.to || body.to.length === 0) {
    return { status: 400, jsonBody: { error: 'subject and to[] are required' } };
  }

  let attachmentRequests;
  try {
    attachmentRequests = parseAttachments(body.attachments);
  } catch (err) {
    return { status: 400, jsonBody: { error: (err as Error).message } };
  }

  // Drafts are written to the Drafts folder — check deny list by display name
  const draftsDeny = await checkDenyList(userId, 'mail', 'Drafts');
  if (draftsDeny) {
    return { status: draftsDeny.status, jsonBody: { error: draftsDeny.error } };
  }

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  let from: string | undefined;
  try {
    from = body.from === undefined ? undefined : validateFromAddress(body.from, await getMailboxAddresses(graph, '/me'));
  } catch (err) {
    if (err instanceof FromAddressError) return { status: 400, jsonBody: { error: err.message } };
    throw err;
  }

  const message = {
    subject: body.subject,
    body: {
      contentType: body.bodyType === 'html' ? 'HTML' : 'Text',
      content: body.body ?? '',
    },
    toRecipients: body.to.map((addr) => ({ emailAddress: { address: addr } })),
    ...(body.cc ? { ccRecipients: body.cc.map((addr) => ({ emailAddress: { address: addr } })) } : {}),
    ...(body.bcc ? { bccRecipients: body.bcc.map((addr) => ({ emailAddress: { address: addr } })) } : {}),
    ...(from ? { from: toGraphFrom(from) } : {}),
  };

  // Drive-item attachments are fetched after the Drafts deny check above.
  let attachments;
  try {
    attachments = await resolveAttachments(graph, attachmentRequests, {
      session: auth.session,
      userId,
      tenantId: getTenantIdFromSession(auth.session),
      operation: 'mail.draft',
      source: 'http',
    });
  } catch (err) {
    if (err instanceof AttachmentSourceError) return { status: err.status, jsonBody: { error: err.message } };
    throw err;
  }

  const result = await createDraftWithAttachments(graph, message, attachments);

  return {
    status: 201,
    jsonBody: {
      id: result.id,
      subject: result.subject,
      webLink: result.webLink,
      createdDateTime: result.createdDateTime,
      status: 'draft',
    },
  };
}

app.http('createDraft', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/mail/drafts',
  handler: withSecurity(withPolicyEnforcement('mail', createDraftHandler)),
});
