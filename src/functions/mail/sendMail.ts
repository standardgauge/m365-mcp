import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantIdFromSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement, checkDenyList } from '../../services/policyEnforcement.js';
import { getUserEmailSettings } from '../../services/userEmailSettings.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import {
  parseAttachments,
  createDraftWithAttachments,
  sendMailWithAttachments,
  requiresUploadSession,
  type MailAttachmentInput,
} from '../../services/mailAttachments.js';
import { resolveAttachments, AttachmentSourceError } from '../../services/driveAttachments.js';
import { getMailboxAddresses, validateFromAddress, toGraphFrom, FromAddressError } from '../../services/mailFrom.js';

interface SendRequest {
  subject: string;
  body: string;
  bodyType?: 'text' | 'html';
  to: string[];
  cc?: string[];
  bcc?: string[];
  saveToSentItems?: boolean;
  /** Inline `{ name, contentType, content }` items, or drive references. */
  attachments?: unknown[];
  /** Send as this proxy address of the mailbox. */
  from?: string;
}

/**
 * Fetch any drive-item attachments. Returns the resolved list, or the HTTP response to
 * send when a reference is refused by policy (403) or can't be attached (400).
 */
async function resolveOrReject(
  graph: ReturnType<typeof createGraphClient>,
  requests: ReturnType<typeof parseAttachments>,
  auth: AuthResult,
): Promise<MailAttachmentInput[] | HttpResponseInit> {
  try {
    return await resolveAttachments(graph, requests, {
      session: auth.session,
      userId: auth.userId,
      tenantId: getTenantIdFromSession(auth.session),
      operation: 'mail.send',
      source: 'http',
    });
  } catch (err) {
    if (err instanceof AttachmentSourceError) return { status: err.status, jsonBody: { error: err.message } };
    throw err;
  }
}

async function sendMailHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const tenantId = getTenantIdFromSession(auth.session);

  const body = (await request.json()) as SendRequest;
  if (!body.subject || !body.to || body.to.length === 0) {
    return { status: 400, jsonBody: { error: 'subject and to[] are required' } };
  }

  let attachmentRequests;
  try {
    attachmentRequests = parseAttachments(body.attachments);
  } catch (err) {
    return { status: 400, jsonBody: { error: (err as Error).message } };
  }

  const { emailOutputMode } = await getUserEmailSettings(tenantId, userId);

  if (emailOutputMode === 'draft') {
    // User is in draft mode — save to Drafts instead of sending so they can review first
    const draftsDeny = await checkDenyList(userId, 'mail', 'Drafts', auth.session);
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

    const attachments = await resolveOrReject(graph, attachmentRequests, auth);
    if (!Array.isArray(attachments)) return attachments;
    const result = await createDraftWithAttachments(graph, message, attachments);

    return {
      status: 200,
      jsonBody: {
        status: 'queued_as_draft',
        subject: body.subject,
        to: body.to,
        draftId: result.id,
        draftLink: result.webLink,
      },
    };
  }

  // sendMail saves to Sent Items unless saveToSentItems is false, so check that folder's deny
  // list up front, before any drive-item attachment is read.
  if (body.saveToSentItems !== false) {
    const sentDeny = await checkDenyList(userId, 'mail', 'Sent Items', auth.session);
    if (sentDeny) {
      return { status: sentDeny.status, jsonBody: { error: sentDeny.error } };
    }
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

  // Resolved before the size guard: a drive item's size is only known once it is fetched.
  const attachments = await resolveOrReject(graph, attachmentRequests, auth);
  if (!Array.isArray(attachments)) return attachments;

  // Large attachments force the draft + /messages/{id}/send fallback, which always saves to
  // Sent Items (Graph gives no saveToSentItems knob on /send). We cannot honor
  // saveToSentItems:false on that path, so reject the combination instead of silently saving —
  // that would both break the caller contract and slip past the Sent Items deny-list guard,
  // which only ran above when saveToSentItems was not false.
  if (body.saveToSentItems === false && attachments.some(requiresUploadSession)) {
    return {
      status: 400,
      jsonBody: {
        error:
          'saveToSentItems:false is not supported when any attachment exceeds 3 MB — ' +
          'large attachments are sent via an upload-session draft, which always saves to Sent Items. ' +
          'Omit saveToSentItems (or set it true) to send, or keep attachments under 3 MB.',
      },
    };
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

  await sendMailWithAttachments(graph, message, attachments, body.saveToSentItems !== false);

  return {
    status: 200,
    jsonBody: {
      status: 'sent',
      subject: body.subject,
      to: body.to,
    },
  };
}

app.http('sendMail', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/mail/send',
  handler: withSecurity(withPolicyEnforcement('mail', sendMailHandler)),
});
