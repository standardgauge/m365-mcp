/**
 * Attach a file that already lives in OneDrive or SharePoint by its item ID.
 *
 * `create_draft` and `send_mail` (MCP and REST) accept `attachments[]` items of
 * `{ driveItemId }` or `{ siteId, itemId, driveId? }` as well as inline base64 `content`.
 * This module turns those references into bytes: it fetches each item through Graph with
 * the caller's own token and hands back the same `MailAttachmentInput` an inline attachment
 * normalizes to, so the inline/upload-session split in mailAttachments.ts applies unchanged.
 *
 * Every reference goes through the policy the equivalent read tool enforces, because an
 * attachment is a read of that file followed by an outbound send:
 *   - the source service (`onedrive` / `sharepoint`) must be enabled for the tenant and
 *     not disabled for the user, even though the tool itself is a mail tool;
 *   - a SharePoint `siteId` must be in the tenant's allowedSites list;
 *   - the item's path is checked against the source service's deny list, the same
 *     `parentReference.path + name` that `read_onedrive_file` and `read_file` check;
 *   - folders are refused, and the 150 MB Graph attachment cap is enforced from the item
 *     metadata before any bytes are pulled and again while streaming.
 *
 */

import type { Client } from '@microsoft/microsoft-graph-client';
import type { UserSession } from './tokenCache.js';
import { checkServiceEnabled, checkAllowedSite } from './policyEnforcement.js';
import { isPathDenied } from './denyList.js';
import { logAccess } from './auditLog.js';
import {
  DEFAULT_CONTENT_TYPE,
  MAX_ATTACHMENT_BYTES,
  type MailAttachmentInput,
  type MailAttachmentRequest,
} from './mailAttachments.js';

/** A drive reference that could not be attached. `status` is the HTTP status the REST routes return. */
export class AttachmentSourceError extends Error {
  constructor(
    message: string,
    readonly status: 400 | 403,
  ) {
    super(message);
    this.name = 'AttachmentSourceError';
  }
}

export interface AttachmentResolveContext {
  session: UserSession;
  userId: string;
  tenantId: string;
  /** Calling tool or route, recorded on the audit entry (e.g. `create_draft`). */
  operation: string;
  source: 'mcp' | 'http';
}

type DriveRef = Exclude<MailAttachmentRequest, { source: 'inline' }>;

interface DriveItemMetadata {
  id?: string;
  name?: string;
  size?: number;
  file?: { mimeType?: string };
  parentReference?: { path?: string };
}

function drivePathFor(ref: DriveRef): string {
  if (ref.source === 'onedrive') return '/me/drive';
  return ref.driveId ? `/sites/${ref.siteId}/drives/${ref.driveId}` : `/sites/${ref.siteId}/drive`;
}

function audit(
  ctx: AttachmentResolveContext,
  resource: string,
  result: 'allowed' | 'denied',
  reason?: string,
): void {
  logAccess({
    tenantId: ctx.tenantId,
    userId: ctx.userId,
    userEmail: ctx.session.email,
    deviceLabel: ctx.session.deviceLabel,
    operation: `${ctx.operation}.attachment`,
    resource,
    result,
    ...(reason ? { reason } : {}),
    source: ctx.source,
  });
}

async function fetchDriveAttachment(
  graph: Client,
  ref: DriveRef,
  ctx: AttachmentResolveContext,
): Promise<MailAttachmentInput> {
  const label = `attachments[${ref.index}]`;
  const service = ref.source;
  const refId = ref.source === 'onedrive' ? `onedrive:${ref.itemId}` : `sharepoint:${ref.siteId}/${ref.itemId}`;

  const deny = (message: string, resource: string): never => {
    audit(ctx, resource, 'denied', message);
    throw new AttachmentSourceError(`${label}: ${message}`, 403);
  };

  const serviceViolation = await checkServiceEnabled(ctx.userId, service, ctx.session);
  if (serviceViolation) deny(serviceViolation.error, refId);
  if (ref.source === 'sharepoint') {
    const siteViolation = await checkAllowedSite(ctx.userId, ref.siteId, ctx.session);
    if (siteViolation) deny(siteViolation.error, refId);
  }

  const itemPath = `${drivePathFor(ref)}/items/${ref.itemId}`;
  const metadata = (await graph
    .api(itemPath)
    .select('id,name,size,file,parentReference')
    .get()) as DriveItemMetadata;

  const filePath = `${metadata.parentReference?.path ?? ''}/${metadata.name ?? ''}`;
  if (await isPathDenied(ctx.tenantId, ctx.userId, service, filePath)) {
    deny('Access restricted by deny list', `${service}:${filePath}`);
  }
  if (!metadata.file) {
    throw new AttachmentSourceError(`${label}: "${metadata.name ?? ref.itemId}" is a folder, not a file`, 400);
  }
  const declaredSize = metadata.size ?? 0;
  if (declaredSize > MAX_ATTACHMENT_BYTES) {
    throw new AttachmentSourceError(
      `${label} ("${metadata.name}") is ${declaredSize} bytes, over the ${MAX_ATTACHMENT_BYTES}-byte (150 MB) Graph limit`,
      400,
    );
  }

  const stream = await graph.api(`${itemPath}/content`).getStream();
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of stream) {
    const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array);
    total += buf.length;
    // The metadata size can be stale; don't trust it to bound the download.
    if (total > MAX_ATTACHMENT_BYTES) {
      throw new AttachmentSourceError(`${label} ("${metadata.name}") is over the 150 MB Graph limit`, 400);
    }
    chunks.push(buf);
  }
  if (total === 0) {
    throw new AttachmentSourceError(`${label} ("${metadata.name}") is empty`, 400);
  }

  audit(ctx, `${service}:${filePath}`, 'allowed');
  return {
    name: ref.name ?? metadata.name ?? ref.itemId,
    contentType: ref.contentType ?? metadata.file.mimeType ?? DEFAULT_CONTENT_TYPE,
    content: Buffer.concat(chunks).toString('base64'),
  };
}

/**
 * Resolve validated attachment requests into bytes, in the caller's order. Inline items pass
 * through; drive references are fetched one at a time so a denial stops the call before
 * anything further is read. Throws `AttachmentSourceError` on a policy denial (403) or an
 * unattachable item (400); Graph errors (item not found, no access) propagate as-is.
 *
 * Call this only after the mail-side checks (Drafts / Sent Items deny list, email output
 * mode) have passed, so a call that is going to be refused anyway never reads a file.
 */
export async function resolveAttachments(
  graph: Client,
  requests: MailAttachmentRequest[],
  ctx: AttachmentResolveContext,
): Promise<MailAttachmentInput[]> {
  const out: MailAttachmentInput[] = [];
  for (const req of requests) {
    if (req.source === 'inline') {
      out.push({ name: req.name, contentType: req.contentType, content: req.content });
    } else {
      out.push(await fetchDriveAttachment(graph, req, ctx));
    }
  }
  return out;
}
