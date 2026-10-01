/**
 * Shared email-attachment handling for the mail tools.
 *
 * Both the REST Azure Functions (`sendMail`, `createDraft`) and the JSON-RPC MCP
 * handlers (`send_mail`, `create_draft`) accept an optional `attachments` array. Each
 * item is either `{ name, contentType, content }` where `content` is the base64-encoded
 * file bytes, or a reference to a file already in M365: `{ driveItemId }` for
 * OneDrive or `{ siteId, itemId, driveId? }` for SharePoint. The server fetches referenced
 * files itself (driveAttachments.ts), so their bytes never pass through the caller.
 *
 * Microsoft Graph has two ways to attach a file:
 *   - Files under 3 MB ride along inline on the message POST as a
 *     `#microsoft.graph.fileAttachment` (a single request).
 *   - Files from 3 MB up to 150 MB must be streamed to a saved draft via a
 *     chunked upload session, because the inline form is rejected at that size.
 *
 * This module validates caller input, splits attachments into the inline vs.
 * upload-session buckets, and drives both Graph paths so the four call sites stay
 * consistent (a Codex-review concern: parallel HTTP + MCP paths must not drift).
 *
 * Practical ceiling note: base64 inflates payload ~33%, and the attachment bytes
 * arrive inside the JSON request body, so very large files are bounded by the
 * MCP/HTTP request-size limit long before Graph's 150 MB cap. Keep total request
 * bodies to a few MB in practice; for larger files, share a SharePoint link.
 */

import 'isomorphic-fetch';
import type { Client } from '@microsoft/microsoft-graph-client';
import { assertOpaqueId } from './opaqueId.js';

/**
 * Attachment bytes ready to hand to Graph. `content` is base64-encoded file bytes. Inline caller
 * input normalizes straight to this; a drive reference becomes one once the server has fetched
 * the item (`resolveAttachments` in driveAttachments.ts).
 */
export interface MailAttachmentInput {
  name: string;
  contentType?: string;
  content: string;
}

/**
 * One validated `attachments[]` item, before any bytes are fetched. An item carries
 * exactly one source: inline base64 `content`, a OneDrive `driveItemId`, or a SharePoint
 * `siteId` + `itemId` (optionally `driveId` for a library other than the site's default).
 * For drive items `name` and `contentType` are optional overrides of the item's own values.
 */
export type MailAttachmentRequest =
  | ({ source: 'inline' } & MailAttachmentInput)
  | { source: 'onedrive'; index: number; itemId: string; name?: string; contentType?: string }
  | {
      source: 'sharepoint';
      index: number;
      siteId: string;
      driveId?: string;
      itemId: string;
      name?: string;
      contentType?: string;
    };

/** Graph `#microsoft.graph.fileAttachment` resource for an inline (single-request) attachment. */
export interface GraphFileAttachment {
  '@odata.type': '#microsoft.graph.fileAttachment';
  name: string;
  contentType: string;
  contentBytes: string;
}

// Graph attaches files < 3 MB inline on the message POST; larger files (up to 150 MB)
// require a chunked upload session against a saved draft.
export const INLINE_ATTACHMENT_LIMIT_BYTES = 3 * 1024 * 1024;
export const MAX_ATTACHMENT_BYTES = 150 * 1024 * 1024;
// Upload-session chunk size MUST be a multiple of 320 KiB per the Graph contract
// (except the final chunk). 1.25 MiB keeps request count low without large buffers.
const UPLOAD_CHUNK_BYTES = 4 * 320 * 1024;

const BASE64_RE = /^[A-Za-z0-9+/]*={0,2}$/;
export const DEFAULT_CONTENT_TYPE = 'application/octet-stream';

/** Decoded byte length of a validated base64 string. */
export function decodedByteLength(base64: string): number {
  return Buffer.from(base64, 'base64').length;
}

function optionalString(value: unknown, field: string): string | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== 'string' || value.trim() === '') throw new Error(`${field} must be a non-empty string`);
  return value;
}

function requiredId(value: unknown, field: string): string {
  if (typeof value !== 'string' || value === '') throw new Error(`${field} must be a non-empty string`);
  // Drive and site IDs are interpolated into Graph paths, so hold them to the same
  // rule as top-level opaque ID arguments. The tools/call gate only checks
  // top-level arguments, not fields nested inside attachments[].
  assertOpaqueId(value, field);
  return value;
}

/**
 * Validate a caller-supplied `attachments` value into `MailAttachmentRequest[]`. Returns `[]`
 * when the field is omitted. Throws a descriptive Error on malformed input so the handler can
 * surface a 400-style message rather than a raw Graph failure. Fetches nothing: drive
 * references are resolved later, after the mail-side policy checks have passed.
 */
export function parseAttachments(raw: unknown): MailAttachmentRequest[] {
  if (raw === undefined || raw === null) return [];
  if (!Array.isArray(raw)) throw new Error('attachments must be an array');

  return raw.map((item, i): MailAttachmentRequest => {
    if (typeof item !== 'object' || item === null || Array.isArray(item)) {
      throw new Error(`attachments[${i}] must be an object`);
    }
    const { name, contentType, content, driveItemId, siteId, itemId, driveId } = item as Record<string, unknown>;

    const sources = [
      content !== undefined && 'content',
      driveItemId !== undefined && 'driveItemId',
      (siteId !== undefined || itemId !== undefined || driveId !== undefined) && 'siteId + itemId',
    ].filter(Boolean);
    if (sources.length === 0) {
      throw new Error(
        `attachments[${i}] needs a source: content (base64 file bytes), driveItemId (OneDrive), or siteId + itemId (SharePoint)`,
      );
    }
    if (sources.length > 1) {
      throw new Error(`attachments[${i}] takes exactly one source, got ${sources.join(' and ')}`);
    }

    if (driveItemId !== undefined) {
      return {
        source: 'onedrive',
        index: i,
        itemId: requiredId(driveItemId, `attachments[${i}].driveItemId`),
        name: optionalString(name, `attachments[${i}].name`),
        contentType: optionalString(contentType, `attachments[${i}].contentType`),
      };
    }

    if (content === undefined) {
      if (siteId === undefined || itemId === undefined) {
        throw new Error(`attachments[${i}] needs both siteId and itemId to attach a SharePoint file`);
      }
      return {
        source: 'sharepoint',
        index: i,
        siteId: requiredId(siteId, `attachments[${i}].siteId`),
        itemId: requiredId(itemId, `attachments[${i}].itemId`),
        ...(driveId !== undefined ? { driveId: requiredId(driveId, `attachments[${i}].driveId`) } : {}),
        name: optionalString(name, `attachments[${i}].name`),
        contentType: optionalString(contentType, `attachments[${i}].contentType`),
      };
    }

    if (typeof name !== 'string' || name.trim() === '') {
      throw new Error(`attachments[${i}].name is required`);
    }
    if (typeof content !== 'string' || content === '') {
      throw new Error(`attachments[${i}].content (base64 file bytes) must be a non-empty string`);
    }
    const stripped = content.replace(/\s/g, '');
    if (!BASE64_RE.test(stripped) || stripped.length % 4 !== 0) {
      throw new Error(`attachments[${i}].content must be base64-encoded`);
    }
    const size = decodedByteLength(stripped);
    if (size === 0) {
      throw new Error(`attachments[${i}].content decoded to 0 bytes`);
    }
    if (size > MAX_ATTACHMENT_BYTES) {
      throw new Error(
        `attachments[${i}] ("${name}") is ${size} bytes, over the ${MAX_ATTACHMENT_BYTES}-byte (150 MB) Graph limit`,
      );
    }
    return {
      source: 'inline',
      name,
      contentType: typeof contentType === 'string' && contentType ? contentType : DEFAULT_CONTENT_TYPE,
      content: stripped,
    };
  });
}

/** Convert a normalized attachment into Graph's inline `fileAttachment` shape. */
export function toFileAttachment(a: MailAttachmentInput): GraphFileAttachment {
  return {
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: a.name,
    contentType: a.contentType ?? DEFAULT_CONTENT_TYPE,
    contentBytes: a.content,
  };
}

/** True when the attachment must use an upload session rather than the inline form. */
export function requiresUploadSession(a: MailAttachmentInput): boolean {
  return decodedByteLength(a.content) > INLINE_ATTACHMENT_LIMIT_BYTES;
}

/** Split attachments into inline (small) and upload-session (large) buckets. */
export function partitionAttachments(atts: MailAttachmentInput[]): {
  inline: MailAttachmentInput[];
  large: MailAttachmentInput[];
} {
  const inline: MailAttachmentInput[] = [];
  const large: MailAttachmentInput[] = [];
  for (const a of atts) (requiresUploadSession(a) ? large : inline).push(a);
  return { inline, large };
}

/**
 * Upload a single large attachment to a saved draft message via a Graph upload session.
 * `basePath` is `/me` or `/users/{id}`. The upload URL returned by createUploadSession is
 * pre-authenticated, so the chunk PUTs carry no auth header.
 */
export async function uploadAttachmentViaSession(
  graph: Client,
  basePath: string,
  messageId: string,
  att: MailAttachmentInput,
): Promise<void> {
  const bytes = Buffer.from(att.content, 'base64');
  const size = bytes.length;

  const session = (await graph
    .api(`${basePath}/messages/${messageId}/attachments/createUploadSession`)
    .post({
      AttachmentItem: {
        attachmentType: 'file',
        name: att.name,
        size,
        contentType: att.contentType ?? DEFAULT_CONTENT_TYPE,
      },
    })) as { uploadUrl?: string };

  const uploadUrl = session?.uploadUrl;
  if (!uploadUrl) {
    throw new Error(`createUploadSession returned no uploadUrl for attachment "${att.name}"`);
  }

  for (let start = 0; start < size; start += UPLOAD_CHUNK_BYTES) {
    const end = Math.min(start + UPLOAD_CHUNK_BYTES, size) - 1;
    const chunk = bytes.subarray(start, end + 1);
    const res = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': String(chunk.length),
        'Content-Range': `bytes ${start}-${end}/${size}`,
      },
      body: chunk,
    });
    // 200 = more chunks expected; 201/204 = final chunk accepted (attachment created).
    if (!res.ok) {
      const detail = await res.text().catch(() => '');
      throw new Error(
        `attachment "${att.name}" upload failed at bytes ${start}-${end}/${size}: ${res.status} ${detail}`.trim(),
      );
    }
  }
}

/**
 * Create a draft message with attachments. Inline attachments ride along on the message POST;
 * large ones are streamed to the created draft via upload sessions. Returns the created draft.
 */
export async function createDraftWithAttachments(
  graph: Client,
  message: Record<string, unknown>,
  atts: MailAttachmentInput[],
  basePath = '/me',
): Promise<{ id: string; [k: string]: unknown }> {
  const { inline, large } = partitionAttachments(atts);
  const payload = inline.length ? { ...message, attachments: inline.map(toFileAttachment) } : message;
  const draft = (await graph.api(`${basePath}/messages`).post(payload)) as { id: string };
  for (const a of large) {
    await uploadAttachmentViaSession(graph, basePath, draft.id, a);
  }
  return draft;
}

/**
 * Send a message with attachments. With only inline (small) attachments this is a single
 * `POST /sendMail`. When any attachment needs an upload session, the message is saved as a
 * draft, the large files are streamed in, and the draft is sent via `POST /messages/{id}/send`.
 *
 * NOTE: the upload-session path always saves to Sent Items (Graph's default for `/send`), so
 * `saveToSentItems: false` can only be honored on the inline (single `/sendMail`) path.
 */
export async function sendMailWithAttachments(
  graph: Client,
  message: Record<string, unknown>,
  atts: MailAttachmentInput[],
  saveToSentItems: boolean,
  basePath = '/me',
): Promise<void> {
  const { large } = partitionAttachments(atts);
  if (large.length === 0) {
    const { inline } = partitionAttachments(atts);
    const msg = inline.length ? { ...message, attachments: inline.map(toFileAttachment) } : message;
    await graph.api(`${basePath}/sendMail`).post({ message: msg, saveToSentItems });
    return;
  }
  const draft = await createDraftWithAttachments(graph, message, atts, basePath);
  await graph.api(`${basePath}/messages/${draft.id}/send`).post({});
}
