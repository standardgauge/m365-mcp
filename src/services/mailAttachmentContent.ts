/**
 * Read one email attachment's content for the MCP `get_attachments` tool and the
 * REST `GET /api/mail/messages/{messageId}/attachments?attachmentId=` route.
 *
 * Graph returns three attachment shapes from `GET .../attachments/{id}`:
 *   - `#microsoft.graph.fileAttachment` carries the bytes inline as `contentBytes`.
 *   - `#microsoft.graph.itemAttachment` (an email, event or contact attached to an
 *     email, e.g. a forwarded thread) carries NO `contentBytes`. Its content is only
 *     reachable through `GET .../attachments/{id}/$value`, which returns the item's
 *     raw MIME (`message/rfc822` for an attached email).
 *   - `#microsoft.graph.referenceAttachment` is a link to a cloud file and has no
 *     content to return.
 *
 * Before both call sites read `contentBytes` unconditionally, so an attached
 * email came back as metadata only and the forwarded thread was unreadable. Both
 * call sites now go through this one helper so the HTTP and MCP paths cannot drift.
 */

import type { Client } from '@microsoft/microsoft-graph-client';

type GraphLike = Pick<Client, 'api'>;

const TEXT_MIME_PREFIXES = ['text/', 'application/json', 'application/xml', 'application/javascript', 'application/csv', 'application/vnd.ms-excel', 'message/rfc822'];

export function isTextMime(mimeType: string): boolean {
  return TEXT_MIME_PREFIXES.some((prefix) => mimeType.startsWith(prefix));
}

export type AttachmentKind = 'file' | 'item' | 'reference';

export interface MailAttachmentContent {
  id: string;
  name: string;
  contentType: string;
  size: number;
  attachmentType: AttachmentKind;
  content?: string;
  encoding?: 'utf-8' | 'base64';
}

function kindOf(odataType: unknown): AttachmentKind {
  if (odataType === '#microsoft.graph.itemAttachment') return 'item';
  if (odataType === '#microsoft.graph.referenceAttachment') return 'reference';
  return 'file';
}

const strictUtf8 = new TextDecoder('utf-8', { fatal: true });

/**
 * Encode bytes for the JSON response: text MIME types come back as a UTF-8 string
 * when the bytes are valid UTF-8, everything else (and text that is not valid
 * UTF-8, e.g. an 8-bit latin-1 MIME body) as base64 so no byte is lost.
 */
function encodeContent(bytes: Buffer, mime: string): { content: string; encoding: 'utf-8' | 'base64' } {
  if (isTextMime(mime)) {
    try {
      return { content: strictUtf8.decode(bytes), encoding: 'utf-8' };
    } catch {
      // fall through to base64
    }
  }
  return { content: bytes.toString('base64'), encoding: 'base64' };
}

async function readStream(stream: AsyncIterable<unknown>): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array));
  return Buffer.concat(chunks);
}

/**
 * Fetch one attachment and its content. `attachmentPath` is the Graph path to the
 * attachment resource, e.g. `/me/messages/{messageId}/attachments/{attachmentId}`.
 * The caller is responsible for the deny-list check on the parent message.
 */
export async function readMailAttachment(graph: GraphLike, attachmentPath: string): Promise<MailAttachmentContent> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const att: any = await graph.api(attachmentPath).get();
  const attachmentType = kindOf(att['@odata.type']);
  const meta = { id: att.id, name: att.name, size: att.size, attachmentType };

  if (attachmentType === 'reference') {
    return { ...meta, contentType: att.contentType ?? 'application/octet-stream' };
  }

  if (attachmentType === 'item') {
    // An attached Outlook item has no contentBytes; its raw MIME is served at /$value.
    const mime: string = att.contentType ?? 'message/rfc822';
    const bytes = await readStream(await graph.api(`${attachmentPath}/$value`).getStream());
    return { ...meta, contentType: mime, ...encodeContent(bytes, mime) };
  }

  const mime: string = att.contentType ?? 'application/octet-stream';
  return { ...meta, contentType: mime, ...encodeContent(Buffer.from(att.contentBytes ?? '', 'base64'), mime) };
}
