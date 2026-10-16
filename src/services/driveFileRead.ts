/**
 * Bounded download of a OneDrive or SharePoint file's content for the read tools.
 *
 * `read_file` and `read_onedrive_file` (MCP and REST) return the whole file inline in
 * the response, so an unbounded read of a multi-gigabyte file holds all of it in memory,
 * then again as base64, and runs into function timeouts or a restart. Every read goes
 * through `readDriveFileContent`, which:
 *   - refuses before any bytes are pulled when the item metadata's `size` is over the cap,
 *     so `getStream()` is never called for a file known to be too large;
 *   - counts bytes while streaming and stops as soon as the cap is crossed, because the
 *     metadata size can be stale or missing.
 *
 * The cap is `MAX_READ_FILE_BYTES` (bytes), defaulting to the 150 MB Graph attachment
 * limit that bounds the mail attachment path.
 */

import type { Client } from '@microsoft/microsoft-graph-client';

type GraphLike = Pick<Client, 'api'>;

export const DEFAULT_MAX_READ_FILE_BYTES = 150 * 1024 * 1024;

/**
 * The readable-size cap from MAX_READ_FILE_BYTES. Unset, empty, zero or non-integer
 * values fall back to DEFAULT_MAX_READ_FILE_BYTES, so a typo never removes the cap.
 */
export function maxReadFileBytes(): number {
  const raw = (process.env.MAX_READ_FILE_BYTES ?? '').trim();
  if (raw === '') return DEFAULT_MAX_READ_FILE_BYTES;
  if (!/^\d+$/.test(raw) || Number(raw) === 0 || !Number.isSafeInteger(Number(raw))) {
    console.error(`[driveFileRead] Ignoring invalid MAX_READ_FILE_BYTES=${JSON.stringify(raw)}; using ${DEFAULT_MAX_READ_FILE_BYTES}`);
    return DEFAULT_MAX_READ_FILE_BYTES;
  }
  return Number(raw);
}

/** A file over the readable-size cap. The REST routes return it as a 413. */
export class FileTooLargeError extends Error {
  constructor(
    readonly fileName: string,
    readonly size: number | undefined,
    readonly limit: number,
  ) {
    super(
      size !== undefined
        ? `"${fileName}" is ${size} bytes, over the ${limit}-byte limit for reading a file inline (MAX_READ_FILE_BYTES)`
        : `"${fileName}" is over the ${limit}-byte limit for reading a file inline (MAX_READ_FILE_BYTES)`,
    );
    this.name = 'FileTooLargeError';
  }
}

export interface DriveFileMetadata {
  name?: string;
  size?: number;
}

/**
 * Download the content at `contentPath` (a Graph `.../items/{id}/content` path), refusing
 * with FileTooLargeError when `metadata.size` or the streamed byte count exceeds the cap.
 * The caller has already fetched `metadata` and run the deny-list and is-a-file checks.
 */
export async function readDriveFileContent(
  graph: GraphLike,
  contentPath: string,
  metadata: DriveFileMetadata,
  limit: number = maxReadFileBytes(),
): Promise<Buffer> {
  const name = metadata.name ?? '';
  if (typeof metadata.size === 'number' && metadata.size > limit) {
    throw new FileTooLargeError(name, metadata.size, limit);
  }

  // Throwing out of for-await calls the iterator's return(), which destroys the stream.
  const stream = (await graph.api(contentPath).getStream()) as AsyncIterable<unknown>;
  const chunks: Buffer[] = [];
  let total = 0;
  for await (const chunk of stream) {
    const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array);
    total += buf.length;
    // The metadata size can be stale or absent; don't trust it to bound the download.
    if (total > limit) {
      throw new FileTooLargeError(name, undefined, limit);
    }
    chunks.push(buf);
  }
  return Buffer.concat(chunks);
}
