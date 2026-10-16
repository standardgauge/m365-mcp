/**
 * What an MCP tool failure is allowed to tell the client.
 *
 * A tools/call handler can fail for two kinds of reason. Some failures are
 * meant for the agent: a policy refusal, a bad argument, a Graph 4xx the user
 * can act on. The agent needs that text to recover. Everything else is an
 * internal failure: a Graph 5xx, a storage error, a network error, a bug. Its
 * message can carry Graph or provider internals, and the client on the other
 * end may be an agent key that has been compromised, so it gets a generic
 * message and a reference id, and the detail goes to the server log.
 *
 * The split is an allow-list by type: a message reaches the client only when
 * the error is one of the caller-facing classes or an actionable real
 * GraphError. Message text never decides it, so a plain Error that happens to
 * contain a policy phrase is still internal. Anything new defaults to generic.
 */
import { randomUUID } from 'node:crypto';
import { GraphError } from '@microsoft/microsoft-graph-client';
import { ValidationError } from './opaqueId.js';
import { FromAddressError } from './mailFrom.js';
import { OutboundPolicyError } from './outboundPolicy.js';
import { AttachmentSourceError } from './driveAttachments.js';
import { ToolCallError } from './toolCallError.js';

export { ToolCallError };

/**
 * Graph statuses whose message tells the user something they can act on: the
 * item is missing, the request was malformed, they lack access, the item
 * changed underneath them, it is locked, or they are being throttled. 401 is
 * left out (a token problem the server handles), as are statuses that only
 * mean this server built the request wrong (405, 406, 411, 415).
 */
const ACTIONABLE_GRAPH_STATUSES = new Set([400, 403, 404, 409, 410, 412, 413, 416, 422, 423, 429]);

/** Error classes defined in this codebase whose messages are written for the caller. */
const CLIENT_FACING_CLASSES = [
  ToolCallError,
  ValidationError,
  FromAddressError,
  OutboundPolicyError,
  AttachmentSourceError,
] as const;

/**
 * The text a failure may show the client, or null when it is internal and must
 * be replaced by the generic message. A policy refusal reaches the agent
 * because it is thrown as a caller-facing class, not because of its wording.
 */
export function clientFacingMessage(err: unknown): string | null {
  if (!(err instanceof Error)) return null;
  if (CLIENT_FACING_CLASSES.some((cls) => err instanceof cls)) return err.message;
  // A Graph error is judged by its status alone: its text is Graph's, so a 5xx
  // that happens to contain a policy phrase is still internal.
  if (err instanceof GraphError) {
    if (!ACTIONABLE_GRAPH_STATUSES.has(err.statusCode)) return null;
    const code = typeof err.code === 'string' && err.code ? ` ${err.code}` : '';
    return `Microsoft Graph returned ${err.statusCode}${code}: ${err.message}`;
  }
  return null;
}

/** A fresh id that ties the generic client message to the server log line. */
export function newErrorReference(): string {
  return randomUUID();
}

/** The message an internal failure shows the client. */
export function genericToolErrorMessage(reference: string): string {
  return (
    'The tool call failed because of an internal server error. ' +
    `Reference: ${reference}. An administrator can find the details in the server logs under this reference.`
  );
}

/**
 * The detail logged server-side for a failure: its message, the Graph status
 * and code when there is one, and the message of each cause in the chain.
 */
export function errorLogDetail(err: unknown): string {
  const parts: string[] = [];
  let current: unknown = err;
  for (let depth = 0; current !== undefined && current !== null && depth < 5; depth++) {
    if (current instanceof Error) {
      const status = (current as { statusCode?: unknown }).statusCode;
      const code = (current as { code?: unknown }).code;
      const tags = [
        typeof status === 'number' ? `status=${status}` : '',
        typeof code === 'string' && code ? `code=${code}` : '',
      ].filter(Boolean).join(' ');
      parts.push(`${current.name}: ${current.message}${tags ? ` (${tags})` : ''}`);
      current = (current as { cause?: unknown }).cause;
    } else {
      parts.push(String(current));
      break;
    }
  }
  return parts.join(' <- caused by ');
}
