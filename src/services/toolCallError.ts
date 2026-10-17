/**
 * The error a tool handler throws when its message is written for the MCP
 * client. Its own module, with no imports, so any service can throw it without
 * pulling in the classifier in toolErrors.ts and the modules it depends on.
 */

/**
 * A failure whose message is written for the MCP client: a refusal, an
 * argument the caller has to fix, or a state the user has to change. Throw it
 * from a tool handler instead of a plain Error when the message must reach the
 * agent. The text must not quote an upstream error verbatim; pass that as
 * `cause` so it is logged and not returned.
 */
export class ToolCallError extends Error {
  constructor(message: string, options?: { cause?: unknown }) {
    super(message, options);
    this.name = 'ToolCallError';
  }
}
