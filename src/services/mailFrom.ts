/**
 * Send-as-alias support for the mail write tools.
 *
 * `mailboxId` selects which mailbox a draft is written to; it does not select the From
 * address. Without an explicit `from`, Graph stamps the mailbox's primary SMTP address, so a
 * reply to a thread that came in on an alias (nate@standardgauge.ai) goes out from the primary
 * (nate@example.com).
 *
 * Two rules live here:
 *
 *   1. An explicit `from` must be one of the mailbox's own proxy addresses. Anything else is
 *      refused before a draft is created. Sending as a different mailbox is a Send-As grant on
 *      that mailbox, reached through `mailboxId`, not through `from`.
 *   2. Replies and forwards with no `from` default to the address of ours the original was
 *      sent to (its From when it is our own message, else the first match in To, then CC).
 *      That is Outlook's own reply-from-alias behavior. When the match is the primary, nothing
 *      is set and Graph's default applies, so non-alias mail behaves exactly as before.
 *
 * Whether Exchange honors the alias on the wire depends on the tenant's
 * `SendFromAliasEnabled`; with it off, Exchange rewrites From to the primary. See
 * the send-as-alias notes for the verification procedure.
 */

import type { Client } from '@microsoft/microsoft-graph-client';

/** A mailbox's SMTP addresses, keyed by lowercase address to the address as Exchange spells it. */
export interface MailboxAddresses {
  primary: string | null;
  addresses: Map<string, string>;
}

/** A caller-supplied `from` that is not one of the mailbox's addresses. Maps to a 400 on REST. */
export class FromAddressError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'FromAddressError';
  }
}

/**
 * Parse a Graph user's `proxyAddresses` (`SMTP:primary@x`, `smtp:alias@y`, plus non-SMTP
 * entries such as `SIP:` / `X500:` which are ignored) and `mail` into a lookup.
 */
export function parseMailboxAddresses(proxyAddresses: unknown, mail: unknown): MailboxAddresses {
  const addresses = new Map<string, string>();
  let primary: string | null = null;
  for (const entry of Array.isArray(proxyAddresses) ? proxyAddresses : []) {
    if (typeof entry !== 'string') continue;
    const m = /^(smtp):(.+)$/i.exec(entry.trim());
    if (!m) continue;
    const address = m[2].trim();
    addresses.set(address.toLowerCase(), address);
    if (m[1] === 'SMTP') primary = address;
  }
  if (typeof mail === 'string' && mail.trim()) {
    const address = mail.trim();
    if (!addresses.has(address.toLowerCase())) addresses.set(address.toLowerCase(), address);
    primary ??= address;
  }
  return { primary, addresses };
}

/** Read the SMTP addresses of the mailbox at `basePath` (`/me` or `/users/{id}`). */
export async function getMailboxAddresses(graph: Client, basePath: string): Promise<MailboxAddresses> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const user: any = await graph.api(basePath).select('mail,proxyAddresses').get();
  return parseMailboxAddresses(user?.proxyAddresses, user?.mail);
}

/**
 * Validate a caller-supplied `from` against the mailbox's addresses. Returns the address as
 * Exchange spells it, or throws FromAddressError.
 */
export function validateFromAddress(requested: unknown, mailbox: MailboxAddresses): string {
  if (typeof requested !== 'string' || !requested.trim()) {
    throw new FromAddressError('`from` must be a non-empty email address');
  }
  const match = mailbox.addresses.get(requested.trim().toLowerCase());
  if (!match) {
    const allowed = [...mailbox.addresses.values()].join(', ') || '(none found)';
    throw new FromAddressError(
      `\`from\` "${requested.trim()}" is not an address of this mailbox. ` +
      `Use one of its proxy addresses: ${allowed}. To send as a different mailbox, use mailboxId.`,
    );
  }
  return match;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type GraphRecipient = any;

function recipientAddress(r: GraphRecipient): string | null {
  const a = r?.emailAddress?.address;
  return typeof a === 'string' && a.trim() ? a.trim().toLowerCase() : null;
}

/**
 * The From a reply or forward should default to: the mailbox address the original was sent
 * from (replying to our own message) or addressed to (To, then CC). Returns null when the match
 * is the primary address or nothing matches, meaning "leave Graph's default alone".
 */
export function pickReplyFrom(
  original: { from?: GraphRecipient; toRecipients?: GraphRecipient[]; ccRecipients?: GraphRecipient[] },
  mailbox: MailboxAddresses,
): string | null {
  const candidates = [original.from, ...(original.toRecipients ?? []), ...(original.ccRecipients ?? [])];
  for (const r of candidates) {
    const a = recipientAddress(r);
    if (!a) continue;
    const match = mailbox.addresses.get(a);
    if (!match) continue;
    if (mailbox.primary && match.toLowerCase() === mailbox.primary.toLowerCase()) return null;
    return match;
  }
  return null;
}

/** Graph `from` property for an address. */
export function toGraphFrom(address: string): { emailAddress: { address: string } } {
  return { emailAddress: { address } };
}
