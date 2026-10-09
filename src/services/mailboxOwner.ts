/**
 * Resolves whose per-user deny lists apply when a caller reaches a mailbox
 * (and the calendar inside it) through `mailboxId`.
 *
 * Tier 2 deny entries are keyed by the Entra object ID of the user who added
 * them. A delegate's agent that opens another user's mailbox is the caller,
 * but the folders it reads belong to the owner — and the owner's tier 2 list
 * is how they hid those folders from AI. So delegated access is checked
 * against both lists: the caller's (which it always was) and the owner's.
 *
 * The owner's object ID is looked up in the directory (`/users/{mailboxId}`,
 * covered by Directory.Read.All) because `mailboxId` may be a UPN as well as
 * an object ID. A failed lookup throws: an owner we cannot identify is an
 * owner whose list we cannot apply, so the operation is refused rather than
 * run against the caller's list alone.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { encodeGraphId } from './opaqueId.js';
import type { DenySubject } from './denyList.js';

// mailboxId (lower-cased) → owner object ID. Object IDs and UPNs are unique
// across tenants, and the mapping does not change for the life of a mailbox.
const ownerIdCache = new Map<string, string>();

/** Test hook — the cache outlives a single test otherwise. */
export function clearMailboxOwnerCache(): void {
  ownerIdCache.clear();
}

/**
 * Returns the deny subject for an operation on `mailboxId`: the caller alone
 * for their own mailbox (`undefined`, `''` or `'me'`, or an ID that resolves
 * to the caller), otherwise `[callerId, ownerId]`.
 *
 * The caller MUST validate `mailboxId` with assertOpaqueId / assertOpaqueIds
 * first, as for every other ID that reaches a Graph path.
 */
export async function resolveDenySubject(
  graph: Client,
  callerId: string,
  mailboxId: unknown,
): Promise<DenySubject> {
  if (typeof mailboxId !== 'string' || mailboxId === '' || mailboxId === 'me') return callerId;
  const key = mailboxId.toLowerCase();
  if (key === callerId.toLowerCase()) return callerId;

  let ownerId = ownerIdCache.get(key);
  if (ownerId === undefined) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const user: any = await graph.api(`/users/${encodeGraphId(mailboxId, 'mailboxId')}`).select('id').get();
    if (typeof user?.id !== 'string' || user.id === '') {
      throw new Error('Could not identify the owner of that mailbox');
    }
    ownerId = user.id as string;
    ownerIdCache.set(key, ownerId);
  }
  return ownerId.toLowerCase() === callerId.toLowerCase() ? callerId : [callerId, ownerId];
}
