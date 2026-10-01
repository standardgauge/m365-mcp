/**
 * Container-resolution helpers for deny-list enforcement.
 *
 * The deny list stores human-readable identifiers (mail folder display names,
 * calendar IDs, contact folder IDs, notebook IDs, etc.) but runtime operations
 * often only have opaque Graph API identifiers (e.g. parentFolderId on a
 * message is a Graph GUID, not "Inbox"). These helpers resolve the container
 * metadata needed for accurate deny-list matching.
 *
 * All lookups use minimal $select to keep them fast, and results are cached
 * in-memory per process lifetime to avoid repeated Graph calls for the same ID.
 */

import { Client } from '@microsoft/microsoft-graph-client';

// ── In-memory caches (keyed by Graph resource ID) ─────────────────────────

const mailFolderNameCache = new Map<string, string>();
const defaultCalendarIdCache = new Map<string, string>();
const calendarNameCache = new Map<string, string>();
const contactParentFolderCache = new Map<string, string>();
const sectionNotebookIdCache = new Map<string, string>();
const defaultContactFolderCache = new Map<string, string>();

/**
 * Resolve the display name of a mail folder given its Graph folder ID.
 * The deny list stores folder display names (e.g. "Inbox", "Sent Items").
 */
export async function resolveMailFolderName(
  graph: Client,
  folderId: string,
  mailboxBase = '/me',
): Promise<string> {
  const cacheKey = `${mailboxBase}:${folderId}`;
  const cached = mailFolderNameCache.get(cacheKey);
  if (cached !== undefined) return cached;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const folder: any = await graph
    .api(`${mailboxBase}/mailFolders/${folderId}`)
    .select('displayName')
    .get();

  const name: string = folder.displayName ?? '';
  mailFolderNameCache.set(cacheKey, name);
  return name;
}

/**
 * Resolve the default calendar's ID for the current user.
 * When calendar operations omit calendarId and use `/me/events`,
 * they operate on the default calendar — but the deny list may block it.
 */
export async function resolveDefaultCalendarId(
  graph: Client,
  userId: string,
): Promise<string> {
  const cached = defaultCalendarIdCache.get(userId);
  if (cached !== undefined) return cached;

  // /me/calendar returns the default calendar directly
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const cal: any = await graph.api('/me/calendar').select('id').get();
  const calId: string = cal.id ?? '';
  defaultCalendarIdCache.set(userId, calId);
  return calId;
}

/**
 * Resolve the display name of a calendar given its Graph calendar ID.
 *
 * Admins may add a deny-list entry by calendar *name* (human-readable) rather
 * than by the opaque calendar ID. To enforce those entries, single-event reads
 * resolve the effective calendar's name and check it against the deny list in
 * addition to the ID. Returns '' when the name cannot be resolved.
 *
 * The caller MUST validate `calendarId` with assertOpaqueId before passing it
 * here — it is interpolated into the Graph path. The only other source is
 * resolveDefaultCalendarId, whose value originates from Graph itself.
 */
export async function resolveCalendarName(
  graph: Client,
  calendarId: string,
): Promise<string> {
  const cached = calendarNameCache.get(calendarId);
  if (cached !== undefined) return cached;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const cal: any = await graph.api(`/me/calendars/${calendarId}`).select('name').get();
  const name: string = cal.name ?? '';
  calendarNameCache.set(calendarId, name);
  return name;
}

/**
 * Resolve the parentFolderId of a contact by its contactId.
 * For direct contact operations (update/delete by ID), we need to
 * check whether the contact's folder is denied.
 *
 * When the contact lives in the root contacts folder (i.e. its Graph
 * parentFolderId matches the default contacts folder GUID), we return the
 * synthetic `contacts-root` key so that deny-list comparisons match what
 * the admin UI stores.
 */
export async function resolveContactParentFolder(
  graph: Client,
  userId: string,
  contactId: string,
): Promise<string> {
  const cached = contactParentFolderCache.get(contactId);
  if (cached !== undefined) return cached;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const contact: any = await graph
    .api(`/me/contacts/${contactId}`)
    .select('parentFolderId')
    .get();

  const rawFolderId: string = contact.parentFolderId ?? '';

  // Normalize: if this is the default contacts folder, map to the synthetic key
  const folderId = await normalizeContactFolderId(graph, userId, rawFolderId);
  contactParentFolderCache.set(contactId, folderId);
  return folderId;
}

/**
 * Resolve the default contacts folder ID.
 * When creating a contact without specifying a folderId, the contact goes
 * to the root contacts folder. We need to check that folder against the
 * deny list.
 */
export async function resolveDefaultContactFolder(
  graph: Client,
  userId: string,
): Promise<string> {
  const cached = defaultContactFolderCache.get(userId);
  if (cached !== undefined) return cached;

  // The default contacts folder can be read by listing /me/contactFolders
  // and finding the one with no parentFolderId, or by creating a contact
  // and reading its parentFolderId. The simplest approach: fetch one contact
  // and read its parentFolderId, or use the well-known default.
  // Graph doesn't have a direct "default contacts folder" endpoint,
  // but contacts created at /me/contacts go to the root folder whose
  // parentFolderId === null. We use the synthetic 'contacts-root' value
  // that listContactFolders.ts uses.
  const folderId = 'contacts-root';
  defaultContactFolderCache.set(userId, folderId);
  return folderId;
}

// ── Default contacts folder GUID cache ───────────────────────────────────

const defaultContactFolderGuidCache = new Map<string, string>();

/**
 * Resolve the actual Graph GUID of the default contacts folder.
 * We fetch one contact from `/me/contacts` and read its `parentFolderId`.
 * This GUID is what Graph returns for any contact in the root folder.
 *
 * Cache is keyed by userId because folder GUIDs are mailbox-specific —
 * different users have different default contacts folder IDs.
 */
async function resolveDefaultContactFolderGuid(graph: Client, userId: string): Promise<string> {
  const cached = defaultContactFolderGuidCache.get(userId);
  if (cached !== undefined) return cached;

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result: any = await graph
      .api('/me/contacts')
      .select('parentFolderId')
      .top(1)
      .get();
    const guid: string = result.value?.[0]?.parentFolderId ?? '';
    if (guid) defaultContactFolderGuidCache.set(userId, guid);
    return guid;
  } catch {
    return '';
  }
}

/**
 * Normalize a raw parentFolderId from Graph to the synthetic `contacts-root`
 * key when it matches the default contacts folder GUID. This ensures deny-list
 * comparisons align with what the admin UI stores.
 */
export async function normalizeContactFolderId(
  graph: Client,
  userId: string,
  rawFolderId: string,
): Promise<string> {
  if (!rawFolderId) return 'contacts-root';

  const defaultGuid = await resolveDefaultContactFolderGuid(graph, userId);
  if (defaultGuid && rawFolderId.toLowerCase() === defaultGuid.toLowerCase()) {
    return 'contacts-root';
  }
  return rawFolderId;
}

/**
 * Resolve the parent notebook ID for a OneNote section.
 * The deny list may block specific notebooks, but createPage only
 * receives a sectionId — we need to look up which notebook owns it.
 */
export async function resolveSectionNotebook(
  graph: Client,
  sectionId: string,
): Promise<string> {
  const cached = sectionNotebookIdCache.get(sectionId);
  if (cached !== undefined) return cached;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const section: any = await graph
    .api(`/me/onenote/sections/${sectionId}`)
    .select('parentNotebook')
    .get();

  const notebookId: string = section.parentNotebook?.id ?? '';
  sectionNotebookIdCache.set(sectionId, notebookId);
  return notebookId;
}

/** Clear all caches — useful for testing. */
export function clearResolverCaches(): void {
  mailFolderNameCache.clear();
  defaultCalendarIdCache.clear();
  calendarNameCache.clear();
  contactParentFolderCache.clear();
  sectionNotebookIdCache.clear();
  defaultContactFolderCache.clear();
  defaultContactFolderGuidCache.clear();
}
