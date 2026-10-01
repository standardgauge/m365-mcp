// Initialize telemetry FIRST — patches console.log to forward to App Insights
import '../../services/telemetry.js';

import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantIdFromSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedPaths, filterDeniedSearchHits } from '../../services/denyList.js';
import { isPathDenied } from '../../services/denyList.js';
import { checkCalendarAccess } from '../../services/calendarAccess.js';
import { resolveMailFolderName, resolveContactParentFolder, resolveDefaultContactFolder, resolveSectionNotebook, normalizeContactFolderId, resolveDefaultCalendarId } from '../../services/containerResolver.js';
import { authenticateRequest } from '../../services/authMiddleware.js';
import type { UserSession } from '../../services/tokenCache.js';
import { filterAndDisambiguateSites } from '../../services/sharepointFilter.js';
import { getEnabledServices, getAllowedSites, getReadOnlyServices } from '../../services/serviceSettings.js';
import { getUserServiceOverrides } from '../../services/userServiceOverrides.js';
import { getUserEmailSettings, setUserEmailSettings, type EmailOutputMode } from '../../services/userEmailSettings.js';
import { isMailIndexingDisabled } from '../../services/userMailConfig.js';
import { logAccess } from '../../services/auditLog.js';
import { assertOpaqueIds, encodeGraphId } from '../../services/opaqueId.js';
import { findUnsupportedArgs, unsupportedArgsMessage } from '../../services/toolArgs.js';
import { resolveMaxResults, toEnvelope, graphCollectionHasMore } from '../../services/resultEnvelope.js';
import { normalizeGraphDateTime, resolveWindow } from '../../services/calendarWindow.js';
import { resolveMailboxTimeZone } from '../../services/mailboxTimeZone.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { resolveAuthUrlBase } from '../../services/frontendUrl.js';
import type { Client } from '@microsoft/microsoft-graph-client';
import { searchMail, listMessages, toMessageSummary, outcomeMeta, type MessageSummary } from '../../services/mailSearch.js';
import {
  buildContactBody,
  shapeContact,
  CONTACT_SELECT_FIELDS,
  ADDRESS_SCHEMA,
  contactsApiPath,
  runContactBatch,
} from '../../services/contactFields.js';
import {
  parseAttachments,
  createDraftWithAttachments,
  sendMailWithAttachments,
} from '../../services/mailAttachments.js';
import { resolveAttachments } from '../../services/driveAttachments.js';
import { getMailboxAddresses, validateFromAddress, pickReplyFrom, toGraphFrom } from '../../services/mailFrom.js';
import {
  SEARCH_DRIVE_ITEM_FIELDS,
  searchFetchSize,
  siteRelativePathFromWebUrl,
  filterHitsToAllowedSites,
  type SearchHitItem,
} from '../../services/sharepointSearch.js';
import { collectPage } from '../../services/graphPaging.js';
import { readMailAttachment } from '../../services/mailAttachmentContent.js';

/**
 * Streamable HTTP MCP endpoint.
 *
 * Implements a simplified MCP-over-HTTP protocol:
 * - POST /api/mcp with JSON-RPC messages
 * - Handles initialize, tools/list, tools/call
 * - User authentication via the `Authorization: Bearer <sessionToken>` header
 *   (the legacy `x-user-id` + shared `x-api-key` impersonation model was removed)
 *
 * This allows Claude Desktop/Code to connect directly as a remote MCP server
 * with no local bridge process needed.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type ToolHandler = (args: Record<string, any>, session: UserSession) => Promise<any>;

const TEXT_MIME_PREFIXES = ['text/', 'application/json', 'application/xml', 'application/javascript', 'application/csv', 'application/vnd.ms-excel'];
function isTextMime(mimeType: string): boolean {
  return TEXT_MIME_PREFIXES.some((prefix) => mimeType.startsWith(prefix));
}

// ── Tool definitions ──

interface ToolDef {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
  handler: ToolHandler;
}

/**
 * Map tool names to their service category for enabledServices enforcement.
 * Tools whose service is not in the tenant's enabledServices list are hidden
 * from tools/list and rejected at tools/call time.
 */
const TOOL_SERVICE_MAP: Record<string, string> = {
  // SharePoint
  list_sites: 'sharepoint', list_folders: 'sharepoint', read_file: 'sharepoint',
  search_sharepoint: 'sharepoint', write_sharepoint_file: 'sharepoint',
  delete_sharepoint_file: 'sharepoint', create_sharepoint_folder: 'sharepoint',
  list_sharepoint_lists: 'sharepoint', list_sharepoint_list_items: 'sharepoint',
  create_sharepoint_list_item: 'sharepoint', update_sharepoint_list_item: 'sharepoint',
  delete_sharepoint_list_item: 'sharepoint',
  // Mail
  list_mailboxes: 'mail', list_folders_mail: 'mail', create_mail_folder: 'mail',
  rename_mail_folder: 'mail', move_mail_folder: 'mail',
  search_mail: 'mail', list_messages: 'mail', read_message: 'mail', get_attachments: 'mail',
  create_draft: 'mail', update_message: 'mail', send_mail: 'mail', send_draft: 'mail',
  reply_to_message: 'mail', reply_all_to_message: 'mail', forward_message: 'mail',
  move_message: 'mail', delete_message: 'mail',
  get_email_output_mode: 'mail', set_email_output_mode: 'mail',
  // OneDrive
  list_onedrive: 'onedrive', read_onedrive_file: 'onedrive',
  write_onedrive_file: 'onedrive', move_onedrive_item: 'onedrive',
  create_onedrive_folder: 'onedrive', delete_onedrive_item: 'onedrive',
  // Calendar
  list_calendars: 'calendar', create_calendar: 'calendar', list_events: 'calendar', get_event: 'calendar',
  create_event: 'calendar', update_event: 'calendar', delete_event: 'calendar',
  move_event: 'calendar', respond_to_event: 'calendar',
  get_schedule: 'calendar', find_meeting_times: 'calendar', list_rooms: 'calendar',
  // Contacts
  search_contacts: 'contacts', create_contact: 'contacts',
  update_contact: 'contacts', delete_contact: 'contacts',
  list_contact_folders: 'contacts', create_contact_folder: 'contacts',
  create_contacts_batch: 'contacts',
  // OneNote
  list_notebooks: 'onenote', create_notebook: 'onenote', list_sections: 'onenote',
  create_section: 'onenote', create_onenote_page: 'onenote',
  // Teams
  list_teams: 'teams', list_channels: 'teams',
  send_chat_message: 'teams', send_channel_message: 'teams',
};

/** Tools that accept a siteId parameter and should be validated against allowedSites. */
const SITE_SCOPED_TOOLS = new Set([
  'list_folders', 'read_file', 'search_sharepoint', 'write_sharepoint_file',
  'delete_sharepoint_file', 'create_sharepoint_folder', 'list_sharepoint_lists',
  'list_sharepoint_list_items', 'create_sharepoint_list_item',
  'update_sharepoint_list_item', 'delete_sharepoint_list_item',
]);

/**
 * Prefix of the error set_email_output_mode throws when an administrator enforces draft
 * mode. The tools/call catch block matches on it to record the refusal as a
 * denied access in the audit log, the same way deny-list and read-only refusals are.
 */
const ENFORCED_MODE_MARKER = 'Email output mode is enforced to draft by your administrator';

/**
 * Tools that mutate data (create / update / delete / send). When a tool's
 * service is in the tenant's readOnlyServices set, these are hidden from
 * tools/list and refused at tools/call time. Read tools — including read-via-POST ones
 * like search_mail, get_schedule, and find_meeting_times — are intentionally
 * excluded so read-only mode never blocks a genuine read. Read-only defaults to
 * empty per tenant, so this set changes nothing until an admin opts a service
 * in (calendar being the driver).
 */
const WRITE_TOOLS: ReadonlySet<string> = new Set([
  // SharePoint
  'write_sharepoint_file', 'delete_sharepoint_file', 'create_sharepoint_folder',
  'create_sharepoint_list_item', 'update_sharepoint_list_item', 'delete_sharepoint_list_item',
  // Mail
  'create_mail_folder', 'rename_mail_folder', 'move_mail_folder',
  'create_draft', 'update_message', 'send_mail', 'move_message', 'delete_message',
  'reply_to_message', 'reply_all_to_message', 'forward_message', 'send_draft',
  // OneDrive
  'write_onedrive_file', 'move_onedrive_item', 'create_onedrive_folder', 'delete_onedrive_item',
  // Calendar
  'create_calendar', 'create_event', 'update_event', 'delete_event', 'move_event', 'respond_to_event',
  // Contacts
  'create_contact', 'update_contact', 'delete_contact',
  'create_contact_folder', 'create_contacts_batch',
  // OneNote
  'create_notebook', 'create_section', 'create_onenote_page',
  // Teams
  'send_chat_message', 'send_channel_message',
]);

/**
 * Tool argument names that are opaque Graph resource IDs and must be validated
 * before interpolation into a Graph API path.  Arguments NOT in this
 * set (e.g. `q`, `path`, `displayName`, `content`) may contain arbitrary text.
 */
const OPAQUE_ID_PARAMS: ReadonlySet<string> = new Set([
  'siteId', 'driveId', 'parentId', 'itemId', 'listId',
  'messageId', 'folderId', 'parentFolderId', 'destinationFolderId',
  'destinationParentFolderId', 'attachmentId', 'mailboxId',
  'calendarId', 'targetCalendarId', 'eventId', 'contactId', 'notebookId', 'sectionId',
  'teamId', 'channelId', 'chatId', 'messageOrEventId',
]);

function prop(type: string, description: string, opts?: Record<string, unknown>) {
  return { type, description, ...opts };
}

/**
 * Drop messages whose parent folder is deny-listed. A folder-scoped call was
 * already checked up front (`folderScoped`), so only a mailbox-wide result set
 * needs the per-message resolve. Shared by search_mail and list_messages.
 */
async function stripDeniedFolders(
  graph: Client,
  base: string,
  tenantId: string,
  userId: string,
  messages: MessageSummary[],
  folderScoped: boolean,
): Promise<MessageSummary[]> {
  if (folderScoped) return messages;
  const kept: MessageSummary[] = [];
  for (const msg of messages) {
    if (msg.folderId) {
      const fn = await resolveMailFolderName(graph, msg.folderId, base);
      if (fn && await isPathDenied(tenantId, userId, 'mail', fn)) continue;
    }
    kept.push(msg);
  }
  return kept;
}

// Per-page `$top` used when walking a SharePoint drive-children or list-items
// collection for offset pagination. Kept at Graph's 200-item page
// ceiling so deep offsets are reached in the fewest round trips.
const SHAREPOINT_PAGE_SIZE = 200;

// Shared input schema for the `attachments` parameter on send_mail / create_draft.
// Each item is { name, contentType, content } where content is the base64-encoded file bytes,
// or a OneDrive / SharePoint item reference the server fetches itself. No `required`
// list: which fields are required depends on the source, and parseAttachments enforces it.
// Files under 3 MB attach inline; larger files (up to 150 MB) use a Graph upload session.
const ATTACHMENTS_SCHEMA = {
  type: 'array',
  description:
    'Optional file attachments. Each item takes exactly one source: { name, contentType, content } ' +
    'with content as base64-encoded file bytes; { driveItemId } to attach a file from the user\'s ' +
    'OneDrive; or { siteId, itemId } (plus driveId for a non-default library) to attach a SharePoint ' +
    'file. For a OneDrive or SharePoint file the server fetches the bytes itself, so prefer it over ' +
    'reading the file and re-sending it as base64; name and contentType default to the file\'s own. ' +
    'Files under 3 MB attach inline; larger files (up to 150 MB) upload via a Graph upload session. ' +
    'Inline base64 inflates size ~33% and travels inside the request body, so keep inline content to a ' +
    'few MB in practice.',
  items: {
    type: 'object',
    properties: {
      name: prop('string', 'File name shown to the recipient (e.g. "budget.xlsx"). Required with content; optional override for drive items.'),
      contentType: prop('string', 'MIME type (e.g. "application/pdf"). Defaults to application/octet-stream, or the drive file\'s own type.'),
      content: prop('string', 'Base64-encoded file bytes.'),
      driveItemId: prop('string', 'OneDrive item ID of a file to attach (from list_onedrive).'),
      siteId: prop('string', 'SharePoint site ID of a file to attach; use with itemId.'),
      itemId: prop('string', 'SharePoint file item ID to attach (from list_folders or search_sharepoint); use with siteId.'),
      driveId: prop('string', 'SharePoint document library (drive) ID, when the file is not in the site\'s default library.'),
    },
  },
} as const;

// Shared contact field properties for create_contact / update_contact /
// create_contacts_batch. The tool surface is flat; buildContactBody()
// translates it to the Graph `contact` resource. Kept in sync with
// CONTACT_SELECT_FIELDS / shapeContact so written fields round-trip through
// search_contacts.
const CONTACT_FIELD_PROPS: Record<string, unknown> = {
  givenName: prop('string', 'First name'),
  surname: prop('string', 'Last name'),
  middleName: prop('string', 'Middle name'),
  nickName: prop('string', 'Nickname'),
  emailAddresses: { type: 'array', items: { type: 'string' }, description: 'Email addresses' },
  businessPhones: { type: 'array', items: { type: 'string' }, description: 'Business phone numbers' },
  homePhones: { type: 'array', items: { type: 'string' }, description: 'Home phone numbers' },
  mobilePhone: prop('string', 'Mobile phone'),
  companyName: prop('string', 'Company'),
  jobTitle: prop('string', 'Job title'),
  personalNotes: prop('string', 'Free-form notes about the contact (Graph personalNotes)'),
  categories: { type: 'array', items: { type: 'string' }, description: 'Category tags — useful for tagging and later bulk-finding an imported set' },
  birthday: prop('string', 'Birthday as an ISO 8601 date/datetime (e.g. "1974-07-04" or "1974-07-04T00:00:00Z")'),
  spouseName: prop('string', 'Spouse or partner name'),
  homeAddress: ADDRESS_SCHEMA,
  businessAddress: ADDRESS_SCHEMA,
  otherAddress: ADDRESS_SCHEMA,
};

/**
 * respond_to_event: map each `response` value to its Graph action segment
 * (POST .../events/{id}/{action}) and the status word returned to the caller.
 */
const EVENT_RESPONSE_ACTIONS: Record<string, { action: string; status: string }> = {
  accept: { action: 'accept', status: 'accepted' },
  tentative: { action: 'tentativelyAccept', status: 'tentativelyAccepted' },
  decline: { action: 'decline', status: 'declined' },
};

// Shared `from` parameter for the mail write tools. mailboxId picks the mailbox; this
// picks which of that mailbox's own addresses the message is sent as.
const FROM_PROP = prop(
  'string',
  "Send as this address. Must be one of the mailbox's own proxy addresses (an alias such as " +
  'alias@example.com on a user@example.com mailbox); anything else is rejected. Omit to use ' +
  'the primary address.',
);
const REPLY_FROM_PROP = prop(
  'string',
  "Send as this address. Must be one of the mailbox's own proxy addresses; anything else is rejected. " +
  'Omit to default to whichever of the mailbox\'s addresses the original was sent to (To, then CC), ' +
  "matching Outlook's reply-from-alias behavior; the primary address when none match.",
);

/**
 * Shared input-schema properties for the reply/forward draft tools.
 * Recipients are not accepted: Graph derives them from the source message
 * (forward_message adds its own `to`).
 */
const REPLY_SCHEMA_PROPS = {
  messageId: prop('string', 'ID of the message being replied to / forwarded (from search_mail or read_message).'),
  comment: prop('string', 'Your new text. Graph inserts it above the quoted original.'),
  bodyType: prop(
    'string',
    "How to interpret `comment`: 'text' (default — escaped and newlines become line breaks) or 'html' (inserted verbatim).",
    { enum: ['text', 'html'] },
  ),
  mailboxId: prop('string', 'Mailbox (default: me)'),
  from: REPLY_FROM_PROP,
} as const;

/**
 * Resolve an explicit `from` for a new message or a draft about to be sent: validated
 * against the mailbox's proxy addresses, undefined when the caller passed none.
 */
async function resolveExplicitFrom(
  graph: Parameters<typeof getMailboxAddresses>[0],
  base: string,
  from: unknown,
): Promise<string | undefined> {
  if (from === undefined) return undefined;
  return validateFromAddress(from, await getMailboxAddresses(graph, base));
}

/**
 * Render a caller-supplied `comment` for Graph's createReply/createReplyAll/createForward
 * `comment` field.
 *
 * Graph splices `comment` into the reply draft's body ahead of the quoted original without
 * escaping it, and Exchange Online builds those drafts as HTML. So plain text has to be
 * escaped here or a `<` in the caller's prose silently eats the rest of the paragraph, and
 * newlines have to become `<br>` or the whole reply collapses onto one line. `bodyType:
 * 'html'` opts out for callers that are deliberately sending markup.
 */
function renderReplyComment(comment: unknown, bodyType: unknown): string {
  const text = typeof comment === 'string' ? comment : '';
  if (bodyType === 'html') return text;
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\r\n|\r|\n/g, '<br>');
}

/**
 * Shared handler for the three Graph draft-creating reply verbs.
 *
 * `POST /messages/{id}/createReply|createReplyAll|createForward` each return a *draft* —
 * nothing is sent — pre-populated with the quoted original body and, critically, the
 * `conversationId` plus `In-Reply-To`/`References` headers that make the reply collapse into
 * the recipient's existing thread. Building a "RE:" message with create_draft cannot set
 * those headers, which is the bug this replaces.
 *
 * Deny-list enforcement covers both ends: the source message's folder (a reply must not be
 * a way to read a denied folder's contents back out through the quoted body) and Drafts,
 * where the new draft lands.
 */
async function createReplyDraft(
  args: Record<string, unknown>,
  session: UserSession,
  action: 'createReply' | 'createReplyAll' | 'createForward',
  extraBody: Record<string, unknown> = {},
): Promise<Record<string, unknown>> {
  const tenantId = getTenantIdFromSession(session);
  const token = await getValidAccessTokenForSession(session);
  const graph = createGraphClient(token);
  const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';

  // Source message must be readable — same check read_message makes. Its addressing is read
  // too, to default the reply's From to the alias it came in on.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const msgMeta: any = await graph.api(`${base}/messages/${args.messageId}`)
    .select('parentFolderId,from,toRecipients,ccRecipients')
    .get();
  if (msgMeta.parentFolderId) {
    const folderName = await resolveMailFolderName(graph, msgMeta.parentFolderId, base);
    if (folderName && await isPathDenied(tenantId, session.userId, 'mail', folderName)) {
      throw new Error('Access restricted by deny list');
    }
  }
  // The reply draft is written to Drafts — same check create_draft makes.
  if (await isPathDenied(tenantId, session.userId, 'mail', 'Drafts')) {
    throw new Error('Access restricted by deny list — Drafts folder is blocked');
  }

  // Resolve From before anything is written, so a rejected `from` never leaves a draft behind.
  // An explicit `from` must validate; the default is best effort, because failing to read the
  // mailbox's addresses must not break replying altogether.
  let from: string | null = null;
  let fromSource: 'explicit' | 'original-recipient' | 'mailbox-default' = 'mailbox-default';
  let fromWarning: string | undefined;
  if (args.from !== undefined) {
    from = validateFromAddress(args.from, await getMailboxAddresses(graph, base));
    fromSource = 'explicit';
  } else {
    try {
      from = pickReplyFrom(msgMeta, await getMailboxAddresses(graph, base));
      if (from) fromSource = 'original-recipient';
    } catch (err) {
      fromWarning = `Could not read the mailbox's addresses to match the original's alias (${(err as Error).message}); the draft uses the primary address. Pass \`from\` to choose one.`;
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  let draft: any = await graph.api(`${base}/messages/${args.messageId}/${action}`).post({
    comment: renderReplyComment(args.comment, args.bodyType),
    ...extraBody,
  });
  if (from) {
    // Set on the created draft rather than through createReply's `message` parameter: PATCHing
    // `from` on a draft is the documented path. If it fails, remove the draft rather than leave
    // one that would go out from the wrong address.
    try {
      draft = { ...draft, ...(await graph.api(`${base}/messages/${draft.id}`).patch({ from: toGraphFrom(from) })) };
    } catch (err) {
      await graph.api(`${base}/messages/${draft.id}`).delete().catch(() => undefined);
      throw new Error(`Could not set the reply's From to ${from}: ${(err as Error).message}`);
    }
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const addresses = (list: any): string[] => (list ?? []).map((r: any) => r?.emailAddress?.address).filter(Boolean);
  return {
    id: draft?.id,
    subject: draft?.subject,
    webLink: draft?.webLink,
    conversationId: draft?.conversationId,
    to: addresses(draft?.toRecipients),
    cc: addresses(draft?.ccRecipients),
    from: from ?? draft?.from?.emailAddress?.address ?? null,
    fromSource,
    ...(fromWarning ? { fromWarning } : {}),
    inReplyTo: args.messageId,
    status: 'draft',
  };
}

/**
 * True when a Graph error means "no item with this ID" — including an ID of the
 * wrong item type (a message ID on an /events path returns ErrorItemNotFound).
 */
function isGraphNotFound(err: unknown): boolean {
  const e = err as { statusCode?: number; code?: string } | null;
  return e?.statusCode === 404 || e?.code === 'ErrorItemNotFound' || e?.code === 'ErrorInvalidIdMalformed';
}

const tools: ToolDef[] = [
  // SharePoint
  {
    name: 'list_sites',
    description: 'List accessible SharePoint sites in the M365 tenant',
    inputSchema: { type: 'object', properties: {}, required: [] },
    handler: async (_args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const result = await graph.api('/sites?search=*').select('id,displayName,webUrl').top(100).get();
      //: filter system sites (contentTypeHub, appcatalog, etc.) and
      // disambiguate displayName collisions so the LLM (or admin UI dropdown)
      // doesn't pick the wrong site silently when two share a name.
      let sites = filterAndDisambiguateSites(result.value ?? []);
      // Enforce allowedSites — if configured, only return allowed sites
      const tenantId = getTenantIdFromSession(session);
      const allowed = await getAllowedSites(tenantId);
      if (allowed.length > 0) {
        const allowedIds = new Set(allowed.map((s) => s.id));
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        sites = sites.filter((s: any) => allowedIds.has(s.id));
      }
      const { items, count, limit, truncated } = toEnvelope(sites, 100, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'list_folders',
    description: 'List a SharePoint folder\'s children (filtered by deny lists). Returns folders only by default; set includeFiles=true to also return files — each file carries the `id` you pass to read_file (list_folders is the only way to obtain a file\'s itemId). Use offset for paging through large folders.',
    inputSchema: {
      type: 'object',
      properties: {
        siteId: prop('string', 'SharePoint site ID'),
        driveId: prop('string', 'Drive ID (default: site default library)'),
        parentId: prop('string', 'Parent folder item ID (default: root)'),
        includeFiles: prop('boolean', 'When true, also return files (each with id/name/size/lastModifiedDateTime) alongside folders. Files carry the itemId required by read_file. Default false.'),
        offset: prop('number', 'Number of children to skip before returning (for paging deep into a large folder). Default 0.'),
        maxResults: prop('number', 'Max children to return (default 100, max 200).'),
      },
      required: ['siteId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const includeFiles = args.includeFiles === true;
      const offset = Math.max(0, Number(args.offset) || 0);
      const maxResults = Math.min(Math.max(1, Number(args.maxResults) || 100), 200);
      const drivePath = args.driveId ? `/sites/${args.siteId}/drives/${args.driveId}` : `/sites/${args.siteId}/drive`;
      const apiPath = args.parentId ? `${drivePath}/items/${args.parentId}/children` : `${drivePath}/root/children`;
      // Select both facets so files and folders can be distinguished and mapped.
      const select = 'id,name,folder,file,size,parentReference,lastModifiedDateTime';
      const fetchFirst = () => {
        let req = graph.api(apiPath).select(select).top(SHAREPOINT_PAGE_SIZE);
        // Push the folder-only constraint to the source unless files are wanted.
        if (!includeFiles) req = req.filter('folder ne null');
        return req.get();
      };
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fetchNext = (link: string) => graph.api(link).get() as Promise<any>;
      const page = await collectPage(fetchFirst, fetchNext, offset, maxResults);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mapped = (page.items as any[]).map((item: any) => {
        const path = `${item.parentReference?.path ?? ''}/${item.name}`;
        return item.folder
          ? { type: 'folder', id: item.id, name: item.name, path, childCount: item.folder?.childCount ?? 0 }
          : { type: 'file', id: item.id, name: item.name, path, size: item.size ?? null, lastModifiedDateTime: item.lastModifiedDateTime ?? null };
      });
      const tenantId = getTenantIdFromSession(session);
      // Files are subject to the same deny-list filtering as folders.
      const items = await filterDeniedPaths(tenantId, session.userId, 'sharepoint', mapped);
      // hasMore/nextOffset drive offset paging; truncated/limit mirror them so
      // every list verb carries the same completeness contract.
      return { items, count: items.length, hasMore: page.hasMore, nextOffset: page.nextOffset, limit: maxResults, truncated: page.hasMore };
    },
  },
  {
    name: 'read_file',
    description: 'Fetch content of a SharePoint file (text as UTF-8, binary as base64)',
    inputSchema: {
      type: 'object',
      properties: {
        siteId: prop('string', 'SharePoint site ID'),
        itemId: prop('string', 'File item ID'),
        driveId: prop('string', 'Drive ID'),
      },
      required: ['siteId', 'itemId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const dp = args.driveId ? `/sites/${args.siteId}/drives/${args.driveId}` : `/sites/${args.siteId}/drive`;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const metadata: any = await graph.api(`${dp}/items/${args.itemId}`).select('id,name,webUrl,size,file,parentReference,createdDateTime,lastModifiedDateTime').get();
      const filePath = `${metadata.parentReference?.path ?? ''}/${metadata.name}`;
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'sharepoint', filePath)) throw new Error('Access restricted by deny list');
      if (!metadata.file) throw new Error('Item is not a file');
      const stream = await graph.api(`${dp}/items/${args.itemId}/content`).getStream();
      const chunks: Buffer[] = [];
      for await (const chunk of stream) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array));
      const buf = Buffer.concat(chunks);
      const mime = metadata.file.mimeType ?? 'application/octet-stream';
      const isText = isTextMime(mime);
      return { id: metadata.id, name: metadata.name, mimeType: mime, size: metadata.size, content: isText ? buf.toString('utf-8') : buf.toString('base64'), encoding: isText ? 'utf-8' : 'base64' };
    },
  },
  {
    name: 'search_sharepoint',
    description: 'Full-text search across SharePoint (restricted to allow-listed sites; deny-listed folders excluded)',
    inputSchema: {
      type: 'object',
      properties: { q: prop('string', 'Search query'), siteId: prop('string', 'Restrict to site'), maxResults: prop('number', 'Max results (default 25)') },
      required: ['q'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      // The allow-list and deny-list are enforced as post-filters below — the
      // `contentSources` source constraint is invalid for a driveItem query and
      // made every search error. Over-fetch a wider window when a scope
      // is active so filtering can still surface maxResults.
      const allowedSites = await getAllowedSites(tenantId);
      // An explicit siteId (already validated in-allow-list by the dispatch gate)
      // narrows the post-filter to that single site; otherwise the full allow-list
      // applies (empty = allow all). This preserves the "restrict to site"
      // semantics that contentSources used to provide at the source.
      const scopeSites = args.siteId ? [{ id: args.siteId }] : allowedSites;
      const maxResults = resolveMaxResults(args.maxResults, 25, 50);
      const size = searchFetchSize(maxResults, scopeSites.length > 0);
      const body = { requests: [{ entityTypes: ['driveItem'], fields: [...SEARCH_DRIVE_ITEM_FIELDS], query: { queryString: args.q }, from: 0, size }] };
      const result = await graph.api('/search/query').post(body);
      const hitsContainer = result.value?.[0]?.hitsContainers?.[0];
      const hits = hitsContainer?.hits ?? [];
      // Graph's search API reports whether the index holds more matches than this
      // page; carry it into the envelope so a capped search never reads as
      // exhaustive ( — the cap horizon was once mistaken for a mailbox floor).
      const moreAvailable = Boolean(hitsContainer?.moreResultsAvailable);
      // Resolve a site-relative path from each hit's webUrl for deny-list
      // matching (parentReference.path is drive-relative and lacks the
      // document-library segment; an absolute webUrl never matches a
      // path-style deny entry). siteId drives the allow-list post-filter.
      const mapped: SearchHitItem[] = hits.map((h: {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        resource?: any;
      }) => ({ id: h.resource?.id, name: h.resource?.name, webUrl: h.resource?.webUrl, siteId: h.resource?.parentReference?.siteId, path: siteRelativePathFromWebUrl(h.resource?.webUrl) ?? undefined, size: h.resource?.size, createdDateTime: h.resource?.createdDateTime, lastModifiedDateTime: h.resource?.lastModifiedDateTime }));
      // Security boundary: drop hits outside the scope, then drop hits whose
      // path could not be resolved (fail closed — an unmatchable path must not
      // slip past the deny list), then deny-list. Slice to maxResults last.
      const siteFiltered = filterHitsToAllowedSites(mapped, scopeSites);
      const pathResolved = siteFiltered.filter((it) => typeof it.path === 'string' && it.path.length > 0);
      const allowed = await filterDeniedSearchHits(tenantId, session.userId, pathResolved);
      const { items, count, limit, truncated } = toEnvelope(allowed, maxResults, moreAvailable);
      return { results: items, count, limit, truncated };
    },
  },
  {
    name: 'write_sharepoint_file',
    description: 'Upload or overwrite a file in a SharePoint document library',
    inputSchema: {
      type: 'object',
      properties: { siteId: prop('string', 'SharePoint site ID'), path: prop('string', 'File path'), content: prop('string', 'File content'), driveId: prop('string', 'Drive ID'), contentType: prop('string', 'MIME type'), encoding: prop('string', 'Content encoding: "utf-8" (default) or "base64" for binary files') },
      required: ['siteId', 'path', 'content'],
    },
    handler: async (args, session) => {
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'sharepoint', args.path)) throw new Error('Access restricted by deny list');
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const dp = args.driveId ? `/sites/${args.siteId}/drives/${args.driveId}` : `/sites/${args.siteId}/drive`;
      const buf = args.encoding === 'base64' ? Buffer.from(args.content, 'base64') : Buffer.from(args.content, 'utf-8');
      const result = await graph.api(`${dp}/root:/${args.path.replace(/^\//, '')}:/content`).header('Content-Type', args.contentType ?? 'text/plain').put(buf);
      return { id: result.id, name: result.name, webUrl: result.webUrl, status: 'written' };
    },
  },
  {
    name: 'delete_sharepoint_file',
    description: 'Delete a file from a SharePoint document library',
    inputSchema: {
      type: 'object',
      properties: { siteId: prop('string', 'SharePoint site ID'), itemId: prop('string', 'File item ID'), driveId: prop('string', 'Drive ID') },
      required: ['siteId', 'itemId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const dp = args.driveId ? `/sites/${args.siteId}/drives/${args.driveId}` : `/sites/${args.siteId}/drive`;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const meta: any = await graph.api(`${dp}/items/${args.itemId}`).select('name,parentReference').get();
      const filePath = `${meta.parentReference?.path ?? ''}/${meta.name}`;
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'sharepoint', filePath)) throw new Error('Access restricted by deny list');
      await graph.api(`${dp}/items/${args.itemId}`).delete();
      return { status: 'deleted', itemId: args.itemId };
    },
  },
  {
    name: 'create_sharepoint_folder',
    description: 'Create a folder in a SharePoint document library',
    inputSchema: {
      type: 'object',
      properties: { siteId: prop('string', 'SharePoint site ID'), name: prop('string', 'Folder name'), parentId: prop('string', 'Parent folder ID'), driveId: prop('string', 'Drive ID') },
      required: ['siteId', 'name'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const dp = args.driveId ? `/sites/${args.siteId}/drives/${args.driveId}` : `/sites/${args.siteId}/drive`;
      // Resolve the parent path to check deny list before creating
      if (args.parentId) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const parentMeta: any = await graph.api(`${dp}/items/${args.parentId}`).select('name,parentReference').get();
        const parentPath = `${parentMeta.parentReference?.path ?? ''}/${parentMeta.name}`;
        const tenantId = getTenantIdFromSession(session);
        if (await isPathDenied(tenantId, session.userId, 'sharepoint', `${parentPath}/${args.name}`)) throw new Error('Access restricted by deny list');
      }
      const apiPath = args.parentId ? `${dp}/items/${args.parentId}/children` : `${dp}/root/children`;
      const result = await graph.api(apiPath).post({ name: args.name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' });
      return { id: result.id, name: result.name, webUrl: result.webUrl, status: 'created' };
    },
  },
  {
    name: 'list_sharepoint_lists',
    description: 'List all SharePoint lists in a site',
    inputSchema: { type: 'object', properties: { siteId: prop('string', 'SharePoint site ID') }, required: ['siteId'] },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const result = await graph.api(`/sites/${args.siteId}/lists`).select('id,displayName,description,lastModifiedDateTime,list').top(100).get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const lists = (result.value ?? []).map((l: any) => ({ id: l.id, displayName: l.displayName, description: l.description ?? '', template: l.list?.template ?? '', lastModifiedDateTime: l.lastModifiedDateTime }));
      const { items, count, limit, truncated } = toEnvelope(lists, 100, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'list_sharepoint_list_items',
    description: 'List items in a SharePoint list with all field data. Use offset to page through lists larger than one page — a specific item deep in a multi-thousand-item library is only reachable this way.',
    inputSchema: {
      type: 'object',
      properties: {
        siteId: prop('string', 'SharePoint site ID'),
        listId: prop('string', 'List ID'),
        offset: prop('number', 'Number of items to skip before returning (for paging deep into a large list). Default 0.'),
        maxResults: prop('number', 'Max items to return (default 50, max 100).'),
      },
      required: ['siteId', 'listId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const offset = Math.max(0, Number(args.offset) || 0);
      const maxResults = Math.min(Math.max(1, Number(args.maxResults) || 50), 100);
      const listPath = `/sites/${args.siteId}/lists/${args.listId}/items`;
      const fetchFirst = () => graph.api(listPath).expand('fields').top(SHAREPOINT_PAGE_SIZE).get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const fetchNext = (link: string) => graph.api(link).get() as Promise<any>;
      const page = await collectPage(fetchFirst, fetchNext, offset, maxResults);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const items = (page.items as any[]).map((item: any) => ({ id: item.id, fields: item.fields ?? {}, lastModifiedDateTime: item.lastModifiedDateTime }));
      return { items, count: items.length, hasMore: page.hasMore, nextOffset: page.nextOffset, limit: maxResults, truncated: page.hasMore };
    },
  },
  {
    name: 'create_sharepoint_list_item',
    description: 'Create a new item in a SharePoint list',
    inputSchema: {
      type: 'object',
      properties: { siteId: prop('string', 'SharePoint site ID'), listId: prop('string', 'List ID'), fields: prop('object', 'Field values as key-value pairs (e.g. {"Title":"My Item","Status":"Active"})') },
      required: ['siteId', 'listId', 'fields'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const result = await graph.api(`/sites/${args.siteId}/lists/${args.listId}/items`).post({ fields: args.fields });
      return { id: result.id, fields: result.fields ?? {}, status: 'created' };
    },
  },
  {
    name: 'update_sharepoint_list_item',
    description: 'Update an existing item in a SharePoint list',
    inputSchema: {
      type: 'object',
      properties: { siteId: prop('string', 'SharePoint site ID'), listId: prop('string', 'List ID'), itemId: prop('string', 'Item ID'), fields: prop('object', 'Field values to update as key-value pairs') },
      required: ['siteId', 'listId', 'itemId', 'fields'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      await graph.api(`/sites/${args.siteId}/lists/${args.listId}/items/${args.itemId}/fields`).patch(args.fields);
      return { id: args.itemId, status: 'updated' };
    },
  },
  {
    name: 'delete_sharepoint_list_item',
    description: 'Delete an item from a SharePoint list',
    inputSchema: {
      type: 'object',
      properties: { siteId: prop('string', 'SharePoint site ID'), listId: prop('string', 'List ID'), itemId: prop('string', 'Item ID') },
      required: ['siteId', 'listId', 'itemId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      await graph.api(`/sites/${args.siteId}/lists/${args.listId}/items/${args.itemId}`).delete();
      return { id: args.itemId, status: 'deleted' };
    },
  },
  // Mail
  {
    name: 'list_mailboxes',
    description: 'List mailboxes the user has access to',
    inputSchema: { type: 'object', properties: {}, required: [] },
    handler: async (_args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const me = await graph.api('/me').select('id,displayName,mail,userPrincipalName').get();
      // Only the primary mailbox is exposed today, so this set is always complete.
      const mailboxes = [{ id: me.id, displayName: me.displayName, email: me.mail ?? me.userPrincipalName, type: 'primary' }];
      return { items: mailboxes, count: mailboxes.length, limit: mailboxes.length, truncated: false };
    },
  },
  {
    name: 'list_folders_mail',
    description: 'List mail folders in a mailbox (filtered by deny lists)',
    inputSchema: {
      type: 'object',
      properties: { mailboxId: prop('string', 'Mailbox user ID (default: me)'), parentFolderId: prop('string', 'Parent folder ID') },
      required: [],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      const apiPath = args.parentFolderId ? `${base}/mailFolders/${args.parentFolderId}/childFolders` : `${base}/mailFolders`;
      const result = await graph.api(apiPath).select('id,displayName,totalItemCount,unreadItemCount,childFolderCount').top(100).get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const folders = (result.value ?? []).map((f: any) => ({ id: f.id, displayName: f.displayName, path: f.displayName, totalItemCount: f.totalItemCount, unreadItemCount: f.unreadItemCount, childFolderCount: f.childFolderCount }));
      // Filter folders by deny list (uses displayName as path — admin deny-list entries match on display name)
      const tenantId = getTenantIdFromSession(session);
      const allowed = await filterDeniedPaths(tenantId, session.userId, 'mail', folders);
      const { items, count, limit, truncated } = toEnvelope(allowed, 100, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'create_mail_folder',
    description: 'Create a mail folder (top-level or as a child of an existing folder). Returns { id, displayName, parentFolderId }.',
    inputSchema: {
      type: 'object',
      properties: {
        displayName: prop('string', 'Folder display name'),
        parentFolderId: prop('string', 'Parent folder ID (default: root mailFolders)'),
        mailboxId: prop('string', 'Mailbox (default: me)'),
      },
      required: ['displayName'],
    },
    handler: async (args, session) => {
      if (!args.displayName) throw new Error('displayName is required');
      // Block creating a folder whose display name is on the deny list — the deny
      // list keys on displayName for mail, so a denied name must not become creatable.
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'mail', args.displayName)) {
        throw new Error('Access restricted by deny list');
      }
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      const apiPath = args.parentFolderId ? `${base}/mailFolders/${args.parentFolderId}/childFolders` : `${base}/mailFolders`;
      const result = await graph.api(apiPath).post({ displayName: args.displayName });
      return { id: result.id, displayName: result.displayName, parentFolderId: result.parentFolderId, status: 'created' };
    },
  },
  {
    name: 'rename_mail_folder',
    description: 'Rename an existing mail folder. Returns { id, displayName, parentFolderId }.',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: prop('string', 'Mail folder ID to rename (from list_folders_mail)'),
        displayName: prop('string', 'New folder display name'),
        mailboxId: prop('string', 'Mailbox (default: me)'),
      },
      required: ['folderId', 'displayName'],
    },
    handler: async (args, session) => {
      const tenantId = getTenantIdFromSession(session);
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      // Block renaming a folder that is itself deny-listed (resolve its current name).
      const currentName = await resolveMailFolderName(graph, args.folderId, base);
      if (currentName && await isPathDenied(tenantId, session.userId, 'mail', currentName)) {
        throw new Error('Access restricted by deny list');
      }
      // Block renaming TO a deny-listed name (parity with create_mail_folder).
      if (await isPathDenied(tenantId, session.userId, 'mail', args.displayName)) {
        throw new Error('Access restricted by deny list');
      }
      const result = await graph.api(`${base}/mailFolders/${args.folderId}`).patch({ displayName: args.displayName });
      return { id: result.id, displayName: result.displayName, parentFolderId: result.parentFolderId, status: 'renamed' };
    },
  },
  {
    name: 'move_mail_folder',
    description: 'Move a mail folder under a different parent folder. Returns { id, displayName, parentFolderId }.',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: prop('string', 'Mail folder ID to move (from list_folders_mail)'),
        destinationParentFolderId: prop('string', 'Target parent mail folder ID'),
        mailboxId: prop('string', 'Mailbox (default: me)'),
      },
      required: ['folderId', 'destinationParentFolderId'],
    },
    handler: async (args, session) => {
      const tenantId = getTenantIdFromSession(session);
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      // Deny-list check on the folder being moved and on the destination parent.
      const movedName = await resolveMailFolderName(graph, args.folderId, base);
      if (movedName && await isPathDenied(tenantId, session.userId, 'mail', movedName)) {
        throw new Error('Access restricted by deny list');
      }
      const destName = await resolveMailFolderName(graph, args.destinationParentFolderId, base);
      if (destName && await isPathDenied(tenantId, session.userId, 'mail', destName)) {
        throw new Error('Destination folder restricted by deny list');
      }
      const result = await graph.api(`${base}/mailFolders/${args.folderId}/move`).post({ destinationId: args.destinationParentFolderId });
      return { id: result.id, displayName: result.displayName, parentFolderId: result.parentFolderId, status: 'moved' };
    },
  },
  {
    name: 'search_mail',
    description:
      'Search email messages. `q` is free text matched as a phrase against subject, body and sender ' +
      '(an email address or domain in `q`, e.g. "alice@example.com" or "example.com", is also matched ' +
      'against every participant). To find mail sent TO someone, pass `participant` or `to` — plain text ' +
      'in `q` does not match recipients on folder-scoped searches. Results without `folderId` are ' +
      'relevance-ranked, not newest-first: the first N are NOT the most recent N. To enumerate the newest ' +
      'messages in a folder, use list_messages. Every response reports `strategy`, `ordering` and ' +
      '`searchedFields`; folder scans also report how far back they looked (`scanHorizon`, `scanComplete`), ' +
      'so an empty result is not proof of absence unless `scanComplete` is true. At least one of `q`, ' +
      '`participant`, `from`, `to` is required.',
    inputSchema: {
      type: 'object',
      properties: {
        q: prop('string', 'Free-text query (phrase match on subject, body, sender). Optional when participant / from / to is given.'),
        participant: prop('string', 'Email address, domain or name fragment matched against the sender and every To / Cc / Bcc recipient. The right way to ask "any mail with this counterparty".'),
        from: prop('string', 'Email address, domain or name fragment the sender must match.'),
        to: prop('string', 'Email address, domain or name fragment any To / Cc / Bcc recipient must match. Use this to check Sent Items for mail to a counterparty.'),
        since: prop('string', 'ISO-8601 date or date-time; only messages received on or after it (e.g. "2026-09-15").'),
        folderId: prop('string', 'Folder ID to scope the search to (e.g. Sent Items).'),
        mailboxId: prop('string', 'Mailbox (default: me)'),
        maxResults: prop('number', 'Max results (default 25, max 100)'),
      },
      required: [],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      // Block search if the requested folder is denied (resolve display name first)
      if (args.folderId) {
        const folderName = await resolveMailFolderName(graph, args.folderId, base);
        if (folderName && await isPathDenied(tenantId, session.userId, 'mail', folderName)) {
          throw new Error('Access restricted by deny list');
        }
      }
      const maxResults = resolveMaxResults(args.maxResults, 25, 100);

      // Routing, matching, fallbacks and the honest-empty-result metadata all live in
      // services/mailSearch.ts and are shared with the /api/mail/search route.
      // Folder-scoped queries never use $search; address criteria in a folder
      // run a newest-first scan so a counterparty's address matches in Sent Items.
      const outcome = await searchMail(graph, {
        base,
        folderId: args.folderId,
        maxResults,
        q: args.q,
        participant: args.participant,
        from: args.from,
        to: args.to,
        since: args.since,
      });

      const filtered = await stripDeniedFolders(graph, base, tenantId, session.userId, outcome.messages.map(toMessageSummary), Boolean(args.folderId));
      // Surface truncation: the defect where a cap horizon was mistaken for the
      // mailbox's earliest message. `moreAvailable` is Graph's nextLink for $search /
      // $filter, and "stopped before the end" for a scan.
      const { items, count, limit, truncated } = toEnvelope(filtered, maxResults, outcome.moreAvailable);
      return { results: items, count, limit, truncated, ...outcomeMeta(outcome) };
    },
  },
  {
    name: 'list_messages',
    description:
      'List the newest messages in a mailbox or folder, newest first, deterministically (`$orderby ' +
      'receivedDateTime desc`). This is the tool for "what was sent / received since X" and "the last ' +
      'N messages in Sent Items" — do not use search_mail with a throwaway query for that, because ' +
      '$search is relevance-ranked and returns a sample, not the newest N. Optional `since` bounds the ' +
      'window; `truncated: true` means more messages exist beyond `limit`.',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: prop('string', 'Folder ID to list (omit for the whole mailbox).'),
        since: prop('string', 'ISO-8601 date or date-time; only messages received on or after it (e.g. "2026-09-15").'),
        mailboxId: prop('string', 'Mailbox (default: me)'),
        maxResults: prop('number', 'Max results (default 25, max 100)'),
      },
      required: [],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      if (args.folderId) {
        const folderName = await resolveMailFolderName(graph, args.folderId, base);
        if (folderName && await isPathDenied(tenantId, session.userId, 'mail', folderName)) {
          throw new Error('Access restricted by deny list');
        }
      }
      const maxResults = resolveMaxResults(args.maxResults, 25, 100);
      const outcome = await listMessages(graph, { base, folderId: args.folderId, since: args.since, maxResults });
      const filtered = await stripDeniedFolders(graph, base, tenantId, session.userId, outcome.messages.map(toMessageSummary), Boolean(args.folderId));
      const { items, count, limit, truncated } = toEnvelope(filtered, maxResults, outcome.moreAvailable);
      return { items, count, limit, truncated, ...outcomeMeta(outcome) };
    },
  },
  {
    name: 'read_message',
    description: 'Fetch full email content including body, recipients, and metadata',
    inputSchema: {
      type: 'object',
      properties: { messageId: prop('string', 'Message ID'), mailboxId: prop('string', 'Mailbox (default: me)') },
      required: ['messageId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const msg: any = await graph.api(`${base}/messages/${args.messageId}`).get();
      if (msg.parentFolderId) {
        const tenantId = getTenantIdFromSession(session);
        const folderName = await resolveMailFolderName(graph, msg.parentFolderId, base);
        if (folderName && await isPathDenied(tenantId, session.userId, 'mail', folderName)) throw new Error('Access restricted by deny list');
      }
      return msg;
    },
  },
  {
    name: 'get_attachments',
    description: "List attachments on an email, or download a specific attachment's content. A file attachment returns its bytes; an attached email (a forwarded message, contentType message/rfc822) returns its raw MIME, including the headers, body and any nested attachments.",
    inputSchema: {
      type: 'object',
      properties: { messageId: prop('string', 'Message ID'), attachmentId: prop('string', 'Attachment ID to download (omit to list)'), mailboxId: prop('string', 'Mailbox (default: me)') },
      required: ['messageId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      // Check deny list via parent folder of the message
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const msgMeta: any = await graph.api(`${base}/messages/${args.messageId}`).select('parentFolderId').get();
      if (msgMeta.parentFolderId) {
        const tenantId = getTenantIdFromSession(session);
        const folderName = await resolveMailFolderName(graph, msgMeta.parentFolderId, base);
        if (folderName && await isPathDenied(tenantId, session.userId, 'mail', folderName)) throw new Error('Access restricted by deny list');
      }
      if (args.attachmentId) {
        return readMailAttachment(graph, `${base}/messages/${args.messageId}/attachments/${args.attachmentId}`);
      }
      const result = await graph.api(`${base}/messages/${args.messageId}/attachments`).select('id,name,contentType,size').get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mapped = (result.value ?? []).map((a: any) => ({ id: a.id, name: a.name, contentType: a.contentType, size: a.size }));
      // List mode is a Graph collection GET like every other list verb: surface
      // truncation off @odata.nextLink so a message with more attachments than
      // one page is never presented as the complete set ( invariant 2).
      const { items, count, limit, truncated } = toEnvelope(mapped, 100, graphCollectionHasMore(result));
      return { attachments: items, count, limit, truncated };
    },
  },
  {
    name: 'create_draft',
    description:
      'Create an email draft in the mailbox (does not send). Supports file attachments via the ' +
      'optional attachments array. If you are revising an existing draft, use update_message instead ' +
      'so the message ID stays stable and the user sees an in-place edit in Outlook.',
    inputSchema: {
      type: 'object',
      properties: { subject: prop('string', 'Subject'), body: prop('string', 'Body'), bodyType: prop('string', 'text or html'), to: { type: 'array', items: { type: 'string' }, description: 'Recipients' }, cc: { type: 'array', items: { type: 'string' }, description: 'CC' }, bcc: { type: 'array', items: { type: 'string' }, description: 'BCC' }, attachments: ATTACHMENTS_SCHEMA, from: FROM_PROP },
      required: ['subject', 'to'],
    },
    handler: async (args, session) => {
      // Drafts are written to the Drafts folder — check deny list by display name
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'mail', 'Drafts')) throw new Error('Access restricted by deny list — Drafts folder is blocked');
      const attachmentRequests = parseAttachments(args.attachments);
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const from = await resolveExplicitFrom(graph, '/me', args.from);
      const attachments = await resolveAttachments(graph, attachmentRequests, { session, userId: session.userId, tenantId, operation: 'create_draft', source: 'mcp' });
      const msg = { subject: args.subject, body: { contentType: args.bodyType === 'html' ? 'HTML' : 'Text', content: args.body ?? '' }, toRecipients: args.to.map((a: string) => ({ emailAddress: { address: a } })), ...(args.cc ? { ccRecipients: args.cc.map((a: string) => ({ emailAddress: { address: a } })) } : {}), ...(args.bcc ? { bccRecipients: args.bcc.map((a: string) => ({ emailAddress: { address: a } })) } : {}), ...(from ? { from: toGraphFrom(from) } : {}) };
      const result = await createDraftWithAttachments(graph, msg, attachments);
      return { id: result.id, subject: result.subject, webLink: result.webLink, ...(from ? { from } : {}), status: 'draft' };
    },
  },
  {
    name: 'reply_to_message',
    description:
      'Reply to an email — creates a draft reply (never sends) addressed to the original sender, with the ' +
      'quoted original body below your text and the threading headers set so it collapses into the ' +
      "recipient's existing conversation. Always use this instead of create_draft with a \"RE:\" subject: a " +
      'hand-built reply carries no quoted history and no In-Reply-To/References, so the recipient sees it as a ' +
      'brand-new conversation. Review the returned draft, then dispatch it with send_draft.',
    inputSchema: {
      type: 'object',
      properties: { ...REPLY_SCHEMA_PROPS },
      required: ['messageId', 'comment'],
    },
    handler: (args, session) => createReplyDraft(args, session, 'createReply'),
  },
  {
    name: 'reply_all_to_message',
    description:
      'Reply-all to an email — same as reply_to_message, but the draft is addressed to the original sender ' +
      'and every other recipient (To and CC). Creates a draft; never sends — dispatch it with send_draft. ' +
      'Check the returned `to`/`cc` first, since reply-all widens the audience.',
    inputSchema: {
      type: 'object',
      properties: { ...REPLY_SCHEMA_PROPS },
      required: ['messageId', 'comment'],
    },
    handler: (args, session) => createReplyDraft(args, session, 'createReplyAll'),
  },
  {
    name: 'forward_message',
    description:
      'Forward an email to new recipients — creates a draft (never sends) with the original message quoted ' +
      'below your text and the threading headers intact. Use this instead of pasting the original into ' +
      'create_draft, which loses both the attachments and the conversation link.',
    inputSchema: {
      type: 'object',
      properties: {
        ...REPLY_SCHEMA_PROPS,
        to: { type: 'array', items: { type: 'string' }, description: 'Recipients to forward to' },
      },
      required: ['messageId', 'to', 'comment'],
    },
    handler: (args, session) => {
      const to = args.to;
      if (!Array.isArray(to) || to.length === 0) throw new Error('forward_message requires at least one recipient in `to`');
      return createReplyDraft(args, session, 'createForward', {
        toRecipients: to.map((a: string) => ({ emailAddress: { address: a } })),
      });
    },
  },
  {
    name: 'update_message',
    description:
      'Update an existing email draft in place (subject, body, or recipients) — prefer this over delete_message + create_draft when iterating on a draft, so the message ID stays stable and the user sees an in-place edit in Outlook rather than a new draft replacing the old one. Only works on drafts; on a sent message returns an error pointing to delete_message + create_draft.',
    inputSchema: {
      type: 'object',
      properties: { messageId: prop('string', 'Draft message ID'), subject: prop('string', 'New subject (optional)'), body: prop('string', 'New body (optional)'), bodyType: prop('string', 'text or html'), to: { type: 'array', items: { type: 'string' }, description: 'Replacement To recipients (optional)' }, cc: { type: 'array', items: { type: 'string' }, description: 'Replacement CC recipients (optional)' }, bcc: { type: 'array', items: { type: 'string' }, description: 'Replacement BCC recipients (optional)' }, mailboxId: prop('string', 'Mailbox (default: me)'), from: prop('string', "Replacement From (optional). Must be one of the mailbox's own proxy addresses; anything else is rejected.") },
      required: ['messageId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      if (args.subject === undefined && args.body === undefined && args.to === undefined && args.cc === undefined && args.bcc === undefined && args.from === undefined) {
        throw new Error('At least one of subject, body, to, cc, bcc, from must be provided');
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const msgMeta: any = await graph.api(`${base}/messages/${args.messageId}`).select('isDraft,parentFolderId').get();
      if (msgMeta.parentFolderId) {
        const tenantId = getTenantIdFromSession(session);
        const folderName = await resolveMailFolderName(graph, msgMeta.parentFolderId, base);
        if (folderName && await isPathDenied(tenantId, session.userId, 'mail', folderName)) throw new Error('Access restricted by deny list');
      }
      if (msgMeta.isDraft === false) {
        throw new Error('This message is not a draft. Graph only allows updates to categories/flag/isRead on sent messages — to rewrite the body or recipients, delete the original (delete_message) and create a new draft (create_draft).');
      }
      const patch: Record<string, unknown> = {};
      if (args.subject !== undefined) patch.subject = args.subject;
      if (args.body !== undefined) patch.body = { contentType: args.bodyType === 'html' ? 'HTML' : 'Text', content: args.body };
      if (args.to !== undefined) patch.toRecipients = args.to.map((a: string) => ({ emailAddress: { address: a } }));
      if (args.cc !== undefined) patch.ccRecipients = args.cc.map((a: string) => ({ emailAddress: { address: a } }));
      if (args.bcc !== undefined) patch.bccRecipients = args.bcc.map((a: string) => ({ emailAddress: { address: a } }));
      const from = await resolveExplicitFrom(graph, base, args.from);
      if (from) patch.from = toGraphFrom(from);
      const result = await graph.api(`${base}/messages/${args.messageId}`).patch(patch);
      return { id: result.id, subject: result.subject, webLink: result.webLink, status: 'updated' };
    },
  },
  {
    name: 'send_mail',
    description:
      "Compose an email. Supports file attachments via the optional attachments array. Delivery " +
      "depends on the calling user's email output mode: in 'draft' mode " +
      "(the default for new users) the message is saved to Drafts for review instead of being sent, and " +
      "the response status is 'queued_as_draft' with a draftId/draftLink; in 'send' mode it is delivered " +
      "immediately and the status is 'sent'. Check the mode first with get_email_output_mode if you need to " +
      'tell the user what will happen.',
    inputSchema: {
      type: 'object',
      properties: { subject: prop('string', 'Subject'), body: prop('string', 'Body'), bodyType: prop('string', 'text or html'), to: { type: 'array', items: { type: 'string' }, description: 'Recipients' }, cc: { type: 'array', items: { type: 'string' }, description: 'CC' }, bcc: { type: 'array', items: { type: 'string' }, description: 'BCC' }, attachments: ATTACHMENTS_SCHEMA, from: FROM_PROP },
      required: ['subject', 'to'],
    },
    handler: async (args, session) => {
      const tenantId = getTenantIdFromSession(session);
      const attachmentRequests = parseAttachments(args.attachments);
      const message: Record<string, unknown> = {
        subject: args.subject,
        body: { contentType: args.bodyType === 'html' ? 'HTML' : 'Text', content: args.body ?? '' },
        toRecipients: args.to.map((a: string) => ({ emailAddress: { address: a } })),
        ...(args.cc ? { ccRecipients: args.cc.map((a: string) => ({ emailAddress: { address: a } })) } : {}),
        ...(args.bcc ? { bccRecipients: args.bcc.map((a: string) => ({ emailAddress: { address: a } })) } : {}),
      };

      // Per-user email output mode. New users default to 'draft', so AI-generated
      // email lands in Drafts for review rather than going out immediately. This is the same gate
      // the REST /api/mail/send path enforces — MCP callers must not be able to bypass it.
      const { emailOutputMode } = await getUserEmailSettings(tenantId, session.userId);
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const from = await resolveExplicitFrom(graph, '/me', args.from);
      if (from) message.from = toGraphFrom(from);
      // Drive-item attachments are fetched only after the target folder's deny
      // check below, so a send that is refused anyway never reads a file.
      const attachmentCtx = { session, userId: session.userId, tenantId, operation: 'send_mail', source: 'mcp' as const };

      if (emailOutputMode === 'draft') {
        // Saving to Drafts — check the Drafts folder deny list by display name
        if (await isPathDenied(tenantId, session.userId, 'mail', 'Drafts')) throw new Error('Access restricted by deny list — Drafts folder is blocked');
        const attachments = await resolveAttachments(graph, attachmentRequests, attachmentCtx);
        const result = await createDraftWithAttachments(graph, message, attachments);
        return { status: 'queued_as_draft', subject: args.subject, to: args.to, ...(from ? { from } : {}), draftId: result.id, draftLink: result.webLink };
      }

      // send mode — delivers immediately, saves to Sent Items; check that folder's deny list
      if (await isPathDenied(tenantId, session.userId, 'mail', 'Sent Items')) throw new Error('Access restricted by deny list — Sent Items folder is blocked');
      const attachments = await resolveAttachments(graph, attachmentRequests, attachmentCtx);
      await sendMailWithAttachments(graph, message, attachments, true);
      return { status: 'sent', subject: args.subject, to: args.to, ...(from ? { from } : {}) };
    },
  },
  {
    name: 'send_draft',
    description:
      'Send an existing draft — the dispatch step for a draft made by reply_to_message, ' +
      'reply_all_to_message, forward_message, or create_draft. Takes the draft ID and sends it as-is, ' +
      "preserving the threading headers a reply draft carries. Subject to the user's email output mode: in " +
      "'send' mode it delivers immediately; in 'draft' mode it refuses and returns the draft's webLink, " +
      'because that mode exists to put a human between composition and delivery. Only works on drafts.',
    inputSchema: {
      type: 'object',
      properties: { messageId: prop('string', 'Draft message ID to send'), mailboxId: prop('string', 'Mailbox (default: me)'), from: prop('string', "Set the draft's From before sending. Must be one of the mailbox's own proxy addresses; anything else is rejected. Omit to send with the From the draft already carries (a reply draft already defaults to the alias the original came in on).") },
      required: ['messageId'],
    },
    handler: async (args, session) => {
      const tenantId = getTenantIdFromSession(session);
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const draft: any = await graph.api(`${base}/messages/${args.messageId}`)
        .select('isDraft,parentFolderId,subject,from,toRecipients,ccRecipients,webLink,conversationId')
        .get();

      // Deny list on the folder the draft lives in, then on Sent Items, where Graph files the
      // sent copy — same pair of checks send_mail makes across its two modes.
      if (draft.parentFolderId) {
        const folderName = await resolveMailFolderName(graph, draft.parentFolderId, base);
        if (folderName && await isPathDenied(tenantId, session.userId, 'mail', folderName)) {
          throw new Error('Access restricted by deny list');
        }
      }
      if (await isPathDenied(tenantId, session.userId, 'mail', 'Sent Items')) {
        throw new Error('Access restricted by deny list — Sent Items folder is blocked');
      }

      if (draft.isDraft === false) {
        throw new Error('This message is not a draft — it has already been sent. send_draft only dispatches unsent drafts.');
      }

      // Email output mode gates this exactly as it gates send_mail. A draft that an
      // agent created moments ago has not been reviewed by anyone, so treating its mere existence
      // as the review would turn 'draft' mode — the default for every new user — into a no-op.
      const { emailOutputMode, enforced } = await getUserEmailSettings(tenantId, session.userId);
      if (emailOutputMode === 'draft') {
        // When an admin enforces draft mode the mode cannot be changed from here, so
        // the message must not suggest it: that suggestion is exactly what an injected prompt
        // would act on.
        const howToChange = enforced
          ? 'Draft mode is enforced by your administrator and cannot be changed by this tool.'
          : "To let tools dispatch drafts directly, switch the mode with set_email_output_mode('send').";
        throw new Error(
          "Your email output mode is 'draft', so drafts are sent by a person, not by this tool. " +
          `Open the draft and send it from Outlook: ${draft.webLink ?? '(no link available)'}. ` +
          howToChange
        );
      }

      const from = await resolveExplicitFrom(graph, base, args.from);
      if (from) await graph.api(`${base}/messages/${args.messageId}`).patch({ from: toGraphFrom(from) });

      await graph.api(`${base}/messages/${args.messageId}/send`).post({});
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const addresses = (list: any): string[] => (list ?? []).map((r: any) => r?.emailAddress?.address).filter(Boolean);
      return {
        status: 'sent',
        messageId: args.messageId,
        subject: draft.subject,
        from: from ?? draft.from?.emailAddress?.address ?? null,
        to: addresses(draft.toRecipients),
        cc: addresses(draft.ccRecipients),
        conversationId: draft.conversationId,
      };
    },
  },
  {
    name: 'get_email_output_mode',
    description:
      "Get the calling user's email output mode. Returns 'draft' (the default — send_mail saves to Drafts " +
      "for review) or 'send' (send_mail delivers immediately). Call this before composing an email so you " +
      'can tell the user whether it will be sent or saved as a draft. When `enforced` is true an administrator ' +
      "has pinned the mode to 'draft' (`enforcedBy` says whether tenant-wide or for this user) and " +
      'set_email_output_mode will refuse to change it.',
    inputSchema: { type: 'object', properties: {}, required: [] },
    handler: async (_args, session) => {
      const tenantId = getTenantIdFromSession(session);
      const { emailOutputMode, enforced, enforcedBy } = await getUserEmailSettings(tenantId, session.userId);
      return { emailOutputMode, enforced: enforced === true, enforcedBy: enforcedBy ?? null };
    },
  },
  {
    name: 'set_email_output_mode',
    description:
      "Update the calling user's email output mode. Pass 'draft' to require review in Drafts before " +
      "delivery (the default for new users) or 'send' to allow send_mail to deliver immediately. Only " +
      "changes the calling user's own setting. Refuses when an administrator enforces draft mode for the " +
      'tenant or for this user; only an administrator can lift that in the admin UI.',
    inputSchema: {
      type: 'object',
      properties: { emailOutputMode: prop('string', "'draft' (save to Drafts for review) or 'send' (deliver immediately)", { enum: ['draft', 'send'] }) },
      required: ['emailOutputMode'],
    },
    handler: async (args, session) => {
      const mode = args.emailOutputMode;
      if (mode !== 'draft' && mode !== 'send') throw new Error("emailOutputMode must be 'draft' or 'send'");
      const tenantId = getTenantIdFromSession(session);
      // Enforced draft mode: an agent-facing tool must not be able to lift the one control
      // that puts a person between composition and delivery. Refuse before any write, for either
      // value: even a 'draft' write while enforced would suggest the tool has a say in the matter.
      // The message carries the ENFORCED_MODE_MARKER so the dispatcher logs the refusal as denied.
      const { enforced, enforcedBy } = await getUserEmailSettings(tenantId, session.userId);
      if (enforced) {
        const scope = enforcedBy === 'tenant' ? 'for this tenant' : 'for your account';
        throw new Error(
          `${ENFORCED_MODE_MARKER} ${scope}, so set_email_output_mode cannot change it. ` +
          'Emails will be saved to Drafts for a person to review and send. Only an administrator can lift this in the admin UI.'
        );
      }
      await setUserEmailSettings(tenantId, session.userId, { emailOutputMode: mode as EmailOutputMode });
      return { status: 'updated', emailOutputMode: mode };
    },
  },
  {
    name: 'move_message',
    description: 'Move an email to a different folder',
    inputSchema: {
      type: 'object',
      properties: { messageId: prop('string', 'Message ID'), destinationFolderId: prop('string', 'Target folder ID'), mailboxId: prop('string', 'Mailbox (default: me)') },
      required: ['messageId', 'destinationFolderId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      // Check deny list for both source folder and destination folder
      const tenantId = getTenantIdFromSession(session);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const msgMeta: any = await graph.api(`${base}/messages/${args.messageId}`).select('parentFolderId').get();
      if (msgMeta.parentFolderId) {
        const srcName = await resolveMailFolderName(graph, msgMeta.parentFolderId, base);
        if (srcName && await isPathDenied(tenantId, session.userId, 'mail', srcName)) throw new Error('Access restricted by deny list');
      }
      const destName = await resolveMailFolderName(graph, args.destinationFolderId, base);
      if (destName && await isPathDenied(tenantId, session.userId, 'mail', destName)) throw new Error('Destination folder restricted by deny list');
      const result = await graph.api(`${base}/messages/${args.messageId}/move`).post({ destinationId: args.destinationFolderId });
      return { id: result.id, subject: result.subject, status: 'moved' };
    },
  },
  {
    name: 'delete_message',
    description: 'Delete an email message',
    inputSchema: {
      type: 'object',
      properties: { messageId: prop('string', 'Message ID'), mailboxId: prop('string', 'Mailbox (default: me)') },
      required: ['messageId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const msgMeta: any = await graph.api(`${base}/messages/${args.messageId}`).select('parentFolderId').get();
      if (msgMeta.parentFolderId) {
        const tenantId = getTenantIdFromSession(session);
        const folderName = await resolveMailFolderName(graph, msgMeta.parentFolderId, base);
        if (folderName && await isPathDenied(tenantId, session.userId, 'mail', folderName)) throw new Error('Access restricted by deny list');
      }
      await graph.api(`${base}/messages/${args.messageId}`).delete();
      return { status: 'deleted', messageId: args.messageId };
    },
  },
  // OneDrive
  {
    name: 'list_onedrive',
    description: 'List files and folders in personal OneDrive. Set foldersOnly=true to hide files.',
    inputSchema: {
      type: 'object',
      properties: { parentId: prop('string', 'Parent folder ID (default: root)'), foldersOnly: prop('boolean', 'Only return folders') },
      required: [],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const apiPath = args.parentId ? `/me/drive/items/${args.parentId}/children` : '/me/drive/root/children';
      let query = graph.api(apiPath).select('id,name,folder,file,size,parentReference,lastModifiedDateTime').top(200);
      if (args.foldersOnly) query = query.filter('folder ne null');
      const result = await query.get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const mapped = (result.value ?? []).map((item: any) => {
        const isFolder = item.folder !== undefined;
        return { id: item.id, name: item.name, path: `${item.parentReference?.path ?? ''}/${item.name}`, type: isFolder ? 'folder' : 'file', ...(isFolder ? { childCount: item.folder?.childCount ?? 0 } : { size: item.size, mimeType: item.file?.mimeType ?? null }), lastModifiedDateTime: item.lastModifiedDateTime };
      });
      const tenantId = getTenantIdFromSession(session);
      const allowed = await filterDeniedPaths(tenantId, session.userId, 'onedrive', mapped);
      const { items, count, limit, truncated } = toEnvelope(allowed, 200, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'read_onedrive_file',
    description: 'Fetch content of a file from personal OneDrive',
    inputSchema: { type: 'object', properties: { itemId: prop('string', 'File item ID') }, required: ['itemId'] },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const metadata: any = await graph.api(`/me/drive/items/${args.itemId}`).select('id,name,webUrl,size,file,parentReference,createdDateTime,lastModifiedDateTime').get();
      const filePath = `${metadata.parentReference?.path ?? ''}/${metadata.name}`;
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'onedrive', filePath)) throw new Error('Access restricted by deny list');
      if (!metadata.file) throw new Error('Item is not a file');
      const stream = await graph.api(`/me/drive/items/${args.itemId}/content`).getStream();
      const chunks: Buffer[] = [];
      for await (const chunk of stream) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as Uint8Array));
      const buf = Buffer.concat(chunks);
      const mime = metadata.file.mimeType ?? 'application/octet-stream';
      const isText = isTextMime(mime);
      return { id: metadata.id, name: metadata.name, mimeType: mime, size: metadata.size, content: isText ? buf.toString('utf-8') : buf.toString('base64'), encoding: isText ? 'utf-8' : 'base64' };
    },
  },
  {
    name: 'write_onedrive_file',
    description: "Create or overwrite a file in personal OneDrive",
    inputSchema: {
      type: 'object',
      properties: { path: prop('string', 'File path (e.g. Documents/notes.txt)'), content: prop('string', 'File content'), contentType: prop('string', 'MIME type'), encoding: prop('string', 'Content encoding: "utf-8" (default) or "base64" for binary files') },
      required: ['path', 'content'],
    },
    handler: async (args, session) => {
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'onedrive', args.path)) throw new Error('Access restricted by deny list');
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const buf = args.encoding === 'base64' ? Buffer.from(args.content, 'base64') : Buffer.from(args.content, 'utf-8');
      const result = await graph.api(`/me/drive/root:/${args.path.replace(/^\//, '')}:/content`).header('Content-Type', args.contentType ?? 'text/plain').put(buf);
      return { id: result.id, name: result.name, webUrl: result.webUrl, status: 'written' };
    },
  },
  {
    name: 'move_onedrive_item',
    description: 'Move or rename a file/folder in personal OneDrive',
    inputSchema: {
      type: 'object',
      properties: { itemId: prop('string', 'Item ID'), destinationFolderId: prop('string', 'Destination folder ID'), newName: prop('string', 'New name (optional)') },
      required: ['itemId', 'destinationFolderId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      // Check deny list for source item
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const meta: any = await graph.api(`/me/drive/items/${args.itemId}`).select('name,parentReference').get();
      const itemPath = `${meta.parentReference?.path ?? ''}/${meta.name}`;
      if (await isPathDenied(tenantId, session.userId, 'onedrive', itemPath)) throw new Error('Access restricted by deny list');
      // Check deny list for destination folder
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const destMeta: any = await graph.api(`/me/drive/items/${args.destinationFolderId}`).select('name,parentReference').get();
      const destPath = `${destMeta.parentReference?.path ?? ''}/${destMeta.name}`;
      if (await isPathDenied(tenantId, session.userId, 'onedrive', destPath)) throw new Error('Destination folder restricted by deny list');
      const patch: Record<string, unknown> = { parentReference: { id: args.destinationFolderId } };
      if (args.newName) patch.name = args.newName;
      const result = await graph.api(`/me/drive/items/${args.itemId}`).patch(patch);
      return { id: result.id, name: result.name, status: 'moved' };
    },
  },
  {
    name: 'create_onedrive_folder',
    description: 'Create a new folder in personal OneDrive',
    inputSchema: {
      type: 'object',
      properties: { name: prop('string', 'Folder name'), parentId: prop('string', 'Parent folder ID (default: root)') },
      required: ['name'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Check deny list for the new folder path
      if (args.parentId) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const parentMeta: any = await graph.api(`/me/drive/items/${args.parentId}`).select('name,parentReference').get();
        const parentPath = `${parentMeta.parentReference?.path ?? ''}/${parentMeta.name}`;
        const tenantId = getTenantIdFromSession(session);
        if (await isPathDenied(tenantId, session.userId, 'onedrive', `${parentPath}/${args.name}`)) throw new Error('Access restricted by deny list');
      }
      const apiPath = args.parentId ? `/me/drive/items/${args.parentId}/children` : '/me/drive/root/children';
      const result = await graph.api(apiPath).post({ name: args.name, folder: {}, '@microsoft.graph.conflictBehavior': 'fail' });
      return { id: result.id, name: result.name, webUrl: result.webUrl, status: 'created' };
    },
  },
  {
    name: 'delete_onedrive_item',
    description: 'Delete a file or folder from personal OneDrive',
    inputSchema: { type: 'object', properties: { itemId: prop('string', 'Item ID') }, required: ['itemId'] },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const meta: any = await graph.api(`/me/drive/items/${args.itemId}`).select('name,parentReference').get();
      const itemPath = `${meta.parentReference?.path ?? ''}/${meta.name}`;
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'onedrive', itemPath)) throw new Error('Access restricted by deny list');
      await graph.api(`/me/drive/items/${args.itemId}`).delete();
      return { status: 'deleted', itemId: args.itemId };
    },
  },
  // Calendar
  {
    name: 'list_calendars',
    description: "List the user's calendars",
    inputSchema: { type: 'object', properties: {}, required: [] },
    handler: async (_args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const result = await graph.api('/me/calendars').select('id,name,color,isDefaultCalendar').top(100).get();
      // Exclude deny-listed calendars (by ID or name) so a blocked calendar is
      // never even enumerated — mirrors the HTTP list_calendars route.
      const tenantId = getTenantIdFromSession(session);
      const allowed = await filterDeniedPaths(tenantId, session.userId, 'calendar', result.value ?? []);
      const { items, count, limit, truncated } = toEnvelope(allowed, 100, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'create_calendar',
    description:
      'Create a new calendar in the signed-in user\'s mailbox (POST /me/calendars). Use this to set up a ' +
      'separate calendar — e.g. a "Private" calendar for items that should not be visible on a shared/default ' +
      'calendar feed — that create_event can then target via calendarId. Refused if a calendar with the given ' +
      'name is on the deny list, or the calendar service is in read-only mode.',
    inputSchema: {
      type: 'object',
      properties: {
        name: prop('string', 'Display name for the new calendar (e.g. "Private").'),
        color: prop('string', 'Optional calendar color preset.', {
          enum: ['auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightYellow', 'lightTeal', 'lightPink', 'lightBrown', 'lightRed'],
        }),
      },
      required: ['name'],
    },
    handler: async (args, session) => {
      if (!args.name || typeof args.name !== 'string' || !args.name.trim()) {
        throw new Error('name is required and must be a non-empty string');
      }
      // Normalize once and use the trimmed name for BOTH the deny-list check and
      // the Graph create body. isPathDenied canonicalizes slash/case but not
      // leading/trailing whitespace, so a padded name like " Executive " would
      // otherwise slip past a deny entry for "Executive" and be created verbatim.
      const name = args.name.trim();
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      // Deny-list is keyed by calendar name here (there is no ID yet): an admin
      // who blocks a calendar by name should not be able to have one recreated
      // under the same name through this tool.
      if (await isPathDenied(tenantId, session.userId, 'calendar', name)) {
        throw new Error('Access restricted by deny list');
      }
      const body: Record<string, unknown> = { name };
      if (args.color) body.color = args.color;
      const result = await graph.api('/me/calendars').post(body);
      return {
        id: result.id, name: result.name, color: result.color ?? null,
        isDefaultCalendar: result.isDefaultCalendar ?? false, status: 'created',
      };
    },
  },
  {
    name: 'list_events',
    description:
      'List calendar events. When a date range is given (startDateTime and/or endDateTime), this queries Graph ' +
      'calendarView, which EXPANDS recurring series into their concrete instances on those days — so it correctly ' +
      'answers "what is on this specific day". Times are returned in timeZone (default America/Los_Angeles). With no ' +
      'range it lists raw calendar events (recurring series appear as their master, not expanded).',
    inputSchema: {
      type: 'object',
      properties: {
        calendarId: prop('string', 'Calendar ID'),
        startDateTime: prop('string', 'Window start, ISO 8601 or a bare date (YYYY-MM-DD). Defaults the end 7 days out if omitted while end is set.'),
        endDateTime: prop('string', 'Window end, ISO 8601 or a bare date. Defaults the start 7 days back if omitted while start is set.'),
        timeZone: prop('string', 'IANA time zone the window and returned times are interpreted in (default: "America/Los_Angeles")'),
        maxResults: prop('number', 'Max results'),
      },
      required: [],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Deny-list enforcement (by ID and name) for the explicit or default calendar.
      const listViolation = await checkCalendarAccess(graph, getTenantIdFromSession(session), session.userId, args.calendarId);
      if (listViolation) throw new Error(listViolation.error);
      const tz = args.timeZone ?? 'America/Los_Angeles';
      const top = resolveMaxResults(args.maxResults, 25, 100);
      const select = 'id,subject,start,end,location,organizer,attendees,isAllDay,webLink';

      // With a date range, use calendarView so recurring series are expanded into
      // concrete instances on those days (GET /events returns series masters whose
      // start is the ORIGINAL occurrence — often years in the past — which is
      // useless for "what is on 2026-09-08"). See.
      const window = resolveWindow(args.startDateTime, args.endDateTime);
      let result;
      if (window) {
        const base = args.calendarId ? `/me/calendars/${encodeGraphId(args.calendarId, 'calendarId')}/calendarView` : '/me/calendarView';
        const qs = `startDateTime=${encodeURIComponent(window.start)}&endDateTime=${encodeURIComponent(window.end)}`;
        result = await graph
          .api(`${base}?${qs}`)
          // Interpret the naive-local window and return start/end in this zone.
          .header('Prefer', `outlook.timezone="${tz}"`)
          .select(select)
          .top(top)
          .orderby('start/dateTime')
          .get();
      } else {
        const apiPath = args.calendarId ? `/me/calendars/${encodeGraphId(args.calendarId, 'calendarId')}/events` : '/me/events';
        result = await graph.api(apiPath).select(select).top(top).orderby('start/dateTime').get();
      }
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const events = (result.value ?? []).map((e: any) => ({ id: e.id, subject: e.subject, start: e.start, end: e.end, location: e.location?.displayName, organizer: e.organizer?.emailAddress?.address, attendees: (e.attendees ?? []).map((a: any) => a.emailAddress?.address), isAllDay: e.isAllDay, webLink: e.webLink }));
      const { items, count, limit, truncated } = toEnvelope(events, top, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'get_event',
    description: 'Fetch full details of a single calendar event (body, attendees, location)',
    inputSchema: {
      type: 'object',
      properties: { eventId: prop('string', 'Event ID'), calendarId: prop('string', 'Calendar ID (default: default calendar)') },
      required: ['eventId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Deny-list enforcement (by ID and name) for the explicit or default calendar.
      const getViolation = await checkCalendarAccess(graph, getTenantIdFromSession(session), session.userId, args.calendarId);
      if (getViolation) throw new Error(getViolation.error);
      const encEventId = encodeGraphId(args.eventId, 'eventId');
      const apiPath = args.calendarId
        ? `/me/calendars/${encodeGraphId(args.calendarId, 'calendarId')}/events/${encEventId}`
        : `/me/events/${encEventId}`;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const e: any = await graph.api(apiPath).select('id,subject,start,end,location,organizer,attendees,isAllDay,body,webLink').get();
      return {
        id: e.id, subject: e.subject, start: e.start, end: e.end,
        location: e.location?.displayName ?? null,
        organizer: e.organizer?.emailAddress?.address ?? null,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        attendees: (e.attendees ?? []).map((a: any) => a.emailAddress?.address),
        isAllDay: e.isAllDay ?? false, body: e.body, webLink: e.webLink,
      };
    },
  },
  {
    name: 'create_event',
    description:
      'Create a calendar event. When attendees are involved and you have not already confirmed everyone is free, ' +
      'call find_meeting_times or get_schedule first to pick a slot that works, and list_rooms to choose a real ' +
      'conference room for the location. Set isOnlineMeeting=true to auto-generate a Microsoft Teams join link — ' +
      'the response then includes onlineMeetingUrl (the Teams "Join" URL) so you never have to ask "what\'s the link?".',
    inputSchema: {
      type: 'object',
      properties: { subject: prop('string', 'Title'), start: prop('string', 'Start (ISO)'), end: prop('string', 'End (ISO)'), timeZone: prop('string', 'IANA or Windows time zone for start/end (default: the mailbox\u2019s configured time zone)'), location: prop('string', 'Location'), body: prop('string', 'Description'), bodyType: prop('string', 'text or html'), attendees: { type: 'array', items: { type: 'string' }, description: 'Attendee emails' }, isAllDay: prop('boolean', 'All-day'), showAs: prop('string', 'Free/busy status shown to others (default: Graph decides \u2014 busy for timed, free for all-day)', { enum: ['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'] }), isOnlineMeeting: prop('boolean', 'Set true to attach a Teams online meeting (returns onlineMeetingUrl)'), onlineMeetingProvider: prop('string', 'Online meeting provider (default: teamsForBusiness)', { enum: ['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer'] }), calendarId: prop('string', 'Calendar ID') },
      required: ['subject', 'start', 'end'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Deny-list enforcement (by ID and name) for the explicit or default calendar.
      const createViolation = await checkCalendarAccess(graph, getTenantIdFromSession(session), session.userId, args.calendarId);
      if (createViolation) throw new Error(createViolation.error);
      // Honor an explicit timeZone; otherwise default to the mailbox's own
      // configured zone rather than a hardcoded one. Fails loud if the
      // zone can't be resolved — never guesses.
      const tz = args.timeZone ?? await resolveMailboxTimeZone(graph);
      const event: Record<string, unknown> = { subject: args.subject, start: { dateTime: args.start, timeZone: tz }, end: { dateTime: args.end, timeZone: tz }, isAllDay: args.isAllDay ?? false };
      if (args.location) event.location = { displayName: args.location };
      if (args.body) event.body = { contentType: args.bodyType === 'html' ? 'HTML' : 'Text', content: args.body };
      if (args.attendees) event.attendees = args.attendees.map((a: string) => ({ emailAddress: { address: a }, type: 'required' }));
      // Free/busy status. Only set it when supplied so an omitted showAs leaves Graph's default.
      if (args.showAs) event.showAs = args.showAs;
      // Attach a Teams (or Skype) online meeting so the invite carries a join URL.
      if (args.isOnlineMeeting) {
        event.isOnlineMeeting = true;
        event.onlineMeetingProvider = args.onlineMeetingProvider ?? 'teamsForBusiness';
      }
      const apiPath = args.calendarId ? `/me/calendars/${encodeGraphId(args.calendarId, 'calendarId')}/events` : '/me/events';
      const result = await graph.api(apiPath).post(event);
      return {
        id: result.id, subject: result.subject, start: result.start, end: result.end,
        webLink: result.webLink,
        isOnlineMeeting: result.isOnlineMeeting ?? false,
        // Surface the Teams "Join" URL directly so the caller doesn't have to re-fetch the event.
        onlineMeetingUrl: result.onlineMeeting?.joinUrl ?? null,
        status: 'created',
      };
    },
  },
  {
    name: 'update_event',
    description: 'Update an existing calendar event',
    inputSchema: {
      type: 'object',
      properties: { eventId: prop('string', 'Event ID'), subject: prop('string', 'New subject'), start: prop('string', 'New start'), end: prop('string', 'New end'), timeZone: prop('string', 'IANA or Windows time zone for start/end (default: the mailbox\u2019s configured time zone)'), location: prop('string', 'Location'), body: prop('string', 'Description'), bodyType: prop('string', 'text or html'), attendees: { type: 'array', items: { type: 'string' }, description: 'Attendees' }, isAllDay: prop('boolean', 'All-day'), showAs: prop('string', 'Free/busy status shown to others', { enum: ['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'] }), calendarId: prop('string', 'Calendar ID') },
      required: ['eventId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Deny-list enforcement (by ID and name) for the explicit or default calendar.
      const updateViolation = await checkCalendarAccess(graph, getTenantIdFromSession(session), session.userId, args.calendarId);
      if (updateViolation) throw new Error(updateViolation.error);
      // Resolve a time zone only when the caller is actually changing start/end.
      // Honor an explicit timeZone, else default to the mailbox's own zone
      //; never guess a hardcoded one, and never make an extra Graph
      // call to fetch it when no time is being rescheduled.
      const tz = (args.start !== undefined || args.end !== undefined)
        ? (args.timeZone ?? await resolveMailboxTimeZone(graph))
        : args.timeZone;
      const patch: Record<string, unknown> = {};
      if (args.subject !== undefined) patch.subject = args.subject;
      if (args.isAllDay !== undefined) patch.isAllDay = args.isAllDay;
      if (args.start !== undefined) patch.start = { dateTime: args.start, timeZone: tz };
      if (args.end !== undefined) patch.end = { dateTime: args.end, timeZone: tz };
      if (args.location !== undefined) patch.location = { displayName: args.location };
      if (args.body !== undefined) patch.body = { contentType: args.bodyType === 'html' ? 'HTML' : 'Text', content: args.body };
      if (args.attendees !== undefined) patch.attendees = args.attendees.map((a: string) => ({ emailAddress: { address: a }, type: 'required' }));
      // Free/busy status. Only patch it when supplied so an omitted showAs leaves the current status untouched.
      if (args.showAs !== undefined) patch.showAs = args.showAs;
      const encUpdEventId = encodeGraphId(args.eventId, 'eventId');
      const apiPath = args.calendarId
        ? `/me/calendars/${encodeGraphId(args.calendarId, 'calendarId')}/events/${encUpdEventId}`
        : `/me/events/${encUpdEventId}`;
      const result = await graph.api(apiPath).patch(patch);
      return { id: result.id, subject: result.subject, status: 'updated' };
    },
  },
  {
    name: 'delete_event',
    description: 'Delete a calendar event',
    inputSchema: {
      type: 'object',
      properties: { eventId: prop('string', 'Event ID'), calendarId: prop('string', 'Calendar ID') },
      required: ['eventId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Deny-list enforcement (by ID and name) for the explicit or default calendar.
      const deleteViolation = await checkCalendarAccess(graph, getTenantIdFromSession(session), session.userId, args.calendarId);
      if (deleteViolation) throw new Error(deleteViolation.error);
      const encDelEventId = encodeGraphId(args.eventId, 'eventId');
      const apiPath = args.calendarId
        ? `/me/calendars/${encodeGraphId(args.calendarId, 'calendarId')}/events/${encDelEventId}`
        : `/me/events/${encDelEventId}`;
      await graph.api(apiPath).delete();
      return { status: 'deleted', eventId: args.eventId };
    },
  },
  {
    name: 'move_event',
    description:
      'Move a calendar event from one calendar to another. IMPORTANT: Microsoft Graph has no native move for ' +
      'events (unlike mail), so this is implemented as COPY-THEN-DELETE — it is NOT atomic. The new event gets a ' +
      'new id and webLink; the original is deleted only after the copy is confirmed created (never the other way ' +
      'round). It preserves subject, body, start/end/timeZone, all-day flag, location, categories, sensitivity, ' +
      'showAs, importance, reminder settings, recurrence (the series master is copied as a series, not expanded ' +
      'into single events), and file attachments. SIDE EFFECTS: if the event has attendees, recreating it sends ' +
      'FRESH invitations to everyone and deleting the original sends cancellations — real outbound mail to real ' +
      'people. For that reason a move is REFUSED when the event has attendees unless you pass force=true, and is ' +
      'always refused when you are not the organizer or the event is a single occurrence of a recurring series ' +
      '(move the series master instead). Returns the new event id.',
    inputSchema: {
      type: 'object',
      properties: {
        eventId: prop('string', 'ID of the event to move (from list_events / get_event).'),
        targetCalendarId: prop('string', 'Destination calendar ID (from list_calendars / create_calendar).'),
        calendarId: prop('string', 'Source calendar ID. Omit to move from the default calendar.'),
        force: prop('boolean', 'Required (true) to move an event that has attendees — acknowledges that fresh invites and a cancellation will be sent. Default false.'),
      },
      required: ['eventId', 'targetCalendarId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      const userId = session.userId;
      // Same-calendar no-op guard. Resolve the EFFECTIVE source calendar ID
      // first: when calendarId is omitted the source is the user's default
      // calendar, so targetCalendarId must be compared against the resolved
      // default ID — not skipped. Otherwise move_event({ eventId,
      // targetCalendarId: <default calendar id> }) recreates the event in the
      // same calendar and deletes the original: a no-op that churns the event
      // id/webLink and, with force=true on an attendee event, fires fresh
      // invites plus a cancellation for nothing. The explicit case resolves to
      // args.calendarId with no Graph call; the default case reuses the cached
      // resolveDefaultCalendarId that checkCalendarAccess also consults.
      const effectiveSourceId = args.calendarId || (await resolveDefaultCalendarId(graph, userId));
      if (effectiveSourceId && effectiveSourceId === args.targetCalendarId) {
        throw new Error('Source and target calendars are the same — nothing to move');
      }
      // Deny-list enforcement on BOTH ends: the source calendar (explicit or
      // default) and the target calendar. A blocked calendar on either side
      // refuses the move before any data is touched.
      const sourceViolation = await checkCalendarAccess(graph, tenantId, userId, args.calendarId);
      if (sourceViolation) throw new Error(sourceViolation.error);
      const targetViolation = await checkCalendarAccess(graph, tenantId, userId, args.targetCalendarId);
      if (targetViolation) throw new Error(targetViolation.error);

      const encMoveEventId = encodeGraphId(args.eventId, 'eventId');
      const encTargetCalId = encodeGraphId(args.targetCalendarId, 'targetCalendarId');
      const sourcePath = args.calendarId
        ? `/me/calendars/${encodeGraphId(args.calendarId, 'calendarId')}/events/${encMoveEventId}`
        : `/me/events/${encMoveEventId}`;
      // Pull every field a faithful copy needs — create_event alone only carries
      // subject/start/end/body, which would silently drop the rest.
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let src: any;
      try {
        src = await graph
          .api(sourcePath)
          .select('id,subject,body,start,end,location,locations,categories,sensitivity,showAs,importance,isAllDay,isReminderOn,reminderMinutesBeforeStart,recurrence,attendees,isOrganizer,type,seriesMasterId,hasAttachments')
          .get();
      } catch (err: unknown) {
        if (isGraphNotFound(err)) throw new Error(`No calendar event found with ID "${args.eventId}"`);
        throw err;
      }

      // A single occurrence / exception of a series would be recreated as a lone
      // event, silently severing it from the series. Refuse and point the caller
      // at the series master.
      if (src.type === 'occurrence' || src.type === 'exception' || src.seriesMasterId) {
        throw new Error(
          'This event is a single occurrence of a recurring series. Move the series master instead ' +
          '(get_event on the series master, then move_event with its id) so the whole series moves as a unit.',
        );
      }
      // Moving an event you do not organize has the same blast radius as moving
      // someone else\'s meeting — refuse outright (no force override).
      if (src.isOrganizer === false) {
        throw new Error(
          'You are not the organizer of this event. Moving an invite you received would recreate the meeting ' +
          'and notify everyone — refused.',
        );
      }
      const attendees = Array.isArray(src.attendees) ? src.attendees : [];
      if (attendees.length > 0 && !args.force) {
        throw new Error(
          `This event has ${attendees.length} attendee(s). Moving it will send fresh invitations to all of them ` +
          'and a cancellation for the original — real outbound mail. Re-run with force=true to proceed.',
        );
      }

      // Build the copy. Only set fields Graph will accept on create; carry the
      // full location/recurrence objects verbatim so nothing is lossily flattened.
      const copy: Record<string, unknown> = {
        subject: src.subject,
        start: src.start,
        end: src.end,
        isAllDay: src.isAllDay ?? false,
      };
      if (src.body) copy.body = { contentType: src.body.contentType, content: src.body.content };
      if (src.location) copy.location = src.location;
      if (Array.isArray(src.locations) && src.locations.length > 0) copy.locations = src.locations;
      if (Array.isArray(src.categories) && src.categories.length > 0) copy.categories = src.categories;
      if (src.sensitivity) copy.sensitivity = src.sensitivity;
      if (src.showAs) copy.showAs = src.showAs;
      if (src.importance) copy.importance = src.importance;
      if (src.isReminderOn !== undefined) copy.isReminderOn = src.isReminderOn;
      if (src.reminderMinutesBeforeStart !== undefined) copy.reminderMinutesBeforeStart = src.reminderMinutesBeforeStart;
      if (src.recurrence) copy.recurrence = src.recurrence;
      if (attendees.length > 0) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        copy.attendees = attendees.map((a: any) => ({
          emailAddress: a.emailAddress,
          type: a.type ?? 'required',
        }));
      }

      // Create in the target FIRST and confirm it landed. Never delete the
      // source before we hold a new event id.
      const created = await graph.api(`/me/calendars/${encTargetCalId}/events`).post(copy);
      if (!created?.id) {
        throw new Error('Move aborted: the target calendar did not return a created event. The original is untouched.');
      }

      // Copy file attachments. Best-effort: item/reference attachments can\'t be
      // re-posted as bytes, so they\'re reported back rather than silently lost.
      const attachmentWarnings: string[] = [];
      if (src.hasAttachments) {
        try {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const atts: any = await graph.api(`${sourcePath}/attachments`).get();
          for (const att of atts.value ?? []) {
            if (att['@odata.type'] === '#microsoft.graph.fileAttachment' && att.contentBytes) {
              await graph.api(`/me/calendars/${encTargetCalId}/events/${encodeGraphId(created.id, 'created.id')}/attachments`).post({
                '@odata.type': '#microsoft.graph.fileAttachment',
                name: att.name,
                contentType: att.contentType,
                contentBytes: att.contentBytes,
                isInline: att.isInline ?? false,
              });
            } else {
              attachmentWarnings.push(`Skipped non-file attachment "${att.name ?? att.id}" (${att['@odata.type']})`);
            }
          }
        } catch (attErr: unknown) {
          attachmentWarnings.push(
            `Failed to copy attachments: ${attErr instanceof Error ? attErr.message : 'unknown error'}`,
          );
        }
      }

      // Source copied successfully — now delete the original.
      await graph.api(sourcePath).delete();

      return {
        status: 'moved',
        eventId: created.id,
        sourceEventId: args.eventId,
        targetCalendarId: args.targetCalendarId,
        webLink: created.webLink ?? null,
        attendeesNotified: attendees.length > 0,
        warnings: attachmentWarnings,
      };
    },
  },
  {
    name: 'respond_to_event',
    description:
      'Respond to a calendar event invitation — accept, tentatively accept, or decline it. Takes either the event ID ' +
      '(from list_events) or the meeting-invite message ID from the inbox (from search_mail / read_message); a message ' +
      'ID is resolved to its underlying event automatically. Set sendResponse=false to update the RSVP without ' +
      'notifying the organizer.',
    inputSchema: {
      type: 'object',
      properties: {
        messageOrEventId: prop('string', 'Event ID (from list_events) or meeting-invite message ID (from search_mail)'),
        response: prop('string', 'How to respond: accept, tentative, or decline', { enum: ['accept', 'tentative', 'decline'] }),
        comment: prop('string', 'Optional note sent to the organizer with the response'),
        sendResponse: prop('boolean', 'Send the response to the organizer (default true)'),
        mailboxId: prop('string', 'Mailbox (default: me)'),
      },
      required: ['messageOrEventId', 'response'],
    },
    handler: async (args, session) => {
      const responseChoice = EVENT_RESPONSE_ACTIONS[args.response];
      if (!responseChoice) throw new Error(`Invalid response "${args.response}" — use accept, tentative, or decline`);
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      const base = args.mailboxId && args.mailboxId !== 'me' ? `/users/${args.mailboxId}` : '/me';
      // Deny-list enforcement (by ID and name). Invite responses act on the
      // mailbox's default calendar: checkCalendarAccess covers the /me case; an
      // explicit mailbox is checked against ITS default calendar directly.
      if (base === '/me') {
        const respondViolation = await checkCalendarAccess(graph, tenantId, session.userId, undefined);
        if (respondViolation) throw new Error(respondViolation.error);
      } else {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const cal: any = await graph.api(`${base}/calendar`).select('id,name').get();
        if (
          (cal?.id && (await isPathDenied(tenantId, session.userId, 'calendar', cal.id))) ||
          (cal?.name && (await isPathDenied(tenantId, session.userId, 'calendar', cal.name)))
        ) {
          throw new Error('Access restricted by deny list');
        }
      }
      // Resolve the target event: try the ID as an event first; when Graph says
      // no such event exists, fall back to treating it as a meeting-invite
      // message and follow the eventMessage → event link.
      let eventId = args.messageOrEventId;
      try {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const ev: any = await graph.api(`${base}/events/${encodeGraphId(eventId, 'messageOrEventId')}`).select('id').get();
        eventId = ev?.id ?? eventId;
      } catch (err: unknown) {
        if (!isGraphNotFound(err)) throw err;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        let msg: any;
        try {
          msg = await graph
            .api(`${base}/messages/${encodeGraphId(args.messageOrEventId, 'messageOrEventId')}`)
            .select('id')
            .expand('microsoft.graph.eventMessage/event($select=id)')
            .get();
        } catch (msgErr: unknown) {
          if (isGraphNotFound(msgErr)) {
            throw new Error(`No calendar event or message found with ID "${args.messageOrEventId}"`);
          }
          throw msgErr;
        }
        if (!msg?.event?.id) {
          throw new Error('The message is not a meeting invitation — there is no linked calendar event to respond to');
        }
        eventId = msg.event.id;
      }
      const sendResponse = args.sendResponse ?? true;
      const body: Record<string, unknown> = { sendResponse };
      if (args.comment) body.comment = args.comment;
      await graph.api(`${base}/events/${encodeGraphId(eventId, 'eventId')}/${responseChoice.action}`).post(body);
      return { status: responseChoice.status, eventId, sendResponse };
    },
  },
  {
    name: 'get_schedule',
    description:
      'Look up free/busy availability for one or more people over a time window (Graph POST /me/calendar/getSchedule). ' +
      'Use this before scheduling a meeting with other attendees so you propose a slot everyone is actually free for, ' +
      'instead of guessing. Returns, per person, a coarse availabilityView string and the list of busy blocks plus ' +
      'their working hours. Only works for users whose free/busy the caller is permitted to see in the tenant.',
    inputSchema: {
      type: 'object',
      properties: {
        schedules: { type: 'array', items: { type: 'string' }, description: 'Email addresses (SMTP) of the people/rooms to check availability for' },
        startDateTime: prop('string', 'Window start, ISO 8601 (e.g. "2026-07-06T09:00:00")'),
        endDateTime: prop('string', 'Window end, ISO 8601'),
        timeZone: prop('string', 'IANA time zone for the window and results (default: "America/Los_Angeles")'),
        availabilityViewInterval: prop('number', 'Granularity of the availabilityView string, in minutes (default 30)'),
      },
      required: ['schedules', 'startDateTime', 'endDateTime'],
    },
    handler: async (args, session) => {
      if (!Array.isArray(args.schedules) || args.schedules.length === 0) {
        throw new Error('schedules must be a non-empty array of email addresses');
      }
      if (!args.startDateTime || !args.endDateTime) {
        // A missing bound leaves dateTime undefined, which Graph reports as the
        // opaque "FreeBusyViewOptions.TimeWindow is invalid" — fail loud instead.
        throw new Error('get_schedule requires both startDateTime and endDateTime');
      }
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tz = args.timeZone ?? 'America/Los_Angeles';
      // Normalize the window: expand bare dates and strip any absolute-time
      // marker that would contradict the timeZone field.
      const startTime = normalizeGraphDateTime(args.startDateTime, 'start');
      const endTime = normalizeGraphDateTime(args.endDateTime, 'end');
      // Graph requires availabilityViewInterval in [5, 1440] minutes.
      let interval = Math.round(Number(args.availabilityViewInterval ?? 30));
      if (!Number.isFinite(interval) || interval < 5) interval = 5;
      if (interval > 1440) interval = 1440;
      const body = {
        schedules: args.schedules,
        startTime: { dateTime: startTime, timeZone: tz },
        endTime: { dateTime: endTime, timeZone: tz },
        availabilityViewInterval: interval,
      };
      // Mirror find_meeting_times: the Prefer hint makes Graph honor the named
      // time zone for the window and the returned free/busy blocks.
      const result = await graph.api('/me/calendar/getSchedule').header('Prefer', `outlook.timezone="${tz}"`).post(body);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (result.value ?? []).map((s: any) => ({
        scheduleId: s.scheduleId,
        availabilityView: s.availabilityView,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        busy: (s.scheduleItems ?? []).map((i: any) => ({ status: i.status, start: i.start, end: i.end, subject: i.subject })),
        workingHours: s.workingHours ?? null,
        error: s.error?.message ?? null,
      }));
    },
  },
  {
    name: 'find_meeting_times',
    description:
      'Ask Microsoft Graph to suggest concrete meeting times that work for a set of attendees (POST /me/findMeetingTimes). ' +
      'Honors each attendee\'s working hours and time zone from their MailboxSettings. Prefer this over get_schedule when ' +
      'you want ranked, ready-to-book slots rather than raw busy blocks. Returns suggestions with a confidence score and ' +
      'the availability of each attendee for that slot.',
    inputSchema: {
      type: 'object',
      properties: {
        attendees: { type: 'array', items: { type: 'string' }, description: 'Required attendee email addresses' },
        meetingDurationMinutes: prop('number', 'Desired meeting length in minutes (default 30)'),
        startDateTime: prop('string', 'Earliest acceptable start, ISO 8601 (optional — defaults to Graph\'s window)'),
        endDateTime: prop('string', 'Latest acceptable end, ISO 8601 (optional)'),
        timeZone: prop('string', 'IANA time zone for the time window (default: "America/Los_Angeles")'),
        maxCandidates: prop('number', 'Maximum number of suggestions to return (default 10)'),
      },
      required: ['attendees'],
    },
    handler: async (args, session) => {
      if (!Array.isArray(args.attendees) || args.attendees.length === 0) {
        throw new Error('attendees must be a non-empty array of email addresses');
      }
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tz = args.timeZone ?? 'America/Los_Angeles';
      const durationMin = args.meetingDurationMinutes ?? 30;
      // ISO 8601 duration, e.g. PT30M
      const body: Record<string, unknown> = {
        attendees: args.attendees.map((a: string) => ({ type: 'required', emailAddress: { address: a } })),
        meetingDuration: `PT${durationMin}M`,
        maxCandidates: args.maxCandidates ?? 10,
        isOrganizerOptional: false,
        returnSuggestionReasons: true,
        minimumAttendeePercentage: 100,
      };
      // Constrain to the requested window whenever EITHER bound is given — with
      // only one bound the other is defaulted 7 days out — so the caller's dates
      // are honored instead of silently falling back to Graph's default (today).
      // Bare dates are expanded and absolute-time markers stripped.
      const window = resolveWindow(args.startDateTime, args.endDateTime);
      if (window) {
        body.timeConstraint = {
          activityDomain: 'work',
          timeSlots: [{ start: { dateTime: window.start, timeZone: tz }, end: { dateTime: window.end, timeZone: tz } }],
        };
      }
      // findMeetingTimes requires a hint header to honor time zone preferences.
      const result = await graph.api('/me/findMeetingTimes').header('Prefer', `outlook.timezone="${tz}"`).post(body);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const suggestions = (result.meetingTimeSuggestions ?? []).map((s: any) => ({
        confidence: s.confidence,
        meetingTimeSlot: s.meetingTimeSlot,
        suggestionReason: s.suggestionReason,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        attendeeAvailability: (s.attendeeAvailability ?? []).map((a: any) => ({ email: a.attendee?.emailAddress?.address, availability: a.availability })),
      }));
      return { emptySuggestionsReason: result.emptySuggestionsReason ?? null, suggestions };
    },
  },
  {
    name: 'list_rooms',
    description:
      'List the conference room mailboxes (places) in the tenant (Graph GET /places/microsoft.graph.room). Use this to ' +
      'pick a real, bookable room for a meeting location instead of leaving it blank — the returned emailAddress can be ' +
      'added as an attendee to book the room, and the displayName/address can be used as the event location.',
    inputSchema: { type: 'object', properties: { maxResults: prop('number', 'Maximum rooms to return (default 100)') }, required: [] },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const maxResults = resolveMaxResults(args.maxResults, 100, 200);
      const result = await graph.api('/places/microsoft.graph.room').top(maxResults).get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const rooms = (result.value ?? []).map((r: any) => ({
        id: r.id,
        displayName: r.displayName,
        emailAddress: r.emailAddress,
        building: r.building ?? null,
        floorNumber: r.floorNumber ?? null,
        capacity: r.capacity ?? null,
        address: r.address ?? null,
      }));
      const { items, count, limit, truncated } = toEnvelope(rooms, maxResults, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  // Contacts
  {
    name: 'search_contacts',
    description: 'Search contacts by name, or list all contacts',
    inputSchema: {
      type: 'object',
      properties: { q: prop('string', 'Search by name'), folderId: prop('string', 'Contact folder ID'), maxResults: prop('number', 'Max results') },
      required: [],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      // Deny-list check on explicit folderId
      if (args.folderId && await isPathDenied(tenantId, session.userId, 'contacts', args.folderId)) {
        throw new Error('Access restricted by deny list');
      }
      // Route through the shared helper so the synthetic 'contacts-root' id
      // (returned by list_contact_folders for the default folder) maps back to
      // /me/contacts instead of /me/contactFolders/contacts-root/contacts.
      const apiPath = contactsApiPath(args.folderId);
      const maxResults = resolveMaxResults(args.maxResults, 25, 50);
      let query = graph.api(apiPath).select(CONTACT_SELECT_FIELDS).top(maxResults);
      if (args.q) {
        // Sanitize input to prevent OData filter injection
        const safeQ = String(args.q).replace(/'/g, "''");
        query = query.filter(`startswith(displayName,'${safeQ}') or startswith(givenName,'${safeQ}') or startswith(surname,'${safeQ}')`);
      }
      const result = await query.get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      let contacts = (result.value ?? []).map((c: any) => shapeContact(c));
      // Post-filter: when no folderId, normalize each contact's parentFolderId
      // to the synthetic 'contacts-root' key before deny-list comparison
      if (!args.folderId) {
        const filtered = [];
        for (const c of contacts) {
          const normalizedFolder = await normalizeContactFolderId(graph, session.userId, c.parentFolderId ?? '');
          if (await isPathDenied(tenantId, session.userId, 'contacts', normalizedFolder)) continue;
          filtered.push(c);
        }
        contacts = filtered;
      }
      // eslint-disable-next-line @typescript-eslint/no-unused-vars
      const cleaned = contacts.map(({ parentFolderId: _pf, ...rest }: { parentFolderId: string; [key: string]: unknown }) => rest);
      // Surface truncation so a recency-capped page never reads as the whole
      // address book. No q returns Graph's default ordering capped at
      // `limit`; @odata.nextLink signals the folder holds more.
      const { items, count, limit, truncated } = toEnvelope(cleaned, maxResults, graphCollectionHasMore(result));
      return { contacts: items, count, limit, truncated };
    },
  },
  {
    name: 'create_contact',
    description: 'Create a new contact. Supports notes, category tags, all three postal addresses, and secondary fields (middle/nick name, home phones, birthday, spouse).',
    inputSchema: {
      type: 'object',
      properties: { ...CONTACT_FIELD_PROPS, folderId: prop('string', 'Contact folder ID (from list_contact_folders; omit for the default folder)') },
      required: ['givenName'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      // Deny-list check: explicit folderId or resolved default contacts folder
      if (args.folderId && args.folderId !== 'contacts-root') {
        if (await isPathDenied(tenantId, session.userId, 'contacts', args.folderId)) throw new Error('Access restricted by deny list');
      } else {
        const defaultFolder = await resolveDefaultContactFolder(graph, session.userId);
        if (defaultFolder && await isPathDenied(tenantId, session.userId, 'contacts', defaultFolder)) throw new Error('Access to the default contacts folder is restricted by the deny list');
      }
      const contact = buildContactBody(args);
      const result = await graph.api(contactsApiPath(args.folderId)).post(contact);
      return { id: result.id, displayName: result.displayName, status: 'created' };
    },
  },
  {
    name: 'update_contact',
    description: "Update an existing contact's fields. Accepts the full field set: notes, category tags, all three postal addresses, and secondary fields. Only provided fields are changed.",
    inputSchema: {
      type: 'object',
      properties: { contactId: prop('string', 'Contact ID'), ...CONTACT_FIELD_PROPS },
      required: ['contactId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Resolve parent folder and check deny list for direct contactId operations
      const parentFolderId = await resolveContactParentFolder(graph, session.userId, args.contactId);
      if (parentFolderId) {
        const tenantId = getTenantIdFromSession(session);
        if (await isPathDenied(tenantId, session.userId, 'contacts', parentFolderId)) throw new Error('Access restricted by deny list');
      }
      const patch = buildContactBody(args);
      const result = await graph.api(`/me/contacts/${args.contactId}`).patch(patch);
      return { id: result.id, displayName: result.displayName, status: 'updated' };
    },
  },
  {
    name: 'delete_contact',
    description: 'Delete a contact',
    inputSchema: { type: 'object', properties: { contactId: prop('string', 'Contact ID') }, required: ['contactId'] },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Resolve parent folder and check deny list for direct contactId operations
      const parentFolderId = await resolveContactParentFolder(graph, session.userId, args.contactId);
      if (parentFolderId) {
        const tenantId = getTenantIdFromSession(session);
        if (await isPathDenied(tenantId, session.userId, 'contacts', parentFolderId)) throw new Error('Access restricted by deny list');
      }
      await graph.api(`/me/contacts/${args.contactId}`).delete();
      return { status: 'deleted', contactId: args.contactId };
    },
  },
  {
    name: 'list_contact_folders',
    description: 'List contact folders. Returns { folders: [{ id, name, parentFolderId }], count }. Use a returned id as the folderId for create_contact / create_contacts_batch. The default folder is reported as the synthetic id "contacts-root".',
    inputSchema: { type: 'object', properties: {}, required: [] },
    handler: async (_args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const [childResult, rootCheck] = await Promise.all([
        graph.api('/me/contactFolders').select('id,displayName,parentFolderId').top(100).get(),
        graph.api('/me/contacts').select('id').top(1).get().catch(() => ({ value: [] })),
      ]);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const children = (childResult.value ?? []).map((f: any) => ({ id: f.id, name: f.displayName, parentFolderId: f.parentFolderId ?? null }));
      // /me/contactFolders never includes the root "Contacts" folder — synthesize it
      // when the root holds contacts or there are no child folders (so it's never empty).
      const rootHasContacts = (rootCheck.value ?? []).length > 0;
      const folders = rootHasContacts || children.length === 0
        ? [{ id: 'contacts-root', name: 'Contacts (Default)', parentFolderId: null }, ...children]
        : children;
      const tenantId = getTenantIdFromSession(session);
      const allowed = await filterDeniedPaths(tenantId, session.userId, 'contacts', folders);
      // The child-folder page is capped at 100; surface truncation from its
      // nextLink so a caller with >100 contact folders is not misled.
      return { folders: allowed, count: allowed.length, limit: 100, truncated: graphCollectionHasMore(childResult) };
    },
  },
  {
    name: 'create_contact_folder',
    description: 'Create a contact folder (top-level, or a child of an existing folder). Returns { id, displayName, parentFolderId }. The returned id is usable as the folderId for create_contact / create_contacts_batch.',
    inputSchema: {
      type: 'object',
      properties: {
        displayName: prop('string', 'Folder display name'),
        parentFolderId: prop('string', 'Parent contact folder ID (from list_contact_folders; omit for a top-level folder)'),
      },
      required: ['displayName'],
    },
    handler: async (args, session) => {
      if (!args.displayName) throw new Error('displayName is required');
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      // Guard the parent: creating a child under a deny-listed folder is denied.
      if (args.parentFolderId && args.parentFolderId !== 'contacts-root') {
        if (await isPathDenied(tenantId, session.userId, 'contacts', args.parentFolderId)) throw new Error('Access restricted by deny list');
      }
      const apiPath = args.parentFolderId && args.parentFolderId !== 'contacts-root'
        ? `/me/contactFolders/${args.parentFolderId}/childFolders`
        : '/me/contactFolders';
      const result = await graph.api(apiPath).post({ displayName: args.displayName });
      return { id: result.id, displayName: result.displayName, parentFolderId: result.parentFolderId ?? null, status: 'created' };
    },
  },
  {
    name: 'create_contacts_batch',
    description: 'Create many contacts in one call using Graph $batch (chunked at 20 per request), for bulk imports where one create_contact call each is impractical. Each item accepts the same fields as create_contact. Returns { created, failed, results: [{ index, id?, status, error? }] }.',
    inputSchema: {
      type: 'object',
      properties: {
        contacts: {
          type: 'array',
          description: 'Contacts to create. Each item accepts the same fields as create_contact (givenName required per item).',
          items: { type: 'object', properties: CONTACT_FIELD_PROPS, required: ['givenName'] },
        },
        folderId: prop('string', 'Contact folder ID applied to every contact in the batch (from list_contact_folders; omit for the default folder)'),
      },
      required: ['contacts'],
    },
    handler: async (args, session) => {
      const input = Array.isArray(args.contacts) ? args.contacts : [];
      if (input.length === 0) throw new Error('contacts must be a non-empty array');
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      // Deny-list check once for the shared target folder.
      if (args.folderId && args.folderId !== 'contacts-root') {
        if (await isPathDenied(tenantId, session.userId, 'contacts', args.folderId)) throw new Error('Access restricted by deny list');
      } else {
        const defaultFolder = await resolveDefaultContactFolder(graph, session.userId);
        if (defaultFolder && await isPathDenied(tenantId, session.userId, 'contacts', defaultFolder)) throw new Error('Access to the default contacts folder is restricted by the deny list');
      }
      const url = contactsApiPath(args.folderId);
      const results = await runContactBatch(graph, input, url);
      const created = results.filter((r) => r.status === 'created').length;
      return { created, failed: results.length - created, results };
    },
  },
  // OneNote
  {
    name: 'list_notebooks',
    description: "List the user's OneNote notebooks",
    inputSchema: { type: 'object', properties: {}, required: [] },
    handler: async (_args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const result = await graph.api('/me/onenote/notebooks').select('id,displayName,createdDateTime,lastModifiedDateTime').top(100).get();
      const { items, count, limit, truncated } = toEnvelope(result.value ?? [], 100, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'create_notebook',
    description: 'Create a new OneNote notebook',
    inputSchema: { type: 'object', properties: { displayName: prop('string', 'Notebook name') }, required: ['displayName'] },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const result = await graph.api('/me/onenote/notebooks').post({ displayName: args.displayName });
      return { id: result.id, displayName: result.displayName, status: 'created' };
    },
  },
  {
    name: 'list_sections',
    description: 'List sections in a OneNote notebook',
    inputSchema: { type: 'object', properties: { notebookId: prop('string', 'Notebook ID') }, required: ['notebookId'] },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Deny-list check on notebookId
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'onenote', args.notebookId)) throw new Error('Access restricted by deny list');
      const result = await graph.api(`/me/onenote/notebooks/${args.notebookId}/sections`).select('id,displayName,createdDateTime').top(100).get();
      // Post-filter sections by section ID against deny list (section-level blocks)
      const allSections = (result.value ?? []).map((s: { id: string; displayName: string; createdDateTime: string }) => ({ ...s }));
      const allowed = await filterDeniedPaths(tenantId, session.userId, 'onenote', allSections);
      const { items, count, limit, truncated } = toEnvelope(allowed, 100, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'create_section',
    description: 'Create a new section in a OneNote notebook',
    inputSchema: {
      type: 'object',
      properties: { notebookId: prop('string', 'Notebook ID'), displayName: prop('string', 'Section name') },
      required: ['notebookId', 'displayName'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Deny-list check on notebookId
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'onenote', args.notebookId)) throw new Error('Access restricted by deny list');
      const result = await graph.api(`/me/onenote/notebooks/${args.notebookId}/sections`).post({ displayName: args.displayName });
      return { id: result.id, displayName: result.displayName, status: 'created' };
    },
  },
  {
    name: 'create_onenote_page',
    description: 'Create a new page in a OneNote section',
    inputSchema: {
      type: 'object',
      properties: { sectionId: prop('string', 'Section ID'), title: prop('string', 'Page title'), htmlContent: prop('string', 'Page content in HTML') },
      required: ['sectionId', 'title'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      // Check section-level deny first, then parent notebook
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'onenote', args.sectionId)) throw new Error('Access restricted by deny list — section is blocked');
      const notebookId = await resolveSectionNotebook(graph, args.sectionId);
      if (notebookId) {
        if (await isPathDenied(tenantId, session.userId, 'onenote', notebookId)) throw new Error('Access restricted by deny list — parent notebook is blocked');
      }
      // HTML-escape the title to prevent XSS via title injection
      const safeTitle = String(args.title).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      const html = `<!DOCTYPE html><html><head><title>${safeTitle}</title></head><body>${args.htmlContent ?? ''}</body></html>`;
      const result = await graph.api(`/me/onenote/sections/${args.sectionId}/pages`).header('Content-Type', 'text/html').post(html);
      return { id: result.id, title: result.title, contentUrl: result.contentUrl, status: 'created' };
    },
  },
  // Teams
  {
    name: 'list_teams',
    description:
      'List the Microsoft Teams the calling user is a member of (Graph GET /me/joinedTeams). Use this to discover the ' +
      'teamId you need for list_channels or send_channel_message. Teams on the deny list are excluded.',
    inputSchema: { type: 'object', properties: {}, required: [] },
    handler: async (_args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const result = await graph.api('/me/joinedTeams').select('id,displayName,description').top(100).get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const teams = (result.value ?? []).map((t: any) => ({ id: t.id, name: t.displayName, description: t.description ?? null }));
      const tenantId = getTenantIdFromSession(session);
      const allowed = await filterDeniedPaths(tenantId, session.userId, 'teams', teams);
      const { items, count, limit, truncated } = toEnvelope(allowed, 100, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'list_channels',
    description:
      'List the channels in a Team (Graph GET /teams/{teamId}/channels). Use this after list_teams to find the ' +
      'channelId for send_channel_message. Channels on the deny list are excluded.',
    inputSchema: {
      type: 'object',
      properties: { teamId: prop('string', 'Team ID (from list_teams)') },
      required: ['teamId'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      // Block enumerating channels of a denied team (deny list keys on team/channel IDs).
      if (await isPathDenied(tenantId, session.userId, 'teams', args.teamId)) throw new Error('Access restricted by deny list');
      // Graph GET /teams/{teamId}/channels does not support $top — appending it fails the whole request.
      const result = await graph.api(`/teams/${args.teamId}/channels`).select('id,displayName,membershipType').get();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const channels = (result.value ?? []).map((c: any) => ({ id: c.id, displayName: c.displayName, path: c.id, membershipType: c.membershipType ?? 'standard' }));
      const allowed = await filterDeniedPaths(tenantId, session.userId, 'teams', channels);
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const shaped = allowed.map((c: any) => ({ id: c.id, displayName: c.displayName, membershipType: c.membershipType }));
      // Graph does not page /channels (no $top support), so the returned set is the
      // full channel list; `limit` reflects that and truncated tracks any nextLink.
      const { items, count, limit, truncated } = toEnvelope(shaped, shaped.length, graphCollectionHasMore(result));
      return { items, count, limit, truncated };
    },
  },
  {
    name: 'send_chat_message',
    description:
      'Send a message to an existing Teams 1:1 or group chat (Graph POST /chats/{chatId}/messages). The chat must ' +
      'already exist — this does not create new chats. chatId comes from a prior Teams chat the user is part of. ' +
      'Blocked if the chat is on the deny list.',
    inputSchema: {
      type: 'object',
      properties: {
        chatId: prop('string', 'Chat ID (Graph chat thread ID)'),
        content: prop('string', 'Message body'),
        contentType: prop('string', 'text (default) or html', { enum: ['text', 'html'] }),
      },
      required: ['chatId', 'content'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      if (await isPathDenied(tenantId, session.userId, 'teams', args.chatId)) throw new Error('Access restricted by deny list');
      const body = { body: { contentType: args.contentType === 'html' ? 'html' : 'text', content: args.content ?? '' } };
      const result = await graph.api(`/chats/${args.chatId}/messages`).post(body);
      return { id: result.id, chatId: args.chatId, webUrl: result.webUrl ?? null, status: 'sent' };
    },
  },
  {
    name: 'send_channel_message',
    description:
      'Post a message to a Teams channel (Graph POST /teams/{teamId}/channels/{channelId}/messages). Use list_teams then ' +
      'list_channels to resolve the IDs. Good for "ping the team" style notifications. Blocked if the team or channel is ' +
      'on the deny list.',
    inputSchema: {
      type: 'object',
      properties: {
        teamId: prop('string', 'Team ID (from list_teams)'),
        channelId: prop('string', 'Channel ID (from list_channels)'),
        content: prop('string', 'Message body'),
        contentType: prop('string', 'text (default) or html', { enum: ['text', 'html'] }),
      },
      required: ['teamId', 'channelId', 'content'],
    },
    handler: async (args, session) => {
      const token = await getValidAccessTokenForSession(session);
      const graph = createGraphClient(token);
      const tenantId = getTenantIdFromSession(session);
      // Deny-list check on both the team and the channel (deny list keys on IDs for teams).
      if (await isPathDenied(tenantId, session.userId, 'teams', args.teamId)) throw new Error('Access restricted by deny list');
      if (await isPathDenied(tenantId, session.userId, 'teams', args.channelId)) throw new Error('Access restricted by deny list');
      const body = { body: { contentType: args.contentType === 'html' ? 'html' : 'text', content: args.content ?? '' } };
      const result = await graph.api(`/teams/${args.teamId}/channels/${args.channelId}/messages`).post(body);
      return { id: result.id, teamId: args.teamId, channelId: args.channelId, webUrl: result.webUrl ?? null, status: 'sent' };
    },
  },
];

// ── JSON-RPC handler ──

const toolMap = new Map(tools.map((t) => [t.name, t]));

// The login endpoint is a function route served at the site root, never under
// whatever subpath the SPA lives at, so the auth link is built from the ORIGIN
// of the configured URL. This used to strip an exact trailing `/admin`, which
// left every other host-reserved value intact: `FRONTEND_URL=https://host/runtime`
// or `https://host/admin/settings` emitted a login link the Functions host
// intercepts before this app runs. resolveAuthUrlBase drops the path outright
// and shares the reserved-path check with the callback redirect.
//
// Resolved per message rather than at import so the value in force is the one
// reported and the path is exercisable by a test. The problem is reported once
// per process: it is a deployment misconfiguration, not a per-request event.
let authUrlProblemReported = false;

function authLoginUrl(): string {
  const base = resolveAuthUrlBase(undefined, undefined, (message) => {
    if (authUrlProblemReported) return;
    authUrlProblemReported = true;
    console.error(message);
  });
  return `${base}/api/auth/login`;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function handleJsonRpc(msg: any, auth: { userId: string; session: any } | null, context: InvocationContext): Promise<any> {
  const { method, params, id } = msg;

  if (method === 'initialize') {
    return {
      jsonrpc: '2.0',
      id,
      result: {
        protocolVersion: '2024-11-05',
        capabilities: { tools: { listChanged: false } },
        serverInfo: { name: process.env.MCP_INSTANCE_NAME ?? 'm365-mcp', version: '2.0.0' },
      },
    };
  }

  if (method === 'notifications/initialized') {
    return null; // no response for notifications
  }

  if (method === 'tools/list') {
    // Filter tool catalog by enabledServices (tenant) + user overrides, and
    // hide the write tools of read-only services. A connector risk
    // review scores what the server advertises, so a read-only service must
    // not list tools it will refuse at tools/call. The call-time refusal
    // below stays as the backstop.
    let visibleTools = tools;
    if (auth) {
      try {
        const tenantId = getTenantIdFromSession(auth!.session);
        const enabled = await getEnabledServices(tenantId);
        const readOnly = await getReadOnlyServices(tenantId);
        const userDisabled = await getUserServiceOverrides(tenantId, auth.userId);
        const mailIndexingOff = await isMailIndexingDisabled(tenantId, auth.userId);
        visibleTools = tools.filter((t) => {
          const svc = TOOL_SERVICE_MAP[t.name];
          if (!svc) return true;
          if (!enabled.includes(svc)) return false;
          if (userDisabled.includes(svc)) return false;
          if (svc === 'mail' && mailIndexingOff) return false;
          if (readOnly.includes(svc) && WRITE_TOOLS.has(t.name)) return false;
          return true;
        });
      } catch { /* fall through — show all if we can't load settings */ }
    }
    return {
      jsonrpc: '2.0',
      id,
      result: {
        tools: visibleTools.map((t) => ({
          name: t.name,
          description: t.description,
          inputSchema: t.inputSchema,
        })),
      },
    };
  }

  if (method === 'tools/call') {
    const toolName = params?.name;
    const toolArgs = params?.arguments ?? {};
    const tool = toolMap.get(toolName);

    if (!tool) {
      return { jsonrpc: '2.0', id, error: { code: -32601, message: `Unknown tool: ${toolName}` } };
    }

    if (!auth) {
      return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: `Session expired. Please re-authenticate at: ${authLoginUrl()}` }], isError: true } };
    }

    try {
      // Enforce enabledServices — reject tools for disabled service categories
      const tenantId = getTenantIdFromSession(auth!.session);
      const toolService = TOOL_SERVICE_MAP[toolName];
      if (toolService) {
        const enabled = await getEnabledServices(tenantId);
        if (!enabled.includes(toolService)) {
          logAccess({
            tenantId,
            userId: auth!.userId,
            userEmail: auth!.session.email,
            deviceLabel: auth!.session.deviceLabel,
            operation: toolName,
            result: 'denied',
            reason: `Service "${toolService}" is not enabled`,
            source: 'mcp',
          });
          return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: `Service "${toolService}" is not enabled for this tenant. An admin can enable it in the admin UI.` }], isError: true } };
        }
        // Per-user override — admin can disable a tenant-enabled service for specific users
        const userDisabled = await getUserServiceOverrides(tenantId, auth!.userId);
        if (userDisabled.includes(toolService)) {
          logAccess({
            tenantId,
            userId: auth!.userId,
            userEmail: auth!.session.email,
            deviceLabel: auth!.session.deviceLabel,
            operation: toolName,
            result: 'denied',
            reason: `Service "${toolService}" disabled for user`,
            source: 'mcp',
          });
          return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: `Service "${toolService}" is disabled for your account. Contact your admin to enable it.` }], isError: true } };
        }
        // Mail indexing disable flag — IR-adjacent roles
        if (toolService === 'mail' && await isMailIndexingDisabled(tenantId, auth!.userId)) {
          logAccess({
            tenantId,
            userId: auth!.userId,
            userEmail: auth!.session.email,
            deviceLabel: auth!.session.deviceLabel,
            operation: toolName,
            result: 'denied',
            reason: 'mail indexing disabled',
            source: 'mcp',
          });
          return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: 'Mail access is disabled for your account. Contact your admin.' }], isError: true } };
        }
      }

      // Enforce read-only mode — refuse mutating tools when the
      // tool's service is configured read-only for the tenant. This mirrors the
      // HTTP-route enforcement so the MCP surface cannot be used to bypass it.
      if (toolService && WRITE_TOOLS.has(toolName)) {
        const readOnly = await getReadOnlyServices(tenantId);
        if (readOnly.includes(toolService)) {
          logAccess({
            tenantId,
            userId: auth!.userId,
            userEmail: auth!.session.email,
            deviceLabel: auth!.session.deviceLabel,
            operation: toolName,
            result: 'denied',
            reason: `Service "${toolService}" is read-only`,
            source: 'mcp',
          });
          return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: `Service "${toolService}" is in read-only mode; write operations are disabled. An admin can change this in the admin UI.` }], isError: true } };
        }
      }

      // Enforce allowedSites — for SharePoint tools that accept siteId, validate
      // the target site is in the allow-list (empty list = allow all)
      if (SITE_SCOPED_TOOLS.has(toolName) && toolArgs.siteId) {
        const allowed = await getAllowedSites(tenantId);
        if (allowed.length > 0) {
          const siteAllowed = allowed.some((s) => s.id === toolArgs.siteId);
          if (!siteAllowed) {
            logAccess({
              tenantId,
              userId: auth!.userId,
              userEmail: auth!.session.email,
              deviceLabel: auth!.session.deviceLabel,
              operation: toolName,
              resource: toolArgs.siteId,
              result: 'denied',
              reason: `Site not in allowedSites`,
              source: 'mcp',
            });
            return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: `Site "${toolArgs.siteId}" is not in the allowed sites list. An admin can add it in the admin UI.` }], isError: true } };
          }
        }
      }

      // Validate opaque ID parameters: reject values that contain
      // characters that could alter the resulting Graph API URL path or inject
      // OData query options.  Non-ID params (q, path, displayName, content,
      // body text) are intentionally excluded from this check.
      try {
        assertOpaqueIds(toolArgs, OPAQUE_ID_PARAMS);
      } catch (idErr: unknown) {
        const msg = idErr instanceof Error ? idErr.message : 'Invalid parameter';
        return {
          jsonrpc: '2.0',
          id,
          result: { content: [{ type: 'text', text: `Bad request: ${msg}` }], isError: true },
        };
      }

      // Reject any parameter the tool does not declare, grounded in the tool's
      // own inputSchema. The connector must never accept a parameter,
      // ignore it, and return success — a silently-dropped argument looks
      // identical to an honored one from the caller's side. Enforcing it here,
      // once, covers every tool including any added later.
      const unsupported = findUnsupportedArgs(toolArgs, tool.inputSchema as { properties?: Record<string, unknown> });
      if (unsupported.length > 0) {
        logAccess({
          tenantId,
          userId: auth!.userId,
          userEmail: auth!.session.email,
          deviceLabel: auth!.session.deviceLabel,
          operation: toolName,
          result: 'denied',
          reason: `unsupported parameter(s): ${unsupported.join(', ')}`,
          source: 'mcp',
        });
        return {
          jsonrpc: '2.0',
          id,
          result: {
            content: [{ type: 'text', text: unsupportedArgsMessage(toolName, unsupported, tool.inputSchema as { properties?: Record<string, unknown> }) }],
            isError: true,
          },
        };
      }

      const result = await tool.handler(toolArgs, auth.session);
      // Build resource identifier from all relevant tool args (coarser → finer grain)
      const resource = [
        toolArgs.siteId,
        toolArgs.driveId,
        toolArgs.listId,
        toolArgs.itemId,
        toolArgs.messageId,
        toolArgs.eventId,
        toolArgs.messageOrEventId,
        toolArgs.folderId,
        toolArgs.attachmentId,
        toolArgs.mailboxId,
        toolArgs.destinationFolderId,
        toolArgs.contactId,
        toolArgs.teamId,
        toolArgs.channelId,
        toolArgs.chatId,
        toolArgs.path,
      ].filter(Boolean).join('/') || undefined;
      logAccess({
        tenantId,
        userId: auth!.userId,
        userEmail: auth!.session.email,
        deviceLabel: auth!.session.deviceLabel,
        operation: toolName,
        resource,
        result: 'allowed',
        source: 'mcp',
      });
      return {
        jsonrpc: '2.0',
        id,
        result: {
          content: [{ type: 'text', text: JSON.stringify(result, null, 2) }],
        },
      };
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : 'Unknown error';
      context.error(`MCP tool ${toolName} error:`, message);

      // Log denied access for policy/deny-list errors thrown inside tool handlers
      const isDenied = message.includes('restricted by the deny list') ||
        message.includes('Access restricted by deny list') ||
        message.includes('read-only mode') ||
        message.includes('not in the allowed sites') ||
        message.includes('is not enabled') ||
        message.includes('disabled for your account') ||
        message.includes(ENFORCED_MODE_MARKER);
      if (auth && isDenied) {
        const errorResource = [
          toolArgs.siteId, toolArgs.driveId, toolArgs.listId, toolArgs.itemId,
          toolArgs.messageId, toolArgs.eventId, toolArgs.messageOrEventId, toolArgs.folderId,
          toolArgs.attachmentId, toolArgs.mailboxId, toolArgs.path,
        ].filter(Boolean).join('/') || undefined;
        logAccess({
          tenantId: getTenantIdFromSession(auth.session),
          userId: auth.userId,
          userEmail: auth.session.email,
          deviceLabel: auth.session.deviceLabel,
          operation: toolName,
          resource: errorResource,
          result: 'denied',
          reason: message,
          source: 'mcp',
        });
      }

      if (message.includes('Re-authentication required')) {
        return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: `Session expired. Please re-authenticate at: ${authLoginUrl()}\n\nError: ${message}` }], isError: true } };
      }
      return { jsonrpc: '2.0', id, result: { content: [{ type: 'text', text: `Error: ${message}` }], isError: true } };
    }
  }

  return { jsonrpc: '2.0', id, error: { code: -32601, message: `Unknown method: ${method}` } };
}

// ── Azure Function handler ──

async function mcpEndpoint(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  // Authenticate via session token (Bearer header, x-session-token, or mcp_session cookie)
  const auth = await authenticateRequest(request);

  try {
    const body = await request.json();

    // Handle batch (array) or single message
    if (Array.isArray(body)) {
      const responses = [];
      for (const msg of body) {
        const resp = await handleJsonRpc(msg, auth, context);
        if (resp) responses.push(resp);
      }
      return { status: 200, jsonBody: responses };
    }

    const response = await handleJsonRpc(body, auth, context);
    if (!response) {
      return { status: 204 }; // notification, no response
    }
    return { status: 200, jsonBody: response };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : 'Unknown error';
    context.error('MCP endpoint error:', message);
    return { status: 500, jsonBody: { jsonrpc: '2.0', error: { code: -32603, message: 'Internal server error' } } };
  }
}

app.http('mcpEndpoint', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/mcp',
  handler: withSecurity(mcpEndpoint),
});
