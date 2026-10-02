/**
 * MCP tool manifest for the M365 MCP server.
 *
 * Callers (Claude, other agents) invoke tools by POSTing a JSON-RPC `tools/call`
 * request to the MCP endpoint (`/api/mcp`); `getManifest()` and the live
 * `tools/list` response are the two discovery surfaces and MUST agree on each
 * tool's input schema. The `schedulingTeamsTools` suite asserts that sync.
 *
 * Most tools also have a dedicated REST Azure Function (recorded in `endpoint` /
 * `method`); newer MCP-native tools are reachable only through the JSON-RPC
 * surface and carry `endpoint: '/api/mcp'`, `method: 'POST'` to reflect that.
 * Keep `endpoint` truthful — do not point it at a REST route that has no
 * corresponding Azure Function.
 */

export interface McpToolParameter {
  name: string;
  type: 'string' | 'number' | 'boolean' | 'array';
  /** Element type when `type` is 'array'. Emitted as `items: { type }` in the JSON Schema. */
  items?: 'string' | 'number' | 'boolean' | 'object';
  /**
   * Full JSON Schema for array items when they are objects (`items: 'object'`). Emitted verbatim
   * as the `items` schema, overriding the scalar `{ type }` form. Used by `attachments`.
   */
  itemsSchema?: Record<string, unknown>;
  description: string;
  required: boolean;
  enum?: string[];
}

export interface McpTool {
  name: string;
  description: string;
  parameters: McpToolParameter[];
  /** REST route for the tool's Azure Function, or '/api/mcp' for JSON-RPC-only tools. */
  endpoint: string;
  method: 'GET' | 'POST' | 'PATCH' | 'DELETE';
}

// Shared `attachments` parameter for send_mail / create_draft (; drive references).
// Kept in sync with the live tools/list schema in mcpEndpoint.ts (ATTACHMENTS_SCHEMA).
const ATTACHMENTS_PARAM: McpToolParameter = {
  name: 'attachments',
  type: 'array',
  items: 'object',
  itemsSchema: {
    type: 'object',
    properties: {
      name: { type: 'string', description: 'File name shown to the recipient (e.g. "budget.xlsx"). Required with content; optional override for drive items.' },
      contentType: { type: 'string', description: 'MIME type (e.g. "application/pdf"). Defaults to application/octet-stream, or the drive file\'s own type.' },
      content: { type: 'string', description: 'Base64-encoded file bytes.' },
      driveItemId: { type: 'string', description: 'OneDrive item ID of a file to attach (from list_onedrive).' },
      siteId: { type: 'string', description: 'SharePoint site ID of a file to attach; use with itemId.' },
      itemId: { type: 'string', description: 'SharePoint file item ID to attach (from list_folders or search_sharepoint); use with siteId.' },
      driveId: { type: 'string', description: 'SharePoint document library (drive) ID, when the file is not in the site\'s default library.' },
    },
  },
  description:
    'Optional file attachments. Each item takes exactly one source: { name, contentType, content } ' +
    'with content as base64-encoded file bytes; { driveItemId } to attach a file from the user\'s ' +
    'OneDrive; or { siteId, itemId } (plus driveId for a non-default library) to attach a SharePoint ' +
    'file. For a OneDrive or SharePoint file the server fetches the bytes itself, so prefer it over ' +
    'reading the file and re-sending it as base64; name and contentType default to the file\'s own. ' +
    'Files under 3 MB attach inline; larger files (up to 150 MB) upload via a Graph upload session. ' +
    'Inline base64 inflates size ~33% and travels inside the request body, so keep inline content to a ' +
    'few MB in practice.',
  required: false,
};

// Shared `from` parameter for send_mail / create_draft / send_draft. Kept in sync with the
// live tools/list schema in mcpEndpoint.ts (FROM_PROP).
const FROM_PARAM: McpToolParameter = {
  name: 'from',
  type: 'string',
  description:
    "Send as this address. Must be one of the mailbox's own proxy addresses (an alias such as " +
    'alias@example.com on a user@example.com mailbox); anything else is rejected. Omit to use ' +
    'the primary address.',
  required: false,
};

// Shared parameters for the reply/forward draft tools. Kept in sync with the live
// tools/list schema in mcpEndpoint.ts (REPLY_SCHEMA_PROPS). No recipient parameters: Graph
// derives them from the source message (forward_message adds its own `to`).
const REPLY_PARAMS: McpToolParameter[] = [
  {
    name: 'messageId',
    type: 'string',
    description: 'ID of the message being replied to / forwarded (from search_mail or read_message).',
    required: true,
  },
  {
    name: 'comment',
    type: 'string',
    description: 'Your new text. Graph inserts it above the quoted original.',
    required: true,
  },
  {
    name: 'bodyType',
    type: 'string',
    description:
      "How to interpret `comment`: 'text' (default — escaped and newlines become line breaks) or " +
      "'html' (inserted verbatim).",
    required: false,
    enum: ['text', 'html'],
  },
  {
    name: 'mailboxId',
    type: 'string',
    description: 'Mailbox (default: me)',
    required: false,
  },
  {
    name: 'from',
    type: 'string',
    description:
      "Send as this address. Must be one of the mailbox's own proxy addresses; anything else is rejected. " +
      "Omit to default to whichever of the mailbox's addresses the original was sent to (To, then CC), " +
      "matching Outlook's reply-from-alias behavior; the primary address when none match.",
    required: false,
  },
];

export const TOOLS: McpTool[] = [
  // ── SharePoint ──────────────────────────────────────────────────────────────
  {
    name: 'list_sites',
    description:
      'Enumerate accessible SharePoint sites in the Microsoft 365 tenant. Returns site IDs, ' +
      'display names, and web URLs. Use the returned `id` values when calling list_folders or ' +
      'search_sharepoint.',
    parameters: [],
    endpoint: '/api/sharepoint/sites',
    method: 'GET',
  },
  {
    name: 'list_folders',
    description:
      'List a SharePoint folder\'s children. Returns folders only by default; set ' +
      'includeFiles=true to also return files — each file carries the `id` you pass to ' +
      'read_file, so this is how you obtain a file\'s itemId. Results are automatically ' +
      'filtered by the admin global deny list and the user\'s personal deny list. Pass ' +
      '`parentId` to navigate into sub-folders and `offset` to page through large folders.',
    parameters: [
      {
        name: 'siteId',
        type: 'string',
        description: 'SharePoint site ID (from list_sites).',
        required: true,
      },
      {
        name: 'driveId',
        type: 'string',
        description: 'Drive ID within the site. Omit to use the site\'s default document library.',
        required: false,
      },
      {
        name: 'parentId',
        type: 'string',
        description: 'Item ID of the parent folder. Omit to list the drive root.',
        required: false,
      },
      {
        name: 'includeFiles',
        type: 'boolean',
        description: 'When true, also return files (each with id/name/size/lastModifiedDateTime) alongside folders. Default false.',
        required: false,
      },
      {
        name: 'offset',
        type: 'number',
        description: 'Number of children to skip before returning, for paging into a large folder. Default 0.',
        required: false,
      },
      {
        name: 'maxResults',
        type: 'number',
        description: 'Maximum number of children to return (default 100, max 200).',
        required: false,
      },
    ],
    endpoint: '/api/sharepoint/folders',
    method: 'GET',
  },
  {
    name: 'read_file',
    description:
      'Fetch the content of a specific SharePoint file. Text-based files (plain text, JSON, ' +
      'XML, etc.) are returned as UTF-8 strings; binary files are base64-encoded. Access is ' +
      'blocked if the file\'s path is on the deny list.',
    parameters: [
      {
        name: 'siteId',
        type: 'string',
        description: 'SharePoint site ID.',
        required: true,
      },
      {
        name: 'itemId',
        type: 'string',
        description: 'File item ID.',
        required: true,
      },
      {
        name: 'driveId',
        type: 'string',
        description: 'Drive ID (optional, uses default drive if omitted).',
        required: false,
      },
    ],
    endpoint: '/api/sharepoint/files/{itemId}',
    method: 'GET',
  },
  {
    name: 'search_sharepoint',
    description:
      'Full-text search across SharePoint using the Microsoft Search API. Results are ' +
      'restricted to the tenant allow-listed sites, and results from deny-listed folders ' +
      'are excluded, before the response is returned. Optionally scope the search to a ' +
      'single site.',
    parameters: [
      {
        name: 'q',
        type: 'string',
        description: 'Search query string.',
        required: true,
      },
      {
        name: 'siteId',
        type: 'string',
        description: 'Restrict search to this site ID (optional).',
        required: false,
      },
      {
        name: 'maxResults',
        type: 'number',
        description: 'Maximum number of results to return (default 25, max 50).',
        required: false,
      },
    ],
    endpoint: '/api/sharepoint/search',
    method: 'GET',
  },

  // ── Exchange / Outlook ──────────────────────────────────────────────────────
  {
    name: 'send_mail',
    description:
      'Compose and deliver an email message, optionally with file attachments (see the ' +
      'attachments parameter). The actual delivery behavior depends on the ' +
      'calling user\'s email output mode: in "draft" mode (the default for new users) the ' +
      'message is saved to Drafts so the user can review it before sending; in "send" mode ' +
      'the message is delivered immediately. The response includes a `status` field — ' +
      '"sent" or "queued_as_draft" — so the caller can inform the user of the outcome. ' +
      'Use get_email_output_mode to check the user\'s current setting.',
    parameters: [
      {
        name: 'subject',
        type: 'string',
        description: 'Email subject line.',
        required: true,
      },
      {
        name: 'body',
        type: 'string',
        description: 'Email body content.',
        required: true,
      },
      {
        name: 'to',
        type: 'string',
        description: 'Comma-separated list of recipient email addresses.',
        required: true,
      },
      {
        name: 'cc',
        type: 'string',
        description: 'Comma-separated CC recipients (optional).',
        required: false,
      },
      {
        name: 'bcc',
        type: 'string',
        description: 'Comma-separated BCC recipients (optional).',
        required: false,
      },
      {
        name: 'bodyType',
        type: 'string',
        description: '"text" (default) or "html".',
        required: false,
        enum: ['text', 'html'],
      },
      ATTACHMENTS_PARAM,
      FROM_PARAM,
    ],
    endpoint: '/api/mail/send',
    method: 'POST',
  },
  {
    name: 'create_draft',
    description:
      'Save an email as a draft in the Drafts folder without sending, optionally with file ' +
      'attachments (see the attachments parameter). Use this when you ' +
      'explicitly want a draft regardless of the user\'s email output mode setting.',
    parameters: [
      {
        name: 'subject',
        type: 'string',
        description: 'Email subject line.',
        required: true,
      },
      {
        name: 'body',
        type: 'string',
        description: 'Email body content.',
        required: true,
      },
      {
        name: 'to',
        type: 'string',
        description: 'Comma-separated list of recipient email addresses.',
        required: true,
      },
      {
        name: 'cc',
        type: 'string',
        description: 'Comma-separated CC recipients (optional).',
        required: false,
      },
      {
        name: 'bcc',
        type: 'string',
        description: 'Comma-separated BCC recipients (optional).',
        required: false,
      },
      {
        name: 'bodyType',
        type: 'string',
        description: '"text" (default) or "html".',
        required: false,
        enum: ['text', 'html'],
      },
      ATTACHMENTS_PARAM,
      FROM_PARAM,
    ],
    endpoint: '/api/mail/drafts',
    method: 'POST',
  },
  {
    name: 'reply_to_message',
    description:
      'Reply to an email — creates a draft reply (never sends) addressed to the original sender, with the ' +
      'quoted original body below your text and the threading headers set so it collapses into the ' +
      "recipient's existing conversation. Always use this instead of create_draft with a \"RE:\" subject: a " +
      'hand-built reply carries no quoted history and no In-Reply-To/References, so the recipient sees it as a ' +
      'brand-new conversation. Review the returned draft, then dispatch it with send_draft.',
    parameters: [...REPLY_PARAMS],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'reply_all_to_message',
    description:
      'Reply-all to an email — same as reply_to_message, but the draft is addressed to the original sender ' +
      'and every other recipient (To and CC). Creates a draft; never sends — dispatch it with send_draft. ' +
      'Check the returned `to`/`cc` first, since reply-all widens the audience.',
    parameters: [...REPLY_PARAMS],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'forward_message',
    description:
      'Forward an email to new recipients — creates a draft (never sends) with the original message quoted ' +
      'below your text and the threading headers intact. Use this instead of pasting the original into ' +
      'create_draft, which loses both the attachments and the conversation link.',
    parameters: [
      ...REPLY_PARAMS,
      {
        name: 'to',
        type: 'array',
        items: 'string',
        description: 'Recipients to forward to',
        required: true,
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'send_draft',
    description:
      'Send an existing draft — the dispatch step for a draft made by reply_to_message, ' +
      'reply_all_to_message, forward_message, or create_draft. Takes the draft ID and sends it as-is, ' +
      "preserving the threading headers a reply draft carries. Subject to the user's email output mode: in " +
      "'send' mode it delivers immediately; in 'draft' mode it refuses and returns the draft's webLink, " +
      'because that mode exists to put a human between composition and delivery. Only works on drafts.',
    parameters: [
      {
        name: 'messageId',
        type: 'string',
        description: 'Draft message ID to send',
        required: true,
      },
      {
        name: 'mailboxId',
        type: 'string',
        description: 'Mailbox (default: me)',
        required: false,
      },
      {
        name: 'from',
        type: 'string',
        description:
          "Set the draft's From before sending. Must be one of the mailbox's own proxy addresses; anything " +
          'else is rejected. Omit to send with the From the draft already carries (a reply draft already ' +
          'defaults to the alias the original came in on).',
        required: false,
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'get_email_output_mode',
    description:
      'Get the current email output mode for the calling user. Returns "draft" (the default — ' +
      'send_mail saves to Drafts for review) or "send" (send_mail delivers immediately). ' +
      'Always call this before composing an email so you can tell the user what will happen. ' +
      'When `enforced` is true an administrator has pinned the mode to "draft" (`enforcedBy` says ' +
      'whether tenant-wide or for this user) and set_email_output_mode will refuse to change it.',
    parameters: [],
    endpoint: '/api/mail/settings',
    method: 'GET',
  },
  {
    name: 'set_email_output_mode',
    description:
      'Update the calling user\'s email output mode. Pass "draft" to require review before ' +
      'delivery (the default for new users) or "send" to allow immediate delivery. Refuses when ' +
      'an administrator enforces draft mode for the tenant or for this user; only an administrator ' +
      'can lift that in the admin UI.',
    parameters: [
      {
        name: 'emailOutputMode',
        type: 'string',
        description: '"draft" — save to Drafts for review; "send" — deliver immediately.',
        required: true,
        enum: ['draft', 'send'],
      },
    ],
    endpoint: '/api/mail/settings',
    method: 'POST',
  },
  {
    name: 'list_mailboxes',
    description:
      'List mailboxes the current user has access to — their primary mailbox and any ' +
      'delegated or shared mailboxes. Returns mailbox IDs to use with other mail tools.',
    parameters: [],
    endpoint: '/api/mail/mailboxes',
    method: 'GET',
  },
  {
    name: 'list_folders_mail',
    description:
      'List mail folders in a mailbox. Top-level folders (Inbox, Sent Items, etc.) are ' +
      'returned by default; pass `parentFolderId` to list sub-folders. Results are filtered ' +
      'by both the global and per-user mail deny lists.',
    parameters: [
      {
        name: 'mailboxId',
        type: 'string',
        description: 'Mailbox user ID or "me" for the current user (default: "me").',
        required: false,
      },
      {
        name: 'parentFolderId',
        type: 'string',
        description: 'Parent mail folder ID. Omit to list top-level folders.',
        required: false,
      },
    ],
    endpoint: '/api/mail/folders',
    method: 'GET',
  },
  {
    name: 'create_mail_folder',
    description:
      'Create a new mail folder in a mailbox. Pass `parentFolderId` to create a sub-folder; ' +
      'omit it to create a top-level folder. The folder name is checked against the deny list ' +
      'before creation. Returns the new folder\'s `id`, `displayName`, and `parentFolderId`.',
    parameters: [
      {
        name: 'displayName',
        type: 'string',
        description: 'Display name for the new folder (e.g. "_Notifications").',
        required: true,
      },
      {
        name: 'parentFolderId',
        type: 'string',
        description: 'Parent mail folder ID. Omit to create a top-level folder.',
        required: false,
      },
      {
        name: 'mailboxId',
        type: 'string',
        description: 'Mailbox user ID or "me" for the current user (default: "me").',
        required: false,
      },
    ],
    endpoint: '/api/mail/folders',
    method: 'POST',
  },
  {
    name: 'rename_mail_folder',
    description:
      'Rename an existing mail folder. Both the source folder name and the new name are ' +
      'checked against the deny list. Returns the updated folder\'s `id`, `displayName`, ' +
      'and `parentFolderId`.',
    parameters: [
      {
        name: 'folderId',
        type: 'string',
        description: 'Mail folder ID to rename (from list_folders_mail).',
        required: true,
      },
      {
        name: 'displayName',
        type: 'string',
        description: 'New display name for the folder.',
        required: true,
      },
      {
        name: 'mailboxId',
        type: 'string',
        description: 'Mailbox user ID or "me" for the current user (default: "me").',
        required: false,
      },
    ],
    endpoint: '/api/mail/folders/{folderId}',
    method: 'PATCH',
  },
  {
    name: 'move_mail_folder',
    description:
      'Move a mail folder to a different parent folder. Both the folder being moved and the ' +
      'destination parent are checked against the deny list. Returns the updated folder\'s ' +
      '`id`, `displayName`, and `parentFolderId`.',
    parameters: [
      {
        name: 'folderId',
        type: 'string',
        description: 'Mail folder ID to move (from list_folders_mail).',
        required: true,
      },
      {
        name: 'destinationParentFolderId',
        type: 'string',
        description: 'Folder ID of the destination parent.',
        required: true,
      },
      {
        name: 'mailboxId',
        type: 'string',
        description: 'Mailbox user ID or "me" for the current user (default: "me").',
        required: false,
      },
    ],
    endpoint: '/api/mail/folders/{folderId}/move',
    method: 'POST',
  },
  {
    name: 'search_mail',
    description:
      'Search email messages in a mailbox, optionally scoped to a folder. `q` is free text matched as a ' +
      'phrase against subject, body and sender; an email address or domain in `q` is also matched against ' +
      'every participant. To find mail sent TO someone, pass `participant` or `to` — plain text in `q` does ' +
      'not match recipients on folder-scoped searches. Mailbox-wide results are relevance-ranked, not ' +
      'newest-first; use list_messages to enumerate the newest messages. Every response reports which ' +
      'fields were searched and how results are ordered, and a folder scan reports how far back it looked, ' +
      'so an empty result is never proof of absence on its own. Access is blocked if the target folder is ' +
      'on the deny list.',
    parameters: [
      {
        name: 'q',
        type: 'string',
        description: 'Free-text query (phrase match on subject, body, sender). Optional when participant / from / to is given. KQL property syntax (received>=, from:, participants:) is rejected: use since / from / to / participant.',
        required: false,
      },
      {
        name: 'participant',
        type: 'string',
        description: 'Email address, domain or name fragment matched against the sender and every To / Cc / Bcc recipient.',
        required: false,
      },
      {
        name: 'from',
        type: 'string',
        description: 'Email address, domain or name fragment the sender must match.',
        required: false,
      },
      {
        name: 'to',
        type: 'string',
        description: 'Email address, domain or name fragment any To / Cc / Bcc recipient must match (e.g. to check Sent Items for mail to a counterparty).',
        required: false,
      },
      {
        name: 'since',
        type: 'string',
        description: 'ISO-8601 date or date-time; only messages received on or after it (e.g. "2026-09-15").',
        required: false,
      },
      {
        name: 'folderId',
        type: 'string',
        description: 'Scope search to this mail folder ID (optional).',
        required: false,
      },
      {
        name: 'mailboxId',
        type: 'string',
        description: 'Mailbox to search (default: "me").',
        required: false,
      },
      {
        name: 'maxResults',
        type: 'number',
        description: 'Maximum results to return (default 25, max 100).',
        required: false,
      },
    ],
    endpoint: '/api/mail/search',
    method: 'GET',
  },
  {
    name: 'list_messages',
    description:
      'List the newest messages in a mailbox or folder, newest first, deterministically. This is the tool ' +
      'for "what was sent since X" or "the last N messages in Sent Items"; search_mail is relevance-ranked ' +
      'and returns a sample, not the newest N. `truncated: true` means more messages exist beyond `limit`. ' +
      'Access is blocked if the target folder is on the deny list.',
    parameters: [
      {
        name: 'folderId',
        type: 'string',
        description: 'Folder ID to list (omit for the whole mailbox).',
        required: false,
      },
      {
        name: 'since',
        type: 'string',
        description: 'ISO-8601 date or date-time; only messages received on or after it (e.g. "2026-09-15").',
        required: false,
      },
      {
        name: 'mailboxId',
        type: 'string',
        description: 'Mailbox to list (default: "me").',
        required: false,
      },
      {
        name: 'maxResults',
        type: 'number',
        description: 'Maximum results to return (default 25, max 100).',
        required: false,
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'read_message',
    description:
      'Fetch the full content of a specific email message including body, recipients, and ' +
      'metadata. Access is blocked if the message\'s folder is on the deny list.',
    parameters: [
      {
        name: 'messageId',
        type: 'string',
        description: 'Message ID (from search_mail results).',
        required: true,
      },
      {
        name: 'mailboxId',
        type: 'string',
        description: 'Mailbox user ID (default: "me").',
        required: false,
      },
    ],
    endpoint: '/api/mail/messages/{messageId}',
    method: 'GET',
  },
  // ── Calendar ─────────────────────────────────────────────────────────────────
  {
    name: 'list_calendars',
    description:
      'List the calendars available to the current user. Calendars blocked by the deny ' +
      'list (by ID or by name) are excluded from the results.',
    parameters: [],
    endpoint: '/api/calendar/calendars',
    method: 'GET',
  },
  {
    name: 'create_calendar',
    description:
      'Create a new calendar in the signed-in user\'s mailbox — e.g. a "Private" calendar for items that should ' +
      'not appear on a shared/default calendar feed. create_event can then target it via calendarId. Refused if a ' +
      'calendar with that name is on the deny list or the calendar service is in read-only mode. (MCP-native.)',
    parameters: [
      {
        name: 'name',
        type: 'string',
        description: 'Display name for the new calendar (e.g. "Private").',
        required: true,
      },
      {
        name: 'color',
        type: 'string',
        description: 'Optional calendar color preset.',
        required: false,
        enum: ['auto', 'lightBlue', 'lightGreen', 'lightOrange', 'lightGray', 'lightYellow', 'lightTeal', 'lightPink', 'lightBrown', 'lightRed'],
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'list_events',
    description:
      'List events in a calendar. When a date range is given, this uses Graph calendarView, which EXPANDS recurring ' +
      'series into their concrete instances on those days — so it correctly answers "what is on this specific day". ' +
      'If no `calendarId` is provided, the default calendar is used. Access is blocked if the calendar is on the ' +
      'deny list.',
    parameters: [
      {
        name: 'calendarId',
        type: 'string',
        description: 'Calendar ID (from list_calendars). Omit to use the default calendar.',
        required: false,
      },
      {
        name: 'startDateTime',
        type: 'string',
        description: 'Window start — ISO 8601 datetime or a bare date (YYYY-MM-DD). Recurring events are expanded into instances within the window.',
        required: false,
      },
      {
        name: 'endDateTime',
        type: 'string',
        description: 'Window end — ISO 8601 datetime or a bare date. If only one bound is given, the other defaults 7 days away.',
        required: false,
      },
      {
        name: 'timeZone',
        type: 'string',
        description: 'IANA time zone the window and returned times are interpreted in (default: "America/Los_Angeles").',
        required: false,
      },
      {
        name: 'maxResults',
        type: 'number',
        description: 'Maximum number of events to return (default 25, max 50).',
        required: false,
      },
    ],
    endpoint: '/api/calendar/events',
    method: 'GET',
  },
  {
    name: 'get_event',
    description:
      'Fetch full details of a specific calendar event including body, attendees, ' +
      'and location. Access is blocked if the calendar is on the deny list (by ID or name).',
    parameters: [
      {
        name: 'eventId',
        type: 'string',
        description: 'Event ID (from list_events).',
        required: true,
      },
      {
        name: 'calendarId',
        type: 'string',
        description: 'Calendar ID. Omit to look up the event in the default calendar.',
        required: false,
      },
    ],
    endpoint: '/api/calendar/events/{eventId}',
    method: 'GET',
  },
  {
    name: 'create_event',
    description:
      'Create a new calendar event. When other attendees are involved, call find_meeting_times or get_schedule first to ' +
      'pick a slot everyone is free for, and list_rooms to choose a real conference room. Set isOnlineMeeting=true to ' +
      'attach a Microsoft Teams meeting — the response then includes onlineMeetingUrl (the Teams join link). ' +
      'Blocked if the calendar is on the deny list or the calendar service is in read-only mode.',
    parameters: [
      {
        name: 'subject',
        type: 'string',
        description: 'Event title.',
        required: true,
      },
      {
        name: 'start',
        type: 'string',
        description: 'Start datetime in ISO 8601 format (e.g. "2026-06-20T09:00:00").',
        required: true,
      },
      {
        name: 'end',
        type: 'string',
        description: 'End datetime in ISO 8601 format.',
        required: true,
      },
      {
        name: 'timeZone',
        type: 'string',
        description: 'IANA time zone name (e.g. "America/Los_Angeles"). Default: "America/New_York".',
        required: false,
      },
      {
        name: 'location',
        type: 'string',
        description: 'Physical location or meeting URL.',
        required: false,
      },
      {
        name: 'body',
        type: 'string',
        description: 'Event description or notes.',
        required: false,
      },
      {
        name: 'bodyType',
        type: 'string',
        description: '"text" (default) or "html".',
        required: false,
        enum: ['text', 'html'],
      },
      {
        name: 'isAllDay',
        type: 'boolean',
        description: 'Set to true for an all-day event.',
        required: false,
      },
      {
        name: 'showAs',
        type: 'string',
        description:
          'Free/busy status shown to others. Omit to let Graph default it (busy for timed events, free for all-day).',
        required: false,
        enum: ['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'],
      },
      {
        name: 'isOnlineMeeting',
        type: 'boolean',
        description: 'Set to true to attach a Teams online meeting. The response includes onlineMeetingUrl (the join link).',
        required: false,
      },
      {
        name: 'onlineMeetingProvider',
        type: 'string',
        description: 'Online meeting provider when isOnlineMeeting is true (default: teamsForBusiness).',
        required: false,
        enum: ['teamsForBusiness', 'skypeForBusiness', 'skypeForConsumer'],
      },
      {
        name: 'calendarId',
        type: 'string',
        description: 'Calendar ID. Omit to create in the default calendar.',
        required: false,
      },
    ],
    endpoint: '/api/calendar/events',
    method: 'POST',
  },
  {
    name: 'update_event',
    description:
      'Update an existing calendar event. Only the fields you provide are changed. ' +
      'Blocked if the calendar is on the deny list or the calendar service is in read-only mode.',
    parameters: [
      {
        name: 'eventId',
        type: 'string',
        description: 'Event ID to update.',
        required: true,
      },
      {
        name: 'calendarId',
        type: 'string',
        description: 'Calendar ID. Omit to look up in the default calendar.',
        required: false,
      },
      {
        name: 'subject',
        type: 'string',
        description: 'New event title.',
        required: false,
      },
      {
        name: 'start',
        type: 'string',
        description: 'New start datetime (ISO 8601).',
        required: false,
      },
      {
        name: 'end',
        type: 'string',
        description: 'New end datetime (ISO 8601).',
        required: false,
      },
      {
        name: 'timeZone',
        type: 'string',
        description: 'IANA time zone name for updated start/end times.',
        required: false,
      },
      {
        name: 'location',
        type: 'string',
        description: 'Updated location.',
        required: false,
      },
      {
        name: 'body',
        type: 'string',
        description: 'Updated description.',
        required: false,
      },
      {
        name: 'bodyType',
        type: 'string',
        description: '"text" or "html".',
        required: false,
        enum: ['text', 'html'],
      },
      {
        name: 'isAllDay',
        type: 'boolean',
        description: 'Whether the event spans the full day.',
        required: false,
      },
      {
        name: 'showAs',
        type: 'string',
        description: 'Free/busy status shown to others.',
        required: false,
        enum: ['free', 'tentative', 'busy', 'oof', 'workingElsewhere', 'unknown'],
      },
    ],
    endpoint: '/api/calendar/events/{eventId}',
    method: 'PATCH',
  },
  {
    name: 'delete_event',
    description:
      'Delete a calendar event. Blocked if the calendar is on the deny list or the calendar service is in read-only mode.',
    parameters: [
      {
        name: 'eventId',
        type: 'string',
        description: 'Event ID to delete.',
        required: true,
      },
      {
        name: 'calendarId',
        type: 'string',
        description: 'Calendar ID. Omit to delete from the default calendar.',
        required: false,
      },
    ],
    endpoint: '/api/calendar/events/{eventId}',
    method: 'DELETE',
  },
  {
    name: 'move_event',
    description:
      'Move a calendar event to another calendar. Microsoft Graph has no native event-move, so this is a ' +
      'COPY-THEN-DELETE and is NOT atomic: the moved event gets a new id, and the original is deleted only after ' +
      'the copy is confirmed. It preserves body, location, categories, sensitivity, showAs, importance, reminders, ' +
      'recurrence (series master copied as a series, not expanded), and file attachments. If the event has ' +
      'attendees the move sends fresh invites plus a cancellation, so it is REFUSED unless force=true; it is always ' +
      'refused when you are not the organizer or the event is a single occurrence of a series. Returns the new ' +
      'event id. (MCP-native.)',
    parameters: [
      {
        name: 'eventId',
        type: 'string',
        description: 'ID of the event to move (from list_events / get_event).',
        required: true,
      },
      {
        name: 'targetCalendarId',
        type: 'string',
        description: 'Destination calendar ID (from list_calendars / create_calendar).',
        required: true,
      },
      {
        name: 'calendarId',
        type: 'string',
        description: 'Source calendar ID. Omit to move from the default calendar.',
        required: false,
      },
      {
        name: 'force',
        type: 'boolean',
        description: 'Required (true) to move an event that has attendees — acknowledges fresh invites and a cancellation will be sent. Default false.',
        required: false,
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'respond_to_event',
    description:
      'Respond to a calendar event invitation — accept, tentatively accept, or decline it. Takes either the event ID ' +
      '(from list_events) or the meeting-invite message ID from the inbox (from search_mail / read_message); a message ' +
      'ID is resolved to its underlying event automatically. Set sendResponse=false to update the RSVP without ' +
      'notifying the organizer. Blocked if the calendar is on the deny list or the calendar service is in read-only mode.',
    parameters: [
      {
        name: 'messageOrEventId',
        type: 'string',
        description: 'Event ID (from list_events) or meeting-invite message ID (from search_mail).',
        required: true,
      },
      {
        name: 'response',
        type: 'string',
        description: 'How to respond: accept, tentative, or decline.',
        required: true,
        enum: ['accept', 'tentative', 'decline'],
      },
      {
        name: 'comment',
        type: 'string',
        description: 'Optional note sent to the organizer with the response.',
        required: false,
      },
      {
        name: 'sendResponse',
        type: 'boolean',
        description: 'Send the response to the organizer (default true).',
        required: false,
      },
      {
        name: 'mailboxId',
        type: 'string',
        description: 'Mailbox to act in. Omit for the signed-in user.',
        required: false,
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'get_schedule',
    description:
      'Look up free/busy availability for one or more people over a time window. Call this before scheduling a meeting ' +
      'with other attendees so you propose a slot everyone is free for. Returns each person\'s busy blocks and working ' +
      'hours. Only works for users whose free/busy the caller may see in the tenant.',
    parameters: [
      {
        name: 'schedules',
        type: 'array',
        items: 'string',
        description: 'Email addresses (SMTP) of the people/rooms to check availability for.',
        required: true,
      },
      {
        name: 'startDateTime',
        type: 'string',
        description: 'Window start in ISO 8601 (e.g. "2026-07-06T09:00:00").',
        required: true,
      },
      {
        name: 'endDateTime',
        type: 'string',
        description: 'Window end in ISO 8601.',
        required: true,
      },
      {
        name: 'timeZone',
        type: 'string',
        description: 'IANA time zone for the window and results (default: "America/Los_Angeles").',
        required: false,
      },
      {
        name: 'availabilityViewInterval',
        type: 'number',
        description: 'Granularity of the availabilityView string, in minutes (default 30).',
        required: false,
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'find_meeting_times',
    description:
      'Suggest concrete meeting times that work for a set of attendees, honoring each attendee\'s working hours and time ' +
      'zone from their MailboxSettings. Prefer this over get_schedule when you want ranked, ready-to-book slots. Returns ' +
      'suggestions with a confidence score and per-attendee availability.',
    parameters: [
      {
        name: 'attendees',
        type: 'array',
        items: 'string',
        description: 'Required attendee email addresses.',
        required: true,
      },
      {
        name: 'meetingDurationMinutes',
        type: 'number',
        description: 'Desired meeting length in minutes (default 30).',
        required: false,
      },
      {
        name: 'startDateTime',
        type: 'string',
        description: 'Earliest acceptable start, ISO 8601 (optional).',
        required: false,
      },
      {
        name: 'endDateTime',
        type: 'string',
        description: 'Latest acceptable end, ISO 8601 (optional).',
        required: false,
      },
      {
        name: 'timeZone',
        type: 'string',
        description: 'IANA time zone for the time window (default: "America/Los_Angeles").',
        required: false,
      },
      {
        name: 'maxCandidates',
        type: 'number',
        description: 'Maximum number of suggestions to return (default 10).',
        required: false,
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'list_rooms',
    description:
      'List the conference room mailboxes (places) in the tenant. Use this to pick a real, bookable room for a meeting ' +
      'location instead of leaving it blank — the returned emailAddress can be added as an attendee to book the room.',
    parameters: [
      {
        name: 'maxResults',
        type: 'number',
        description: 'Maximum rooms to return (default 100).',
        required: false,
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },

  // ── Teams ────────────────────────────────────────────────────────────────────
  {
    name: 'list_teams',
    description:
      'List the Microsoft Teams the calling user is a member of. Use this to discover the teamId needed for ' +
      'list_channels or send_channel_message. Teams on the deny list are excluded.',
    parameters: [],
    endpoint: '/api/teams/teams',
    method: 'GET',
  },
  {
    name: 'list_channels',
    description:
      'List the channels in a Team. Use this after list_teams to find the channelId for send_channel_message. ' +
      'Channels on the deny list are excluded.',
    parameters: [
      {
        name: 'teamId',
        type: 'string',
        description: 'Team ID (from list_teams).',
        required: true,
      },
    ],
    endpoint: '/api/teams/teams/{teamId}/channels',
    method: 'GET',
  },
  {
    name: 'send_chat_message',
    description:
      'Send a message to an existing Teams 1:1 or group chat. The chat must already exist — this does not create new ' +
      'chats. Blocked if the chat is on the deny list.',
    parameters: [
      {
        name: 'chatId',
        type: 'string',
        description: 'Chat ID (Graph chat thread ID).',
        required: true,
      },
      {
        name: 'content',
        type: 'string',
        description: 'Message body.',
        required: true,
      },
      {
        name: 'contentType',
        type: 'string',
        description: '"text" (default) or "html".',
        required: false,
        enum: ['text', 'html'],
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
  {
    name: 'send_channel_message',
    description:
      'Post a message to a Teams channel. Use list_teams then list_channels to resolve the IDs. Good for "ping the team" ' +
      'style notifications. Blocked if the team or channel is on the deny list.',
    parameters: [
      {
        name: 'teamId',
        type: 'string',
        description: 'Team ID (from list_teams).',
        required: true,
      },
      {
        name: 'channelId',
        type: 'string',
        description: 'Channel ID (from list_channels).',
        required: true,
      },
      {
        name: 'content',
        type: 'string',
        description: 'Message body.',
        required: true,
      },
      {
        name: 'contentType',
        type: 'string',
        description: '"text" (default) or "html".',
        required: false,
        enum: ['text', 'html'],
      },
    ],
    endpoint: '/api/mcp',
    method: 'POST',
  },
];

/** Returns the full MCP-compatible manifest object. */
export function getManifest(): object {
  return {
    schema_version: '1.0',
    name: process.env.MCP_INSTANCE_NAME ?? 'm365-mcp',
    description:
      'Controlled access to Microsoft 365 SharePoint sites, Exchange/Outlook mailboxes, and ' +
      'Calendar via Microsoft Graph API. Features a two-tier deny list (global admin + per-user) ' +
      'for fine-grained access control over which folders, mail folders, and calendars AI ' +
      'agents may read.',
    auth: {
      type: 'oauth2',
      flow: 'authorization_code',
      login_url: '/api/auth/login',
      callback_url: '/api/auth/callback',
    },
    tools: TOOLS.map((tool) => ({
      name: tool.name,
      description: tool.description,
      inputSchema: {
        type: 'object',
        properties: Object.fromEntries(
          tool.parameters.map((p) => [
            p.name,
            {
              type: p.type,
              description: p.description,
              ...(p.itemsSchema ? { items: p.itemsSchema } : p.items ? { items: { type: p.items } } : {}),
              ...(p.enum ? { enum: p.enum } : {}),
            },
          ])
        ),
        required: tool.parameters.filter((p) => p.required).map((p) => p.name),
      },
    })),
  };
}
