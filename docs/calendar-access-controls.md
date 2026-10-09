# Calendar access controls

The Calendar tools (`list_calendars`, `create_calendar`, `list_events`,
`get_event`, `create_event`, `update_event`, `delete_event`, `move_event`,
`respond_to_event`) ship with two admin-managed access controls.
Both are enforced **identically** on the HTTP routes (`/api/calendar/*`) and the
MCP `tools/call` path, so neither surface can be used to bypass the other — and
no bypass is possible via the built-in M365 connector once it is disabled, because every path runs through the same enforcement.

## 1. Deny list — by calendar ID *or* name

The two-tier deny list (global admin + per-user, `GlobalDenyList` /
`UserDenyList` tables, type `calendar`) blocks a calendar whether the entry was
recorded as the opaque Graph **calendar ID** or the human-readable **display
name**. Blocking a shared/sensitive calendar by name (e.g. `HR`, `Executive`,
`Board`) is the common case, since admins rarely know a calendar's GUID.

Enforcement covers every calendar operation:

- `list_calendars` — deny-listed calendars are excluded from the results.
- `list_events`, `get_event`, `create_event`, `update_event`, `delete_event` —
  the effective calendar (explicit `calendarId`, or the default calendar reached
  via `/me/events`) is checked by ID and by name; a match returns `403`
  (HTTP) / an error result (MCP) before any Graph call.
- `create_calendar` (MCP-native, no HTTP route) — the requested calendar **name**
  is checked against the deny list before creation, so an admin who blocks a
  calendar by name cannot have one recreated under that name.
- `move_event` (MCP-native, no HTTP route) — checks **both** ends: the source
  calendar (explicit `calendarId` or the default) and the target calendar are
  each checked by ID and by name; a deny match on either refuses the move before
  any data is touched. See "move_event is copy-then-delete" below.
- `respond_to_event` (MCP-native, no HTTP route) — always acts on the target
  mailbox's default calendar, which is checked by ID and by name the same way
  (the signed-in user's via `checkCalendarAccess`, an explicit `mailboxId`'s
  via that mailbox's default calendar) before any Graph call.

Add an entry in the `/admin` UI (Deny List Manager, Calendar type) or via the
existing deny-list admin endpoints.

## 2. Read-only mode — `readOnlyServices`

A tenant-level setting (`serviceSettings` table, row `readOnlyServices`) lists
services that stay **readable but not writable**. When `calendar` is in the set:

| Tool | Read-only mode |
|---|---|
| `list_calendars`, `list_events`, `get_event` | allowed |
| `create_calendar`, `create_event`, `update_event`, `delete_event`, `move_event`, `respond_to_event` | refused (`403` / error) and hidden from MCP `tools/list` |

A read-only service does not advertise its write tools. MCP `tools/list` drops
every `WRITE_TOOLS` entry whose service is read-only, because a
connector risk review scores what a connector exposes, not only what it
refuses. The `tools/call` refusal stays in place as the backstop, so a client
that calls a hidden tool by name still gets the read-only error.

Read-only is independent of `enabledServices`: a service must be enabled to be
usable at all, and read-only further restricts an enabled service to reads. It
is per-service, so making calendar read-only does not affect mail, SharePoint,
etc. Default is empty, so existing tenants are unaffected until an admin opts in.

Manage it through the `/api/manage/services` endpoint (Global Admin only; the admin API takes the admin UI's browser session, not an MCP client token, see [Admin API sessions](operations-runbook.md#admin-api-sessions)):

```bash
# Read current config
GET /api/manage/services
# → { "enabledServices": [...], "readOnlyServices": ["calendar"] }

# Make calendar read-only (leaves enabledServices untouched)
POST /api/manage/services   { "readOnlyServices": ["calendar"] }

# Update both lists at once
POST /api/manage/services   { "enabledServices": ["mail","calendar"], "readOnlyServices": ["calendar"] }
```

## move_event is copy-then-delete, not atomic

Microsoft Graph has **no native move for calendar events** — unlike mail, which
has `POST /me/messages/{id}/move`. The Outlook UI's drag-and-drop between
calendars uses EWS `MoveItem`, which Graph does not surface. So `move_event`
copies the event into the target calendar and deletes the original. This is
**not atomic**, and the copy is lossy in the general case, so the tool is
deliberately conservative:

- **Faithful copy.** `create_event` alone only carries subject/start/end/body, so
  `move_event` reads the full source event and re-creates it with body, location,
  categories, sensitivity, showAs, importance, reminder settings, recurrence, and
  file attachments preserved. Non-file attachments (item / reference) cannot be
  re-posted as bytes and are reported back in the response `warnings` array rather
  than being silently dropped.
- **Recurring series.** The series **master** is copied as a series (its
  `recurrence` rule is carried over). A single occurrence / exception is
  **refused** with a pointer to move the master instead, so a recurring series is
  never silently flattened into one event.
- **Attendees → real outbound mail.** Re-creating an event that has attendees
  sends **fresh invitations** to everyone, and deleting the original sends
  **cancellations**. Because that is outbound mail to real people as a side
  effect of a "move," `move_event` **refuses by default** when the event has
  attendees and requires an explicit `force=true`.
- **Organizer only.** Moving an event you did not organize has the same blast
  radius as moving someone else's meeting, so a non-organizer move is **always
  refused** (no `force` override).
- **Create-before-delete.** The event is created in the target and confirmed
  (a new event id is returned) **before** the original is deleted — never the
  other way round — so a failed copy leaves the source intact.
- **New identity.** The moved event has a **new** event id and `webLink`; any
  saved link or id to the original stops resolving.

## Example rollout ( →)

1. **Deploy** — merge this change to canonical `standardgauge/m365-mcp`; the
   the deploy workflow
   roll `your-mcp-host.example.com` to the new revision. No PR is opened against the
   deploy fork (it is a sync target).
2. **Enable calendar for the Example tenant** — as a Global Admin on
   `your-mcp-host.example.com`, add `calendar` to `enabledServices` (admin UI, or
   `POST /api/manage/services`). This is what surfaces the calendar tools to
   Example Claude — tool exposure is a runtime tenant setting, not a client config
   file. (Example's m365-mcp repo holds specs/prototypes only; it has no MCP client
   config to edit.)
3. **Apply Example's calendar policy** — set `readOnlyServices` and any calendar
   deny-list entries per Example's decision (e.g. read-only until write access is
   signed off; deny shared/exec calendars by name).
4. **Disable the built-in connector** — safe once the tools above
   are live, since calendar access now runs entirely through this MCP's
   enforcement.
