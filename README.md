# M365 MCP Server

A Microsoft 365 [MCP](https://modelcontextprotocol.io) server that gives Claude and other AI agents read/write access to Outlook Mail, Calendar, SharePoint, OneDrive, Contacts, OneNote, and Teams through the Microsoft Graph API.

Every user signs in individually with delegated OAuth. There is no service account and no application-level Graph permission, so the server can only ever reach what the signed-in user could reach themselves. On top of that, a two-tier deny list and per-service read-only switches let an administrator carve out what AI is allowed to touch.

Built on Azure Functions v4 (TypeScript, Node 20), deployed as an Azure Container App, with a React admin UI for managing access controls.

**License:** AGPL-3.0. See [License](#license).

---

## Contents

- [How it works](#how-it-works)
- [Access controls](#access-controls)
- [Deploying your own instance](#deploying-your-own-instance)
- [Entra ID app registration and permissions](#entra-id-app-registration-and-permissions)
- [Configuration reference](#configuration-reference)
- [Connecting a client](#connecting-a-client)
- [Tool reference](#tool-reference)
- [Local development](#local-development)
- [License](#license)

---

## How it works

The server is a single HTTP service. Clients speak MCP to it over JSON-RPC at `/api/mcp`; there is no local bridge process to install and no code running on the user's machine beyond a thin stdio-to-HTTP shim.

```
MCP client (Claude Desktop, Claude Code, any MCP client)
      │  JSON-RPC over Streamable HTTP, with an x-user-id header
      ▼
Azure Container App
├── /api/mcp                → MCP protocol endpoint (tools/list, tools/call)
├── /api/auth/login         → MSAL OAuth sign-in
├── /api/auth/callback      → OAuth redirect target
├── /api/mail/*             → Mail and Exchange
├── /api/calendar/*         → Calendar and scheduling
├── /api/sharepoint/*       → SharePoint
├── /api/onedrive/*         → OneDrive
├── /api/contacts/*         → Contacts
├── /api/onenote/*          → OneNote
└── /admin                  → React admin UI
      │
      ▼
Azure Table Storage
├── mcpSessions             → Persistent user sessions
├── mcpMsalCache            → MSAL token cache, including refresh tokens
├── GlobalDenyList          → Administrator-managed deny list
├── UserDenyList            → Per-user deny list
├── serviceSettings         → Service enablement and read-only flags
├── allowedSites            → SharePoint site allowlist
├── UserEmailSettings       → Per-user email output mode (draft or send)
└── EmailOutputModePolicy   → Admin-enforced draft mode, tenant-wide or per user
      │
      ▼
Microsoft Graph
```

### The authentication model

Each user authenticates once through the standard Microsoft sign-in page. The resulting refresh token is held in the MSAL cache in Azure Table Storage, so sessions survive container restarts and redeploys without forcing everyone to sign in again.

Because every token is **delegated**, the server's reach is bounded by the user's own permissions. A user who cannot open a SharePoint site cannot reach it through this server either. There are no application (app-only) Graph permissions anywhere in the registration, which is what rules out a back-door service account with tenant-wide access.

Calls are identified by an `x-user-id` header, which maps to a stored session. A request with no valid session gets no Graph access.

### Deny lists are enforced on both surfaces

Every tool call passes through the deny list before it reaches Graph, and the same check runs on the HTTP routes and on the MCP `tools/call` path. Neither surface can be used to bypass the other. This matters because the REST routes and the MCP endpoint are two doors into the same building; enforcing on one only would be theatre.

---

## Access controls

**Two-tier deny list.** Tier 1 is global and administrator-managed: folders, paths, and calendars blocked for everyone. Tier 2 is per-user and self-managed, so an individual can hide their own folders from AI without asking an administrator. Calendar entries match by calendar ID *or* display name, so a shared calendar such as "HR" or "Executive" can be blocked by name across every calendar tool.

**Per-service read-only mode.** The `readOnlyServices` setting keeps a service readable while refusing all of its write tools. Those write tools are also hidden from the MCP tool list, so a read-only service advertises only what it will actually do. The use case it was built for: let an agent read a calendar for scheduling context without letting it create, modify, or delete events.

**Enforced draft mode.** Every user starts in draft mode, where AI-composed email lands in Drafts for a person to send. The mode is self-service, and that includes the agent: `set_email_output_mode` is a tool, so an injected prompt could switch to send mode and then send. A Global Admin can enforce draft mode for the tenant or for one user. While enforced, the effective mode is draft whatever the user chose, `send_draft` refuses, and both mode-change paths (the MCP tool and `POST /api/mail/settings`) refuse and log the attempt as denied. Only the admin UI lifts it.

**SharePoint site allowlist.** Once the list has any entry, sites must be explicitly allowed rather than blocked, so a newly created site is not reachable by default. An empty list allows every site, so add at least one site before users sign in if you want a closed posture.

**Default deny lists per deployment.** `DEFAULT_MAIL_DENY_FOLDERS` and `DEFAULT_SHAREPOINT_DENY_PATHS` apply on top of the table-managed lists, which gives a new instance a safe baseline before an administrator has configured anything.

---

## Deploying your own instance

You need an Azure subscription and an Entra ID tenant you can register an application in. The instance runs entirely inside your own tenant; nothing routes through a third party.

### 1. Azure resources

Create a resource group, an Azure Container Registry, a storage account, and a Container App. Naming is yours to choose; record the names because they become the environment block of your deploy workflow.

The Bicep templates under `infra/` provision this. `infra/main.bicep` is the full-stack template (storage account and tables, registry, identity, Log Analytics, Application Insights, environment, Container App); `infra/container-app.bicep` is the same without the storage account, for a deployment that brings its own. Entra setup, in order, with verification checkpoints: [`docs/entra-setup.md`](docs/entra-setup.md).

### 2. Entra app registration

See [Entra ID app registration and permissions](#entra-id-app-registration-and-permissions) below. This is the step with the most detail and the one most likely to bite you, so read it rather than skimming.

### 3. Build and deploy

```bash
# Client and tenant IDs are build args because the admin UI needs them at build time.
docker build --platform linux/amd64 \
  --build-arg AZURE_CLIENT_ID=<client-id> \
  --build-arg AZURE_TENANT_ID=<tenant-id> \
  -t <acr-server>/<app-name>:<tag> .

docker push <acr-server>/<app-name>:<tag>

az containerapp update \
  --name <app-name> \
  --resource-group <resource-group> \
  --image <acr-server>/<app-name>:<tag>
```

### 4. Set secrets

Container App environment variables **must** use secret references rather than plain-text values, so the client secret and storage connection string never appear in the revision definition.

```bash
az containerapp secret set --name <app> --resource-group <rg> \
  --secrets azure-client-id="<value>" azure-client-secret="<value>" ...
```

Secrets are picked up by a new revision, so follow this with an `az containerapp update --image ...`.

### 5. Verify

Sign in at `https://<your-host>/api/auth/login`, confirm a session appears at `/api/manage/sessions`, and exercise one tool per service you enabled. A deployment that authenticates but returns nothing usually means consent is incomplete; see the troubleshooting note in the permissions section.

---

## Entra ID app registration and permissions

### Create the registration

In the Azure portal, go to **Microsoft Entra ID → App registrations → New registration**.

- **Name:** anything you like, for example `m365-mcp`.
- **Supported account types:** *Accounts in this organizational directory only* for a single-tenant instance. Choose multi-tenant only if you intend to serve users from other tenants through admin consent.
- **Redirect URI:** leave blank for now.

Record the **Application (client) ID** and **Directory (tenant) ID**. These become `AZURE_CLIENT_ID` and `AZURE_TENANT_ID`.

### Redirect URIs

Under **Authentication → Add a platform → Web**, add one redirect URI per environment:

| Environment | Redirect URI |
|---|---|
| Local development | `http://localhost:7071/api/auth/callback` |
| Deployed instance | `https://<your-host>/api/auth/callback` |

Also enable **ID tokens** under *Implicit grant and hybrid flows*, and set the logout URL to `https://<your-host>/api/auth/logout`.

### Client secret

Under **Certificates & secrets → Client secrets → New client secret**, create a secret with a 24-month expiry and copy the **Value** immediately; it is never shown again. This becomes `AZURE_CLIENT_SECRET`. Diarise the rotation, because expiry takes the whole instance down at once.

### Delegated Graph permissions

All permissions are **delegated**. There are no application permissions, by design.

Add these under **API permissions → Add a permission → Microsoft Graph → Delegated permissions**, then click **Grant admin consent** once for the whole set.

| Permission | Backs |
|---|---|
| `openid`, `profile`, `offline_access` | Sign-in and refresh tokens |
| `User.Read` | The signed-in user's own profile |
| `User.ReadBasic.All` | Resolving other users' basic profiles |
| `Directory.Read.All` | Detecting whether the signed-in user is a Global Administrator. The admin view reads the user's directory roles via `/me/transitiveMemberOf`. This is the one tool-independent scope carried in `GRAPH_SCOPES` (see the note below). |
| `Mail.ReadWrite` | Mail read, draft, move, delete |
| `Mail.ReadBasic` | Enumerating delegated mailboxes |
| `Mail.Send` | Sending mail and dispatching drafts |
| `MailboxSettings.Read` | Working hours and time zones for scheduling |
| `Calendars.ReadWrite` | Calendar read and write |
| `Calendars.Read.Shared` | Colleagues' free/busy for `get_schedule` |
| `Place.Read.All` | Conference room mailboxes for `list_rooms` |
| `Sites.ReadWrite.All` | SharePoint sites |
| `Files.ReadWrite.All` | OneDrive and SharePoint files |
| `Contacts.ReadWrite` | Contacts |
| `Notes.ReadWrite.All` | OneNote |
| `Team.ReadBasic.All` | Listing Teams |
| `Channel.ReadBasic.All` | Listing channels |
| `ChannelMessage.Read.All` | Reading channel messages |
| `ChannelMessage.Send` | Posting to a channel |
| `ChatMessage.Send` | Sending 1:1 and group chat |

`Place.Read.All`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, and `Directory.Read.All` require administrator consent. Granting consent once for the tenant covers the rest, so individual users see no consent prompt on first sign-in.

Trim this list to the services you actually intend to enable. If you are not deploying the Teams tools, leave the four Teams permissions off the registration entirely; a permission not granted is one that cannot be misused.

### The part that catches people: requested scopes are a subset of granted scopes

`GRAPH_SCOPES` in `src/services/graphClient.ts` is **not** the list of permissions the tools use. It is the much shorter list MSAL names when requesting a token.

This works because on the Entra v2.0 endpoint an access token carries *every* delegated permission consented for that resource, not only the ones named in the request. Calendar, OneNote, and Teams all rely on this: none of their scopes appears in `GRAPH_SCOPES`, and the tools work anyway because the permissions are granted on the app registration.

Two consequences, both of which have caused real outages:

> **Do not add tool permissions to `GRAPH_SCOPES`.** Adding a scope the tenant has not consented to makes `acquireTokenSilent` request it and throw `interaction_required`, which breaks *every* tool call for *every* already-signed-in user. New permissions belong on the app registration, not in this array. Once consent lands, they flow into the token on the next silent refresh with no re-login.

> **Everything in `GRAPH_SCOPES` must be consented before first sign-in.** The reverse of the same rule. If a scope is requested but not granted, authentication fails for everyone rather than degrading to a smaller tool set.

At the time of writing `GRAPH_SCOPES` contains `Sites.ReadWrite.All`, `Files.ReadWrite.All`, `Mail.ReadWrite`, `Mail.ReadBasic`, `User.Read`, `Contacts.ReadWrite`, and `Directory.Read.All`. This list and the permission table above are checked against the source array by `src/__tests__/scopeDocDrift.test.ts`, which fails CI if any scope in `GRAPH_SCOPES` is missing from the table or from this sentence, so the two cannot silently drift. Every scope named here therefore requires admin consent before first sign-in, `Directory.Read.All` included.

### Multi-tenant instances

A multi-tenant registration lets users from other tenants sign in after their own administrator consents:

```
https://login.microsoftonline.com/organizations/adminconsent?client_id=<your-client-id>
```

Single-tenant instances need no such URL. Prefer single-tenant unless you have a specific reason not to; it is the tighter boundary.

---

## Configuration reference

Copy `.env.example` to `.env` for local work, or set these as Container App secret references in production.

| Variable | Required | Purpose |
|---|---|---|
| `AZURE_CLIENT_ID` | Yes | Entra application (client) ID |
| `AZURE_CLIENT_SECRET` | Yes | Entra client secret |
| `AZURE_TENANT_ID` | Yes | Entra directory (tenant) ID |
| `OAUTH_REDIRECT_URI` | Yes | Must match a registered redirect URI exactly |
| `FRONTEND_URL` | Yes | Base URL of the admin UI, used for post-login redirects |
| `AZURE_STORAGE_CONNECTION_STRING` | Production | Table Storage account. Unset locally falls back to the Azurite emulator |
| `AUDIT_LOG_RETENTION_DAYS` | No | Days of `auditLog` rows to keep. Older rows are purged about once a day while the server is in use. Defaults to `365`; `0` keeps every row. See the operations runbook |
| `DEFAULT_MAIL_DENY_FOLDERS` | No | Comma-separated mail folder names always denied, matched by display name, case-insensitively |
| `DEFAULT_SHAREPOINT_DENY_PATHS` | No | Comma-separated SharePoint path prefixes always denied |
| `MCP_INSTANCE_NAME` | No | Display name shown in the admin UI, the `initialize` response, and the generated extension bundle (also derives the bundle's slug and keychain service name). Defaults to `M365 MCP` |
| `MCP_SESSION_HMAC_KEY` | Production | 64-hex-char key that hashes session tokens at rest. Generate with `openssl rand -hex 32`; never reuse across tenants |
| `MCP_DATA_ENCRYPTION_KEY` | Production | 64-hex-char AES-256-GCM key for access tokens and the MSAL cache at rest. Same generation rule. Rotating it signs every user out |
| `APPLICATIONINSIGHTS_CONNECTION_STRING` | No | Forwards console output and exceptions to the tenant's own Application Insights resource |
| `GIT_SHA` | Set by the build | Reported by `GET /health` so a probe can assert which commit is running |

---

## Connecting a client

Each instance serves its own installers. They sign you in through the browser, save a small local shim, and write the Claude Desktop and Claude Code configs:

```bash
curl -fsSL https://<your-host>/install.sh | bash        # macOS
```

```powershell
iwr -useb https://<your-host>/install.ps1 | iex         # Windows
```

Claude Desktop users can install the extension from `https://<your-host>/install` instead.

The shim (`src/install/m365-mcp-shim.js`, served at `/install/m365-mcp-shim.js`) is the stdio process the MCP client launches. It forwards JSON-RPC to `/api/mcp` with your session token, and needs only Node.js 18 or later. It also lets `create_draft` and `send_mail` attachments, and `write_onedrive_file`, name a local file by path instead of carrying base64, so the file never passes through the model's context. Local reads are limited to `~/Downloads`, `~/Documents` and the OneDrive sync folders, refuse hidden files, and cap at 10 MB. Set `M365_MCP_ATTACH_ROOTS` in the entry's `env` to change the folders.

To configure a client by hand, download the shim and point the client at it:

```json
{
  "mcpServers": {
    "m365": {
      "command": "node",
      "args": [
        "/path/to/m365-mcp-shim.js",
        "--streamableHttp", "https://<your-host>/api/mcp",
        "--header", "Authorization:Bearer <session-token>"
      ]
    }
  }
}
```

---

## Tool reference

### Mail and Exchange (17 tools)

| Tool | Method | Description |
|---|---|---|
| `list_mailboxes` | Read | List accessible mailboxes |
| `list_folders_mail` | Read | List mail folders, deny-list filtered |
| `create_mail_folder` | Write | Create a mail folder, top-level or nested |
| `rename_mail_folder` | Write | Rename a mail folder |
| `move_mail_folder` | Write | Move a mail folder under a new parent |
| `search_mail` | Read | Search messages by text, or by `participant` / `from` / `to` address. Reports which fields were searched and how results are ordered |
| `list_messages` | Read | Newest N messages in a mailbox or folder, newest first, optionally `since` a date |
| `read_message` | Read | Fetch full message content |
| `get_attachments` | Read | List or download attachments |
| `create_draft` | Write | Create a new draft, never sends. Attachments inline or from OneDrive/SharePoint by item ID |
| `reply_to_message` | Write | Draft a threaded reply to the sender |
| `reply_all_to_message` | Write | Draft a threaded reply to everyone |
| `forward_message` | Write | Draft a threaded forward |
| `send_mail` | Write | Compose and send, or save to Drafts, per the user's output mode |
| `send_draft` | Write | Send an existing draft |
| `move_message` | Write | Move a message between folders |
| `delete_message` | Delete | Delete a message |

**Replying properly matters more than it looks.** A reply built with `create_draft` carries an `RE:` subject and pasted-in text, but no quoted history and, more damagingly, no `In-Reply-To` or `References` headers. The recipient's mail client files it as a brand new conversation instead of collapsing it into the thread they started. Nothing about this is visible from the sending side, since your own Sent folder looks correct either way.

`reply_to_message`, `reply_all_to_message`, and `forward_message` wrap Graph's `createReply`, `createReplyAll`, and `createForward`. Each returns a draft with the quoted original already below your text and the threading headers set correctly. Use them, then `send_draft`.

**Searching for who a message went to is not a text search.** `search_mail` with only `q` matches subject, body and sender. In Sent Items the sender is always you, so a counterparty's address in `q` finds nothing on a folder-scoped search, and Graph reports that empty result with HTTP 200. Pass `participant` or `to` instead. Mailbox-wide search is relevance-ranked, so the first N results are not the newest N; use `list_messages` for "what went out since Tuesday". Both tools report `strategy`, `ordering`, `searchedFields` and, on a folder scan, `scanHorizon` and `scanComplete`, so a caller can tell an exhaustive miss from a bounded one.

`get_email_output_mode` and `set_email_output_mode` control whether `send_mail` delivers immediately or lands in Drafts for review. When an administrator enforces draft mode, `get_email_output_mode` reports `enforced: true` and `set_email_output_mode` refuses; see [Access controls](#access-controls).

### Calendar (8 tools)

| Tool | Method | Description |
|---|---|---|
| `list_calendars` | Read | List the user's calendars |
| `create_calendar` | Write | Create a calendar |
| `list_events` | Read | List events, filterable by date range |
| `create_event` | Write | Create an event with attendees |
| `update_event` | Write | Update event details |
| `delete_event` | Delete | Delete an event |
| `move_event` | Write | Move an event between calendars by copy-then-delete. Not atomic; refuses attendee events without `force`, and refuses non-organizer and single-occurrence events |
| `respond_to_event` | Write | Accept, tentatively accept, or decline an invite |

Scheduling helpers `get_schedule`, `find_meeting_times`, and `list_rooms` round out the calendar surface.

### SharePoint (9 tools)

| Tool | Method | Description |
|---|---|---|
| `list_sites` | Read | Enumerate accessible sites |
| `list_folders` | Read | List folder children, with files when `includeFiles=true`; `offset` pages large folders |
| `read_file` | Read | Fetch file content as text or base64 |
| `search_sharepoint` | Read | Full-text search |
| `write_sharepoint_file` | Write | Upload or overwrite |
| `delete_sharepoint_file` | Delete | Delete a file |
| `create_sharepoint_folder` | Write | Create a folder |
| `list_sharepoint_lists` | Read | List SharePoint lists in a site |
| `list_sharepoint_list_items` | Read | List items, `offset` pages deep lists |

### OneDrive (6 tools)

| Tool | Method | Description |
|---|---|---|
| `list_onedrive` | Read | List files and folders with type, size, and mime |
| `read_onedrive_file` | Read | Fetch file content |
| `write_onedrive_file` | Write | Create or overwrite |
| `create_onedrive_folder` | Write | Create a folder |
| `move_onedrive_item` | Write | Move or rename |
| `delete_onedrive_item` | Delete | Delete a file or folder |

### Contacts (7 tools)

| Tool | Method | Description |
|---|---|---|
| `search_contacts` | Read | Search or list contacts, returning the full field set |
| `create_contact` | Write | Create a contact with notes, categories, all three postal addresses, and secondary fields |
| `update_contact` | Write | Update fields; only provided fields change |
| `delete_contact` | Delete | Delete a contact |
| `list_contact_folders` | Read | List contact folders; the default reports as `contacts-root` |
| `create_contact_folder` | Write | Create a folder, top-level or child |
| `create_contacts_batch` | Write | Bulk-create via Graph `$batch`, chunked at 20 per request, with per-entry results |

Notes, categories, and postal addresses round-trip: anything written reads back through `search_contacts`. Category tags are the practical hook for tagging an imported set and later finding or bulk-removing it.

### OneNote (5 tools)

| Tool | Method | Description |
|---|---|---|
| `list_notebooks` | Read | List notebooks |
| `create_notebook` | Write | Create a notebook |
| `list_sections` | Read | List sections |
| `create_section` | Write | Create a section |
| `create_onenote_page` | Write | Create a page from HTML |

### Teams

`list_teams`, `list_channels`, `send_channel_message`, and `send_chat_message`. These need the four Teams permissions on the registration; leave them ungranted to disable the surface.

---

## Local development

```bash
npm install
cp local.settings.json.example local.settings.json
# Edit local.settings.json with your Entra app credentials

# Table Storage emulator
npx azurite --tableHost 127.0.0.1

npm run build
npm run start

# Admin UI, in a second terminal
npm run dev:admin
```

`docker compose up` runs the whole stack in containers, Azurite included, if you would rather not install the Functions runtime locally.

Run `npm test` before opening a pull request.

---

## License

Copyright (C) 2026 Standard Gauge, LLC.

Licensed under the GNU Affero General Public License, version 3. See [LICENSE](LICENSE)
for the full text.

The network clause is deliberate: if you run a modified version of this server as a service, you owe your users the source of your modifications. Running it unmodified inside your own organisation carries no such obligation, which is the ordinary case for a company deploying it for its own staff.
