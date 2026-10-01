# SharePoint MCP Server — Spec & Design

**Parent project:** m365-mcp (m365-mcp) \
**Status:** Planning \
**Last updated:** 2026-03-19

* * *

## Problem

The built-in Claude M365 connector uses delegated permissions — Claude inherits the user's full SharePoint access. There is no way to restrict Claude from accessing specific folders while the user retains access to those same folders. This is a deal-breaker for Example, where sensitive folders (Finance, HR, Legal, IR, etc.) must be off-limits to AI while remaining accessible to the humans who need them.

## Solution

A custom MCP server hosted on Azure Functions that proxies all SharePoint access through a server-side deny list. The server authenticates each user individually via OAuth 2.0 / Entra ID, so Microsoft's own permission system still applies as a base layer. On top of that, the server enforces an admin-configurable deny list that blocks Claude from accessing designated folders — regardless of whether the user has access.

The built-in M365 connector is disabled org-wide; all SharePoint access flows through this custom server.

## Architecture

  * **Runtime:** Azure Functions (serverless, Node.js/TypeScript)

  * **Auth:** OAuth 2.0 Authorization Code flow, per-user tokens via Microsoft Entra ID

  * **API:** Microsoft Graph API for all SharePoint/OneDrive operations

  * **Config storage:** Azure Blob Storage (deny list JSON + audit log)

  * **Token storage:** Azure Key Vault (refresh tokens, keyed by user)

  * **Monitoring:** Azure Monitor / Application Insights

  * **Admin UI:** Azure Static Web App (Phase 2) — see Admin Experience section

  * **Starting point:** Fork of sekops-ch/sharepoint-mcp-server (TypeScript, MIT)




## Deny List

Prefix-based, case-insensitive path matching. Enforced at two points: before forwarding requests to Graph API, and after receiving search results (stripping denied paths from results before returning to Claude).

### Denied Folders (initial config)

  * Shared Documents/Billing

  * Shared Documents/Business Development

  * Shared Documents/Compliance

  * Shared Documents/HR

  * Shared Documents/Finance

  * Shared Documents/IR

  * Shared Documents/Legal

  * Shared Documents/Management Company

  * Shared Documents/Social Impact

  * Shared Documents/Transfer

  * Shared Documents/Weekly Meeting

  * Shared Documents/IT




Personal OneDrive access is denied by default (denyPersonalOneDrives: true).

### Approved Folders

  * Shared Documents/Investment Team

  * Shared Documents/_SoMa Portfolio

  * Shared Documents/Admin

  * Shared Documents/Compliance Policies and Procedures (distinct from Compliance)

  * Shared Documents/Office

  * Shared Documents/Operations

  * Shared Documents/Trading




## Admin Experience

The deny list needs to be manageable by a non-technical admin who can see what's available and toggle access on or off per folder. This is delivered in two phases.

### Phase 1 — MCP Folder Browser Tool (ships with v1)

A new admin-only MCP tool, `browse_site_folders`, that enumerates the top-level folder structure of a SharePoint site via Graph API and annotates each folder with its current deny/allow status from the config. The admin interacts with Claude conversationally to toggle folders.

**Tool: browse_site_folders**

  * Input: site URL or site ID (optional — defaults to primary site)

  * Input: depth (optional — defaults to 1, max 3)

  * Auth: admin-only (same Entra ID allowlist as other admin tools)

  * Behavior:

    1. Call Graph API to enumerate folders at the specified depth

    2. Load current deny list from Blob Storage

    3. Annotate each folder: denied (explicit match), denied (inherited from parent), or allowed

    4. Return the annotated tree to Claude for display




**Tool: toggle_folder_access**

A convenience wrapper that combines browse + toggle in one step. Accepts a folder path and a desired state (deny or allow). Validates the path exists in SharePoint before updating config. Returns the updated tree view for confirmation.

### Phase 2 — Admin Web UI (future)

An Azure Static Web App (React) that provides a visual folder tree with checkboxes for deny list management. This is the proper admin console for initial setup, periodic reviews, and bulk changes.

  * Authenticates admin via Entra ID (same app registration, admin-only scope)

  * Fetches folder tree from Graph API

  * Loads and saves deny list config from/to Azure Blob Storage

  * Tree view with checkboxes — check to deny, uncheck to allow

  * Shows inherited deny status (parent denied = children auto-denied)

  * Save button writes updated config to Blob Storage

  * MCP server picks up changes on next cache refresh (60s TTL)




This is a separate deployment (Azure Static Web Apps) with its own build pipeline. The backend is the same Blob Storage config — the web UI and MCP admin tools are two interfaces to the same data.

## MCP Tools

### User-Facing Tools

Tool| Description| Deny List Enforcement \
---|---|--- \
sharepoint_search| Keyword search across SharePoint via Graph Search API| Post-filter: strip denied paths from results \
sharepoint_read_file| Read file content by drive item ID or URL| Pre-check: validate path before calling Graph \
sharepoint_list_folder| List folder contents| Pre-check: validate path before calling Graph \
sharepoint_find_site| Discover accessible SharePoint sites| None needed (site-level discovery is safe) \
onedrive_search| Search personal OneDrive| Gated by denyPersonalOneDrives flag \

### Admin Tools (owner-only)

Tool| Description \
---|--- \
browse_site_folders| Enumerate folder tree with deny/allow annotations (Phase 1) \
toggle_folder_access| Set a folder to denied or allowed with path validation (Phase 1) \
list_deny_rules| Return current deny list \
add_deny_rule| Add path to deny list \
remove_deny_rule| Remove path from deny list \
set_personal_onedrive_access| Toggle personal OneDrive access flag \

Admin tools require Entra ID identity validation against an admin allowlist, enforced as middleware before any handler runs.

## Auth Design

  * OAuth 2.0 Authorization Code flow with Entra ID

  * App registration with delegated permissions: Sites.Read.All, Files.Read.All, User.Read

  * Per-user access and refresh tokens

  * Refresh tokens stored in Azure Key Vault, keyed by user identity

  * All Graph API calls made with the individual user's token




## Key Decisions

Decision| Choice| Rationale \
---|---|--- \
Deny list vs. allowlist mode| Start with deny list, consider allowlist-only mode later| Deny list matches Example's current folder review output; allowlist is more secure but needs more planning \
Config storage| Azure Blob Storage| Simple, low-cost, accessible only via managed identity \
Base repo| sekops-ch/sharepoint-mcp-server| Already implements Graph integration and MCP tool structure \
Token storage| Azure Key Vault| Industry standard for secret storage in Azure \
Admin experience| MCP folder browser (v1) + web UI (v2)| MCP tool ships with the server and covers day-to-day changes. Web UI provides visual tree for setup and bulk reviews. Both write to the same Blob Storage config. \

## Security Requirements

  * HTTPS only (no HTTP)

  * Azure Function endpoint protected with Entra ID authentication

  * Deny list config readable only by Function App managed identity

  * Full request audit logging via Application Insights (user, path, permitted/denied)

  * denyPersonalOneDrives: true is the hard default — opt-in only

  * Admin tools gated by Entra ID group membership — not just any authenticated user




## Mailbox Access (Exchange/Outlook)

Exchange/Outlook mailbox access is a first-class feature of the m365-mcp MCP, handled with the same two-tier access control model as SharePoint.

### Why This Is Included Here

The only available off-the-shelf Claude Outlook connector bundles SharePoint and mail together in a single connector with no granular access controls. This makes it impossible to grant Claude mail access without simultaneously granting uncontrolled SharePoint access — a non-starter at Example where SharePoint access is tightly governed. The m365-mcp MCP handles both in one place and applies the same deny list model consistently to both surfaces.

### Bulk Email Triage Workflow

The move and delete tools exist to support a class of workflows where Claude reads, aggregates, and then acts on a large set of related emails in a single conversational exchange. Two representative patterns:

  * **Analyst blast emails per ticker.** During earnings season a portfolio manager may receive 10–20 analyst notes about the same company within a short window. Claude searches for all emails on a given ticker, aggregates them into a single summary, surfaces outlier perspectives (e.g. a differentiated bull/bear call or a management contact note that others lack), and then bulk-deletes or archives the full set on confirmation — leaving the user with one synthesised view rather than a pile of largely redundant messages.
  * **Real-time alert digests.** Services like Street Account fire intraday alerts continuously. Claude can summarise all messages from a given service over a time window and then delete or archive them in bulk. Claude resolves the sender pattern semantically (identifying all relevant sender addresses automatically) rather than requiring the user to specify exact addresses.



Both patterns require **read + delete** at minimum, and **move** to support an archive-instead-of-delete preference. The email triage agent or skill that implements these workflows is a separate deliverable; the tools below are the MCP primitives it will depend on.

### MCP Tools

**Tool**| **Description**| **Deny List Enforcement** \
---|---|--- \
list_mailboxes| List mailboxes accessible to the authenticated user| None needed (mailbox discovery is safe) \
list_folders_mail| List mail folders within a mailbox| Pre-check: validate folder path against deny list before returning \
search_mail| Keyword search across mail in a mailbox| Post-filter: strip messages from denied folders before returning results \
read_message| Read full content of a specific message by ID| Pre-check: validate that the message's parent folder is not denied \
delete_messages| Permanently delete one or more messages by ID. Accepts a list of message IDs to support bulk deletion (e.g. delete all analyst notes on a given ticker after summarising them).| Pre-check: validate each message's parent folder against deny list before deletion; denied-folder messages are skipped and reported \
move_messages| Move one or more messages to a target mail folder. Accepts a list of message IDs and a destination folder ID. Supports bulk archive workflows as an alternative to deletion.| Pre-check: validate source folder is not denied; destination folder must not be a denied folder \

### Two-Tier Deny List (Mail)

The same two-tier deny list model that governs SharePoint access applies to mail folders:

  * **Admin global deny list** — org-wide mail folders blocked for all users (e.g. HR mailboxes, executive mailboxes, legal hold folders).
  * **Per-user personal deny list** — individual users can designate their own mail folders as off-limits to Claude (e.g. personal folders, sensitive client threads).



Deny list entries are matched against the full folder path, prefix-based and case-insensitive, consistent with SharePoint enforcement.

### Admin UI

The Phase 2 React admin UI includes a mail folder browser alongside the SharePoint folder browser. Admins can navigate mailbox folder trees and toggle deny/allow status per folder. Both SharePoint and mail configurations are managed in the same UI and written to the same Blob Storage config. The MCP server picks up changes on the next cache refresh (60s TTL).

### Graph API Permissions

Additional delegated permissions required for mail access:

  * `Mail.ReadWrite` — read, move, and delete messages in the signed-in user's mailbox (replaces `Mail.Read`; the broader scope is required to support delete and move operations)
  * `Mail.ReadBasic.All` — read basic mail properties across mailboxes (delegated)



These are added to the existing app registration alongside `Sites.Read.All` and `Files.Read.All`.

### Connector Strategy

**Do NOT use the off-the-shelf Claude Outlook/mail connector.** All mail access goes through the m365-mcp MCP. This ensures consistent access controls, a single audit log, and a single admin UI across both SharePoint and mail. Installing the off-the-shelf connector alongside this MCP would create a parallel uncontrolled access path that bypasses the deny list entirely.

## Why a Custom MCP (Not Off-the-Shelf Connectors)

Four reasons this MCP is purpose-built rather than assembled from available connectors:

  1. **The only Outlook connector bundles SharePoint access.** The off-the-shelf Claude connector for Outlook/Exchange also provides SharePoint access — they ship as one unit. There is no way to grant Claude mail access without simultaneously granting it uncontrolled SharePoint access. Since SharePoint access at Example must go through the deny list, this connector cannot be used at all.

  2. **No off-the-shelf connector supports our two-tier deny list model.** The admin global deny list + per-user personal deny list combination is a custom access control requirement. No available connector implements this; they either inherit the user's full permissions or operate on a broad service-account basis with no per-folder exclusions.

  3. **A single MCP gives us a single access control surface.** Centralizing SharePoint and mail access in one server means one deny list config, one audit log (via Application Insights), and one admin UI. If we used separate connectors for SharePoint and mail, we would need to replicate access control logic across two systems and keep them in sync — a maintenance and compliance risk.

  4. **Multi-user OAuth from day one.** The m365-mcp MCP uses OAuth 2.0 Authorization Code flow with per-user tokens stored in Azure Key Vault. Each user's Graph API calls are made with their own identity, so Microsoft's native permissions remain the base layer. Some off-the-shelf connectors use shared service accounts, which dilutes audit trails and makes per-user deny lists impossible to enforce reliably.




## Open Questions

  * Should we implement allowlist-only mode from the start, or defer?

  * What is the TTL for deny list config caching? (Spec suggests 60 seconds)

  * How should admin allowlist be maintained — Entra ID group, or static config?

  * Does Example need per-user deny list overrides, or is the global list sufficient?

  * Admin web UI: build in-house with Azure Static Web Apps, or evaluate existing SharePoint admin tooling?




## Local Development Setup & Known Issues

### Entra App Registration

  * **App name:** m365-mcp Sharepoint MCP

  * **Permissions:** Files.ReadWrite.All, Mail.ReadBasic, Mail.ReadWrite, Mail.Send, Sites.ReadWrite.All, User.Read (all delegated)

  * **Admin consent:** granted at both app registration and enterprise application level

  * **Redirect URI:** http://localhost:7071/api/auth/callback




### Local Dev Setup

  * Clone repo: github.com/example-org/m365-mcp-sharepoint-mcp

  * Copy `local.settings.json.example` → `local.settings.json`, fill in AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, AZURE_TENANT_ID

  * Install Azure Functions Core Tools v4: `brew install azure-functions-core-tools@4`

  * `npm install && npm run build`

  * `func start` (port 7071), `npm run dev:admin` (port 5173)




### Known Auth Issue: CA Policy Blocking Interactive Login

  * Browser auth code flow blocked by Example tenant CA policy requiring device compliance

  * Error: AADSTS900144 — request not even reaching Entra sign-in logs

  * Likely cause: Microsoft Enterprise SSO plugin or corporate proxy intercepting OAuth redirects

  * **Workaround:** Device code flow at `GET /api/auth/device`

  * Device code flow also blocked for nate@example-org.com by same CA policy

  * **Next steps:** (a) test with DataFlow_Service@example-org.com account, OR (b) add m365-mcp Sharepoint MCP app to CA policy exclusion list in Entra Security → Conditional Access




### Code Fixes Applied (committed to main)

  * Mail.ReadBasic.All → Mail.ReadBasic scope fix

  * GRAPH_SCOPES aligned with Entra registration (ReadWrite variants)

  * Device code flow endpoint added (`/api/auth/device`)

  * tsconfig + package.json Azure Functions discovery fixes

  * Frontend index.html added for Vite dev server




* * *

Full build specification is maintained as a separate document (soma_mcp_spec.md). This page tracks architecture decisions and current state.
