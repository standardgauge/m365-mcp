# SharePoint search access controls

`search_sharepoint` (and its HTTP twin `GET /api/sharepoint/search`) must honor
the same two tenant controls that `list_folders` and `read_file` already
enforce: the **site allow-list** and the **folder deny-list**. Before it
honored neither, so a search leaked file names, `webUrl` paths, sizes, and
timestamps from sites and folders the tenant had walled off. Content reads were
never affected — only metadata leaked — but a file name can itself be sensitive
(a deal name, a counterparty). This doc records the enforcement so neither
search surface can be used to bypass the other.

## The gap

Two controls, both skipped on search:

1. **Site allow-list never applied.** The MCP pre-dispatch gate
   (`mcpEndpoint.ts`) validates `allowedSites` only when a call carries a
   `siteId`. A `/search/query` carries none, so the gate was skipped and the
   query fanned out **tenant-wide** — every site plus personal OneDrive
   (`*-my.sharepoint.com/personal/…`). The handler also ignored the `siteId`
   arg its own schema advertised.
2. **Deny-list could not match search hits.** `filterDeniedPaths` matches on a
   path, but a Search `driveItem` hit exposes only an absolute `webUrl` (which
   never matches a listing-style deny entry like `/Shared Documents/Finance`)
   and a `parentReference.path` that is drive-relative and lacks the
   document-library segment. So deny-listed folders' contents flowed straight
   through.

## The fix

Both controls now apply on **both** paths (`search_sharepoint` MCP handler and
the HTTP function), sharing `src/services/sharepointSearch.ts`:

### 1. Scope by post-filtering on the resolved site (superseded source scoping)

> ** correction.** originally scoped the query *at the source* by
> sending a `contentSources` constraint on the `driveItem` request. That property
> is only valid for `externalItem` (connector) entity types — Graph rejects it on
> a `driveItem`/`listItem` query with `SearchRequest Invalid (EntityRequest
> Invalid (Content Source is required only for ExternalItem))`, which made
> `search_sharepoint` **error on every query, on every tenant**. Source scoping
> is therefore removed; the site scope is enforced entirely by the post-filter in
> step 2 (which was already present as defense-in-depth and is fail-closed), and
> the query is allowed to fan out.

- An explicit, already-validated `siteId` narrows the post-filter to that one
  site.
- Otherwise a non-empty allow-list restricts results to exactly the allowed
  sites.
- An empty allow-list applies no site restriction — the canonical allow-all
  default, where a fresh tenant restricts nothing until an admin configures it.

Because filtering happens *after* Graph applies its `size` cap, the most-relevant
hits tenant-wide can all sit in non-allowed sites and be dropped, leaving a short
page. To preserve recall the handler **over-fetches** when a scope is active —
`searchFetchSize` widens the requested `size` (`maxResults` clamped up to a floor
of 50 and down to a 200-hit per-request cap) and the result is sliced back to
`maxResults` after filtering. With no scope active it requests exactly
`maxResults` (nothing to drop).

The MCP `siteId` arg is honored (it was advertised but ignored before); the
pre-dispatch gate still validates it against the allow-list when present.

### 2. Filter by resolved site (the site security boundary)

Each hit's `parentReference.siteId` is checked against the active scope, dropping
anything outside it — including OneDrive hits. The match is GUID-aware: a
driveItem `parentReference.siteId` and an allow-list `/sites` id can carry the
same unique site-collection + web GUID in different composite shapes across
tenants, so a matching web identity counts as a match. A hit with no resolvable,
allow-listed site is dropped (fail closed).

### 3. Deny-list against a resolved site-relative path

The site-relative path is resolved from each hit's `webUrl`
(`/sites/<x>` · `/teams/<x>` · `/personal/<x>` prefixes stripped, document
library segment preserved, URL-decoded). Hits whose path can't be resolved are
dropped before matching (fail closed — a raw absolute URL is never passed
through). `denyList.filterDeniedSearchHits` then matches each denied entry as an
**ancestor folder segment**, so a denied folder is honored whether the entry was
recorded as a full path (`/Shared Documents/Finance`), a partial path
(`/Finance`), or a bare folder name (`Finance`) — the shape `list_folders`
records varies, but search must exclude the folder's contents regardless.
Segment comparison is exact per component, so `IR` blocks a folder literally
named `IR` but not `Investor IR Notes` — no accidental over-blocking. The
config-driven `DEFAULT_SHAREPOINT_DENY_PATHS` defaults are enforced on search
too, even with an empty admin deny table.

## Enforcement matrix

| Control | `list_folders` / `read_file` | `search_sharepoint` (before) | `search_sharepoint` (after) |
|---|---|---|---|
| Site allow-list | enforced | **not enforced** | enforced (fail-closed `siteId` post-filter; removed the broken source scoping) |
| `siteId` scoping | enforced | **ignored** | honored (post-filter narrowed to the one site) |
| Folder deny-list | enforced | **could not match** | enforced (ancestor-segment match on resolved path) |
| Fail-closed on storage down | yes | n/a (unenforced) | yes |

## Tests

- `src/__tests__/sharepointSearch.test.ts` — unit coverage of the shared helpers
  (`searchFetchSize` over-fetch sizing, `webUrl` → site-relative path, GUID-aware
  allow-list filter).
- `src/__tests__/sharepointSearchAccessControl.test.ts` — end-to-end through
  both handlers with the **real** deny-list service: non-allow-listed sites and
  personal OneDrive are dropped; items under a denied folder are stripped for
  both path-form and bare-name deny entries; `contentSources` is **never** on the
  wire while `parentReference` is requested; the MCP `siteId` arg is
  honored by post-filter.

Prior art for search-shaped deny filtering: `mailDenyListE2E.test.ts`
("search_mail strips messages in denied folders").
