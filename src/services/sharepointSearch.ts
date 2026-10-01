/**
 * Shared helpers that apply the tenant's SharePoint access controls — the site
 * allow-list and the folder deny-list — to Microsoft Search (`/search/query`)
 * results.
 *
 * Both search surfaces (the MCP `search_sharepoint` tool and the HTTP
 * `/api/sharepoint/search` route) run through these helpers so neither can be
 * used to disclose file names, paths, or metadata from sites or folders the
 * tenant has walled off.
 *
 * Why search needs its own logic rather than reusing the per-item `siteId` gate:
 * a `/search/query` call carries no `siteId`, so the pre-dispatch allow-list
 * gate is skipped and the query would otherwise fan out tenant-wide (every site
 * plus personal OneDrive). And the deny-list matcher works on a path, but
 * Search `driveItem` hits expose only an absolute `webUrl` (which never matches
 * listing-style deny entries) and a `parentReference.path` that is drive-relative
 * and lacks the document-library segment. These helpers close both gaps:
 * post-filter by resolved site, and resolve a matchable site-relative path from
 * each hit's `webUrl` (which `denyList.filterDeniedSearchHits` then matches by
 * ancestor folder segment).
 *
 * Why the allow-list is enforced as a POST-filter rather than at the source
 *: the earlier design sent a `contentSources` constraint on the
 * `driveItem` request to scope the search at the source. That property is only
 * valid for `externalItem` (connector) entity types — Graph rejects it on a
 * `driveItem`/`listItem` request with `SearchRequest Invalid (EntityRequest
 * Invalid (Content Source is required only for ExternalItem))`, which made
 * `search_sharepoint` error on every single query. The security boundary is
 * therefore `filterHitsToAllowedSites` (fail-closed on `siteId`) plus the
 * deny-list — both applied to every hit before anything is returned — and the
 * query is allowed to fan out. To keep recall under an active allow-list (the
 * top-N most-relevant hits tenant-wide may all sit in non-allowed sites and get
 * dropped), the caller over-fetches a wider window from Graph via
 * `searchFetchSize` and slices back to the requested `maxResults` after
 * filtering.
 */

import { canonicalizePath } from './denyList.js';

export interface AllowedSite {
  id: string;
  name?: string;
}

/**
 * Normalized SharePoint search hit, carrying the fields the allow-list and
 * deny-list filters need. `siteId` drives the allow-list post-filter; `path`
 * (site-relative, resolved from `webUrl`) drives deny-list matching.
 */
export interface SearchHitItem {
  id?: string;
  name?: string;
  webUrl?: string;
  siteId?: string | null;
  path?: string;
  size?: number | null;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  summary?: string | null;
}

/**
 * `fields` to request on a `/search/query` driveItem request so the result
 * carries everything the allow-list and deny-list filters need — in particular
 * `parentReference` (for `siteId`) and `webUrl` (for the site-relative path).
 */
export const SEARCH_DRIVE_ITEM_FIELDS = [
  'id',
  'name',
  'webUrl',
  'size',
  'createdDateTime',
  'lastModifiedDateTime',
  'parentReference',
] as const;

/**
 * The widest `size` we will ever ask Microsoft Search for in a single request.
 * The Graph `/search/query` per-request ceiling is 500; we cap well under it so
 * the over-fetch stays cheap while giving the allow-list / deny-list post-filters
 * enough raw hits to still surface `maxResults` allowed results.
 */
export const SEARCH_MAX_FETCH = 200;

/** Floor for the over-fetch window when any post-filter is active. */
const SEARCH_FETCH_FLOOR = 50;

/**
 * How many hits to request from Graph for a search returning `maxResults` to the
 * caller.
 *
 * The allow-list and deny-list are enforced as post-filters (see the module
 * header — `contentSources` source scoping is invalid for `driveItem` and broke
 * search entirely), so a naive `size = maxResults` can come back short
 * whenever the most-relevant hits sit in non-allowed sites or denied folders.
 * When any filter is active we over-fetch a wider window — `maxResults` clamped
 * up to `SEARCH_FETCH_FLOOR` and down to `SEARCH_MAX_FETCH` — and the caller
 * slices back to `maxResults` after filtering. With no filter active (empty
 * allow-list, no deny entries) there is nothing to drop, so we request exactly
 * `maxResults`.
 */
export function searchFetchSize(maxResults: number, filtered: boolean): number {
  if (!filtered) return maxResults;
  return Math.min(SEARCH_MAX_FETCH, Math.max(maxResults, SEARCH_FETCH_FLOOR));
}

/**
 * Resolve the site-relative path from an absolute SharePoint / OneDrive
 * `webUrl` so it can be matched against listing-style deny-list entries such as
 * `/Shared Documents/Finance`.
 *
 * The site-collection prefix (`/sites/<name>`, `/teams/<name>`,
 * `/personal/<name>`) is stripped; root-site URLs carry no such prefix. The
 * document-library segment (`Shared Documents`, `Documents`, …) is preserved
 * because deny entries recorded from folder listings include it. The path is
 * URL-decoded so `%20`-style encoding matches a plain-text deny entry.
 *
 * Returns `null` when the URL is absent or unparseable — callers fail closed on
 * a `null` rather than fall back to the raw absolute URL (which never matches a
 * path-style deny entry).
 */
export function siteRelativePathFromWebUrl(webUrl?: string): string | null {
  if (!webUrl) return null;
  let pathname: string;
  try {
    pathname = decodeURIComponent(new URL(webUrl).pathname);
  } catch {
    return null;
  }
  const stripped = pathname.replace(/^\/(sites|teams|personal)\/[^/]+/i, '');
  const rel = stripped.length > 0 ? stripped : '/';
  return canonicalizePath(rel);
}

const GUID_RE = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi;

interface ParsedSiteId {
  siteCollectionGuid?: string;
  webGuid?: string;
}

/**
 * Parse a Graph site / site-ref composite id into its site-collection and web
 * GUIDs. The canonical shape is `hostname,siteCollectionGuid,webGuid`, so the
 * first embedded GUID is the site-collection GUID and the second (when present)
 * is the web GUID. Returned lower-cased for case-insensitive comparison.
 */
function parseSiteId(id: string): ParsedSiteId {
  const guids = (id.match(GUID_RE) ?? []).map((g) => g.toLowerCase());
  return { siteCollectionGuid: guids[0], webGuid: guids[1] };
}

/**
 * Whether a search hit's `parentReference.siteId` refers to an allow-listed
 * site. A match is either an exact case-insensitive equality on the whole
 * composite id, or an EXACT web-level identity match — the site-collection GUID
 * AND the web GUID both agree.
 *
 * Web identity, not just site-collection identity, is required: a Graph site id
 * is `hostname,siteCollectionGuid,webGuid`, and every web/subsite inside one
 * site collection shares the collection GUID while carrying a distinct web GUID.
 * Matching on the collection GUID alone (the earlier behavior) admitted hits
 * from sibling subsites the tenant never allow-listed — a cross-boundary
 * metadata leak within the same site collection (codex CR).
 *
 * Fail-closed: an id that carries no web GUID (a bare collection GUID or a
 * hostname-only id) cannot be pinned to a single web, so it never matches.
 */
function siteIsAllowed(siteId: string, allowedSites: AllowedSite[]): boolean {
  const target = siteId.toLowerCase();
  const t = parseSiteId(siteId);
  return allowedSites.some((s) => {
    if (s.id.toLowerCase() === target) return true;
    const a = parseSiteId(s.id);
    return (
      t.siteCollectionGuid !== undefined &&
      t.webGuid !== undefined &&
      t.siteCollectionGuid === a.siteCollectionGuid &&
      t.webGuid === a.webGuid
    );
  });
}

/**
 * Drop search hits that resolve to a site outside the allow-list. An empty
 * allow-list allows all. A hit whose `siteId` is missing (or cannot be matched
 * to an allowed site) is dropped when the allow-list is active (fail closed) —
 * a result we cannot attribute to an allowed site must not leak. This is a
 * defense-in-depth layer over the `contentSources` source-level constraint.
 */
export function filterHitsToAllowedSites<T extends { siteId?: string | null }>(
  items: T[],
  allowedSites: AllowedSite[],
): T[] {
  if (allowedSites.length === 0) return items;
  return items.filter(
    (it) => typeof it.siteId === 'string' && siteIsAllowed(it.siteId, allowedSites),
  );
}
