/**
 * SharePoint site filter — applied to the raw output of Graph's
 * /sites?search=* endpoint before returning to MCP / admin UI callers.
 *
 * Two passes:
 *
 * 1. **Drop system sites.** Microsoft Graph returns several internal site
 *    collections that aren't user content (compliance hubs, app catalogs,
 *    content type hubs, personal OneDrive sites). These leak into the
 *    site picker as confusing duplicates and have no folders/files for
 *    users to navigate. Filter by URL path against a known list.
 *
 * 2. **Disambiguate remaining displayName collisions.** Even after step 1,
 *    two real user sites can legitimately share a display name (Graph
 *    doesn't enforce uniqueness — only the path segment is unique).
 *    For any duplicate name, append the trailing URL path segment in
 *    parentheses so the picker shows e.g.
 *      "Marketing (sites/marketing)"
 *      "Marketing (sites/marketing-archive)"
 *    Sites with unique display names are returned unchanged.
 *
 * The filter is centralized here so the MCP `list_sites` tool, the
 * admin UI HTTP endpoint (`/api/sharepoint/sites`), and any future
 * caller all behave consistently. Discovered 2026-04-10 via
 * after the example tenant returned two sites both named "SFPs"
 * (one was the contentTypeHub, the other was /sites/SFPs).
 */

// Known SharePoint system / hub site path segments. Match is suffix-based:
// any URL whose path component ends with one of these is dropped.
//
// Reference: https://learn.microsoft.com/en-us/sharepoint/sharepoint-online-tenant-types
const SYSTEM_SITE_PATH_SEGMENTS: readonly string[] = [
  '/contentTypeHub',
  '/appcatalog',
  '/CompliancePolicyCenter',
  '/CompliancePolicyCenterRedirect',
  '/ComplianceCenterRedirectionSite',
  '/RecordsCenter',
  '/PWA',
  '/SearchCenter',
  '/SitePages',
  '/sitemaster',
];

// Personal OneDrive sites have URLs like /personal/<user>_<tenant>_<tld>
// — match the prefix instead of an exact suffix.
const PERSONAL_SITE_PREFIX = '/personal/';

interface SitePartial {
  webUrl: string;
  displayName?: string;
}

/**
 * Returns true if the site's webUrl indicates a system / internal site
 * that should be hidden from end-user pickers.
 */
export function isSystemSite(site: SitePartial): boolean {
  let url: URL;
  try {
    url = new URL(site.webUrl);
  } catch {
    return false; // Malformed URL — leave it visible so it's at least debuggable
  }
  const pathname = url.pathname;

  if (pathname.startsWith(PERSONAL_SITE_PREFIX)) {
    return true;
  }
  for (const seg of SYSTEM_SITE_PATH_SEGMENTS) {
    // Case-insensitive suffix match: contentTypeHub vs ContentTypeHub
    if (pathname.toLowerCase().endsWith(seg.toLowerCase())) {
      return true;
    }
  }
  return false;
}

/**
 * Extract the trailing URL path segment for use as a disambiguator.
 *   https://x.sharepoint.com/sites/marketing       -> "sites/marketing"
 *   https://x.sharepoint.com                       -> "(root)"
 *   https://x.sharepoint.com/teams/eng/sub         -> "teams/eng/sub"
 *
 * Used when two sites share the same displayName so the picker can
 * distinguish them.
 */
export function siteDisambiguator(webUrl: string): string {
  try {
    const url = new URL(webUrl);
    const path = url.pathname.replace(/^\/+|\/+$/g, '');
    return path === '' ? '(root)' : path;
  } catch {
    return webUrl;
  }
}

/**
 * Apply both passes (system filter + collision disambiguation) and return
 * a new array. Input is not mutated.
 */
export function filterAndDisambiguateSites<T extends SitePartial & { displayName?: string }>(
  sites: readonly T[]
): T[] {
  const visible = sites.filter((s) => !isSystemSite(s));

  // Count occurrences of each displayName so we know which ones need
  // disambiguation. Empty / missing names are not collapsed.
  const nameCounts = new Map<string, number>();
  for (const s of visible) {
    const name = s.displayName ?? '';
    nameCounts.set(name, (nameCounts.get(name) ?? 0) + 1);
  }

  return visible.map((s) => {
    const name = s.displayName ?? '';
    if ((nameCounts.get(name) ?? 0) > 1) {
      // Collision — append the URL disambiguator. Return a shallow copy so
      // the caller's input array isn't mutated.
      return {
        ...s,
        displayName: `${name} (${siteDisambiguator(s.webUrl)})`,
      };
    }
    return s;
  });
}
