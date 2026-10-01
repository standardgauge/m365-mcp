/**
 * Offset-based pagination over a Microsoft Graph collection.
 *
 * SharePoint drive-children and list-items collections do not support `$skip`;
 * Graph pages them with a server-issued `@odata.nextLink`. To let a caller reach
 * an item deep in a multi-thousand-item library — which `top`-only listing could
 * never surface — `collectPage` walks those `nextLink` pages internally until it
 * has buffered the requested `[offset, offset + limit)` window, then returns that
 * slice plus a `hasMore` flag and the `nextOffset` to resume from.
 *
 * Only Graph's own `nextLink` URLs are followed (never a caller-supplied URL), so
 * there is no SSRF surface: the caller controls a numeric `offset`, not a URL.
 * Pagination is bounded by `MAX_PAGES` to defend against a malformed or cyclic
 * `nextLink` chain, mirroring the `getTenantUsers` fetch loop.
 */

/** A single Graph collection page. */
export interface GraphPage<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

/** The requested window of a collection plus continuation metadata. */
export interface PagedResult<T> {
  items: T[];
  /** Whether the raw collection extends past `offset + limit`. */
  hasMore: boolean;
  /** Raw-collection index to resume from (pass as the next `offset`). */
  nextOffset: number;
}

// 50 pages is far above any plausible library depth at the page sizes used here
// and exists only to bound a malformed/cyclic nextLink chain.
const MAX_PAGES = 50;

/**
 * Walk `@odata.nextLink` pages until the `[offset, offset + limit)` window is
 * buffered (or the collection is exhausted / the page cap is hit), then return
 * that slice.
 *
 * `fetchFirst` issues the initial request; `fetchNext` follows a Graph-issued
 * `nextLink`. Both return the raw Graph page so this helper stays independent of
 * the Graph client type and is trivially unit-testable.
 */
export async function collectPage<T>(
  fetchFirst: () => Promise<GraphPage<T>>,
  fetchNext: (nextLink: string) => Promise<GraphPage<T>>,
  offset: number,
  limit: number,
): Promise<PagedResult<T>> {
  const start = Math.max(0, offset);
  const target = start + Math.max(0, limit);
  const all: T[] = [];
  const visited = new Set<string>();

  let page = await fetchFirst();
  if (Array.isArray(page.value)) all.push(...page.value);
  let nextLink = page['@odata.nextLink'];

  let pages = 1;
  while (all.length < target && nextLink && !visited.has(nextLink) && pages < MAX_PAGES) {
    visited.add(nextLink);
    page = await fetchNext(nextLink);
    if (Array.isArray(page.value)) all.push(...page.value);
    nextLink = page['@odata.nextLink'];
    pages++;
  }

  const items = all.slice(start, target);
  // More remain if we buffered past the window, or the server still has pages.
  const hasMore = all.length > target || Boolean(nextLink);
  return { items, hasMore, nextOffset: start + items.length };
}
