/**
 * Truncation-aware result envelopes for list/search tools.
 *
 * The connector's second core invariant is: **a result set is complete, or it
 * says it is truncated — never a short list presented as the whole answer.** A
 * caller that receives 25 rows with no further signal cannot tell "these are all
 * of them" from "these are the first 25 of thousands," and acting on the latter
 * as if it were the former is exactly how a cap horizon got mistaken for a
 * mailbox floor and how a recency-capped contact list read as
 * the whole address book.
 *
 * Every list/search tool therefore returns `count`, `limit`, and `truncated`
 * alongside its rows. `truncated: true` means the underlying collection extends
 * past what was returned; the caller can widen `maxResults`, page with `offset`
 * where the tool supports it, or narrow the query.
 *
 * Truncation is derived from a real server signal, not guessed: for a Graph
 * collection GET it is the presence of an `@odata.nextLink`; for the search API
 * it is `hitsContainers[].moreResultsAvailable`. Either is passed in as
 * `moreAvailable`. As a backstop, receiving more rows than `limit` also counts
 * as truncated, so an over-fetch pattern is reported honestly even without a
 * server continuation token.
 */

/** Standard shape returned by every list/search tool. */
export interface ListEnvelope<T> {
  /** The rows for this page, never longer than `limit`. */
  items: T[];
  /** Number of rows actually returned (`items.length`). */
  count: number;
  /** The cap applied to this call — the most rows it could have returned. */
  limit: number;
  /** Whether the underlying collection extends past what was returned. */
  truncated: boolean;
}

/**
 * Resolve a caller's requested page size into a bounded limit. A missing or
 * non-numeric request falls back to `def`; anything above `cap` is clamped to
 * `cap`; anything below 1 becomes 1. The resolved value is what the tool echoes
 * back as `limit`, so the caller always sees the cap that was actually applied.
 */
export function resolveMaxResults(requested: unknown, def: number, cap: number): number {
  const n = typeof requested === 'number' ? requested : Number(requested);
  if (!Number.isFinite(n) || n <= 0) return Math.min(def, cap);
  return Math.min(Math.floor(n), cap);
}

/**
 * Build a truncation-aware envelope. `rows` is the fully-processed row list
 * (already mapped and, where applicable, deny-list filtered). It is sliced to
 * `limit` defensively; `truncated` reflects the underlying collection, being
 * true when the server reported more (`moreAvailable`) or when more rows than
 * `limit` were collected before slicing.
 */
export function toEnvelope<T>(
  rows: T[],
  limit: number,
  moreAvailable = false,
): ListEnvelope<T> {
  const truncated = moreAvailable || rows.length > limit;
  const items = rows.length > limit ? rows.slice(0, limit) : rows;
  return { items, count: items.length, limit, truncated };
}

/**
 * Extract the "more results exist" signal from a raw Graph collection response.
 * Graph appends an `@odata.nextLink` to a collection GET whenever the server has
 * rows beyond the returned page, which is a reliable truncation signal without
 * an extra round trip.
 */
export function graphCollectionHasMore(raw: unknown): boolean {
  return Boolean(
    raw && typeof raw === 'object' && '@odata.nextLink' in (raw as Record<string, unknown>),
  );
}
