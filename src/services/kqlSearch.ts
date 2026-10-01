/**
 * Sanitize a user-supplied query for embedding inside a KQL $search phrase.
 *
 * Graph mail `$search` wraps the query in a double-quoted KQL phrase:
 *   .search(`"${q}"`)
 * A double-quote inside `q` closes the phrase early and lets the remainder be
 * interpreted as additional KQL operators / property restrictions, widening the
 * search scope beyond the intended phrase. The blast radius stays bounded by the
 * caller's own delegated token, but it bypasses the intended query shape
 * (F9).
 *
 * We strip the double-quotes rather than escape them: KQL has no portable
 * in-phrase escape for `"` across Exchange Online tiers, and a phrase search
 * over the remaining terms is the expected behaviour anyway. This mirrors the
 * defensive OData escaping the calendar/contacts filters already do.
 */
export function sanitizeKqlPhrase(query: string): string {
  return query.replace(/"/g, ' ');
}
