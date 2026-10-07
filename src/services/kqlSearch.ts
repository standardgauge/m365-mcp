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
 *
 * Backslashes go too. The finished KQL is sent inside an outer quoted `$search`
 * string where `\` is the escape character (see `graphSearchParam`), so a value
 * ending in a backslash would otherwise escape the closing quote.
 */
export function sanitizeKqlPhrase(query: string): string {
  return query.replace(/["\\]/g, ' ');
}

/**
 * Wrap a KQL expression as the value of Graph's `$search` query option.
 *
 * Graph parses an unquoted `$search` value as plain search terms, so a property
 * restriction in it fails with `Syntax error: character ':' is not valid`. The
 * whole expression has to be one double-quoted string with its inner quotes
 * escaped as `\"`:
 *
 *   from:"alice@example.com"  ->  "from:\"alice@example.com\""
 *
 * Every caller-supplied value inside `kql` has already been through
 * `sanitizeKqlPhrase`, so the only quotes escaped here are the phrase delimiters
 * the query builder placed itself.
 *
 * The result is percent-encoded. The Graph SDK's `.search()` concatenates its
 * argument into the URL as is, so an `&`, `#` or `+` in a caller's value would
 * otherwise end the parameter early or turn into a space.
 */
export function graphSearchParam(kql: string): string {
  return encodeURIComponent(`"${kql.replace(/"/g, '\\"')}"`);
}
