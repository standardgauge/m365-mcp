/**
 * Validation helpers for opaque Microsoft Graph resource IDs.
 *
 * The Microsoft Graph SDK v3 does not percent-encode path segments on the way
 * out: characters like `/`, `?`, and `#` are forwarded verbatim to the API.
 * Interpolating caller-supplied IDs directly into a path therefore allows an
 * attacker to:
 *   - traverse to an unintended resource  (/sites/allowed/../../../users/victim)
 *   - inject OData query options          (siteId?$top=1000&$select=secret)
 *   - escape the allowedSites / deny-list scope
 *
 * Legitimate Graph opaque IDs (base64url-encoded strings, GUIDs, UPNs,
 * SharePoint composite IDs) do not contain any of the blocked characters.
 */

/** Thrown by assertOpaqueId / assertOpaqueIds when a value fails validation. */
export class ValidationError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ValidationError';
  }
}

/** Characters that can alter a URL path or inject OData query parameters. */
const UNSAFE_ID_RE = /[/?#%\\\r\n\x00]/;

/**
 * Assert that `value` is safe to interpolate as a single path segment in a
 * Microsoft Graph API URL.  Throws an `Error` with a descriptive message when
 * the value contains characters that could manipulate the resulting URL.
 *
 * Pass `paramName` to make the error actionable (it appears in the message).
 *
 * Special token: the literal string `'me'` is accepted without validation —
 * it is the Graph API self-reference token and routes to `/me`, not to
 * `/users/me`.  Callers that handle `'me'` separately before building a path
 * (e.g. `mailboxId`) may still call this function on the raw value; it will
 * return normally because `'me'` contains none of the blocked characters.
 */
export function assertOpaqueId(value: string, paramName: string): void {
  if (!value || typeof value !== 'string') {
    throw new ValidationError(`${paramName} must be a non-empty string`);
  }
  if (UNSAFE_ID_RE.test(value)) {
    throw new ValidationError(`${paramName} contains characters that are not permitted in a resource identifier`);
  }
}

/**
 * Percent-encode an opaque Graph resource ID for safe interpolation as a single
 * path segment.
 *
 * The Microsoft Graph SDK v3 forwards path segments verbatim — it does NOT
 * percent-encode them. Graph event and message IDs are base64 strings that
 * routinely contain `+` and `=` (and, when standard base64, `/`). A raw `+` in
 * the URL path is decoded to a space by Graph's front end, corrupting the ID so
 * that GET / PATCH / DELETE fail with "The Id is invalid" — silently, since the
 * ID looks fine in every echoed response. This is exactly why get_event /
 * update_event / delete_event / respond_to_event could not read back or amend
 * an event that create_event had just returned: the write path
 * worked because POST /me/events carries the ID in the body, and every read/
 * mutate path carries it in the URL.
 *
 * `encodeURIComponent` turns `+` → `%2B`, `/` → `%2F`, `=` → `%3D`, etc., which
 * Graph decodes back to the original ID. It also neutralizes path-traversal and
 * OData-injection payloads (`/`, `?`, `#` all become percent-escapes that stay
 * inside the segment), so encoding is both the correctness fix and a second
 * layer of the protection that `assertOpaqueId` provides up front.
 *
 * Callers should still validate with `assertOpaqueId` / `assertOpaqueIds`
 * first; this function only guards against a non-string slipping through.
 */
export function encodeGraphId(value: string, paramName: string): string {
  if (!value || typeof value !== 'string') {
    throw new ValidationError(`${paramName} must be a non-empty string`);
  }
  return encodeURIComponent(value);
}

/**
 * Validate every opaque-ID argument in a tool-args map.
 *
 * `idParams` is the set of argument names that are expected to be opaque
 * Graph resource IDs.  Any key in `args` that appears in `idParams` is
 * validated with `assertOpaqueId`.  Non-string values (numbers, booleans) and
 * nullish values are skipped — type coercion is the caller's responsibility.
 *
 * Throws on the first invalid value found.
 */
export function assertOpaqueIds(
  args: Record<string, unknown>,
  idParams: ReadonlySet<string>,
): void {
  for (const [param, rawValue] of Object.entries(args)) {
    if (!idParams.has(param)) continue;
    if (rawValue == null || rawValue === '') continue;
    if (typeof rawValue !== 'string') continue;
    assertOpaqueId(rawValue, param);
  }
}
