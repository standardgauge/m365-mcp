/**
 * Resolution and validation of FRONTEND_URL — the post-OAuth redirect target.
 *
 * This exists because of. Every tenant had FRONTEND_URL set to
 * `<host>/admin`, and the Azure Functions host reserves /admin/* for its own
 * key-protected admin API, so the request never reached our catch-all and
 * every login ended on an empty 404. It shipped in March and was found in
 * September: nothing validated the value, and nothing could.
 *
 * The rule is narrow on purpose. We do not try to decide whether a URL is
 * "good" — only whether it targets a path the Functions host will intercept
 * before any of our code runs, which is knowable and absolute.
 */

/**
 * Path prefixes claimed by the Azure Functions host runtime. A request to one
 * of these is answered by the host (401/404) and never dispatched to a
 * registered function, whatever routes we declare.
 */
export const HOST_RESERVED_PREFIXES = ['/admin', '/runtime'] as const;

/** Where the admin SPA is actually served: the site root. */
export const DEFAULT_FRONTEND_URL = '/';

/** True when the value carries an http(s) scheme, so `new URL` can parse it. */
function isAbsoluteUrl(value: string): boolean {
  return /^https?:\/\//i.test(value);
}

/**
 * Returns the reserved prefix a candidate redirect target collides with, or
 * null when the value is safe. Accepts absolute URLs and site-relative paths.
 *
 * A trailing-boundary check keeps this from flagging legitimate paths that
 * merely start with the same letters — `/administration` is ours, `/admin/x`
 * is the host's.
 */
export function reservedPrefixFor(value: string | undefined): string | null {
  if (!value) return null;

  let pathname: string;
  try {
    // Absolute URL — take its path. Relative value — treat as the path itself.
    pathname = isAbsoluteUrl(value)
      ? new URL(value).pathname
      : value.split('?')[0].split('#')[0];
  } catch {
    return null; // unparseable is a different problem; don't claim it's reserved
  }

  const normalised = pathname.replace(/\/+$/, '').toLowerCase() || '/';
  return (
    HOST_RESERVED_PREFIXES.find(
      (p) => normalised === p || normalised.startsWith(`${p}/`)
    ) ?? null
  );
}

/**
 * Resolves the redirect target, refusing a value the Functions host would
 * intercept. A reserved value is logged and replaced with the site root rather
 * than thrown: a misconfigured redirect must not take authentication down, and
 * the root is always correct because that is where serveAdmin serves the SPA.
 */
export function resolveFrontendUrl(
  raw: string | undefined = process.env.FRONTEND_URL,
  onProblem: (message: string) => void = (m) => console.error(m)
): string {
  const reserved = reservedPrefixFor(raw);
  if (reserved) {
    onProblem(
      `FRONTEND_URL is set to '${raw}', whose path is reserved by the Azure ` +
        `Functions host ('${reserved}'). Requests there are answered by the host ` +
        `and never reach this app, so the post-login redirect would 404. ` +
        `Falling back to '${DEFAULT_FRONTEND_URL}'. Fix the frontend-url secret ` +
        `on the Container App — see docs/operations-runbook.md → "Reserved paths".`
    );
    return DEFAULT_FRONTEND_URL;
  }
  return raw && raw.length > 0 ? raw : DEFAULT_FRONTEND_URL;
}

/**
 * Returns the origin of a configured URL, or null when it has none (a
 * site-relative value such as `/` or an unparseable string).
 */
function originOf(value: string | undefined): string | null {
  if (!value || !isAbsoluteUrl(value)) return null;
  try {
    return new URL(value).origin;
  } catch {
    return null;
  }
}

/**
 * Resolves the base to prefix onto `/api/auth/login` for the MCP
 * re-authentication message.
 *
 * The login endpoint is a function route, so it is always served at the site
 * root — never under whatever subpath the SPA happens to live at. This returns
 * the *origin* of the configured URL and drops its path entirely, which is the
 * same intent the previous `\/admin$`-stripping had but without the failure
 * mode codex flagged on this PR: stripping only an exact trailing `/admin`
 * leaves `FRONTEND_URL=https://host/admin/settings` and
 * `FRONTEND_URL=https://host/runtime` emitting login links under a path the
 * Functions host intercepts before this app ever runs.
 *
 * `OAUTH_REDIRECT_URI` is the fallback because it is required (see README) and
 * is by definition this app's own callback endpoint, so its origin is the API
 * origin. A reserved `FRONTEND_URL` is still reported, because the value is
 * wrong even though this path no longer propagates it — silence is the defect
 * exists to remove.
 */
export function resolveAuthUrlBase(
  frontendUrl: string | undefined = process.env.FRONTEND_URL,
  redirectUri: string | undefined = process.env.OAUTH_REDIRECT_URI,
  onProblem: (message: string) => void = (m) => console.error(m)
): string {
  const reserved = reservedPrefixFor(frontendUrl);
  if (reserved) {
    onProblem(
      `FRONTEND_URL is set to '${frontendUrl}', whose path is reserved by the ` +
        `Azure Functions host ('${reserved}'). The MCP re-authentication link ` +
        `ignores the path and uses the origin, so it still resolves, but the ` +
        `post-login redirect will 404. Fix the frontend-url secret on the ` +
        `Container App — see docs/operations-runbook.md → "Reserved paths".`
    );
  }
  return originOf(frontendUrl) ?? originOf(redirectUri) ?? '';
}
