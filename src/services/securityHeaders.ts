import type { HttpHandler, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { runWithClientAddress } from './clientAddress.js';

// Default response headers applied to every wrapped Azure Functions handler.
//
// Picked deliberately for a JSON API exposed publicly via Azure Container Apps:
//   - HSTS with long max-age + preload, since the service is HTTPS-only behind ACA ingress.
//   - X-Frame-Options: DENY because no admin UI is meant to be embedded.
//   - Referrer-Policy: strict-origin-when-cross-origin to avoid leaking paths
//     to upstream OAuth providers / Graph in Referer.
//   - Permissions-Policy: deny powerful features the API never needs.
//   - CSP default-src 'none': API endpoints serve JSON only; no resources to load.
//     Routes that serve HTML (serveAdmin) override this by including their own
//     Content-Security-Policy in the response headers, which wins under the
//     merge order in withSecurity (handler headers override secureHeaders).
export const secureHeaders = {
  'Strict-Transport-Security': 'max-age=63072000; includeSubDomains; preload',
  'X-Content-Type-Options': 'nosniff',
  'X-Frame-Options': 'DENY',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
  'Permissions-Policy': 'geolocation=(), microphone=(), camera=(), payment=()',
  'Content-Security-Policy': "default-src 'none'; frame-ancestors 'none'",
} as const;

// CSP for the /install HTML landing page.
// The page uses a single inline <style> block and no JavaScript — 'unsafe-inline'
// for style-src only.  Handler headers win over secureHeaders in withSecurity's
// merge order, so this effectively replaces the API default CSP for this route.
export const installLandingHeaders = {
  'Content-Security-Policy': "default-src 'none'; style-src 'unsafe-inline'; frame-ancestors 'none'",
} as const;

// CSP for the admin React SPA. Overrides the API-focused default above when
// returned in the handler's own headers object (handler headers win over secureHeaders).
//
// script-src is 'self' only: the page carries no inline script. serveAdmin injects
// its runtime config as a non-executable JSON data block, which script-src does not
// govern and the SPA reads with JSON.parse (threat model 7.4). style-src keeps
// 'unsafe-inline' for React style attributes.
//
// connect-src covers same-origin API calls plus MSAL's token endpoints and Microsoft
// Graph. form-action 'self' restricts form submissions; MSAL redirect flows target
// login.microsoftonline.com so we include it there too.
export const adminSpaHeaders = {
  'Content-Security-Policy': [
    "default-src 'none'",
    "script-src 'self'",
    "style-src 'self' 'unsafe-inline'",
    "img-src 'self' data:",
    "font-src 'self'",
    "connect-src 'self' https://login.microsoftonline.com https://graph.microsoft.com https://login.windows.net",
    "frame-ancestors 'none'",
    "form-action 'self' https://login.microsoftonline.com",
    "base-uri 'self'",
  ].join('; '),
} as const;

// Wrap an HttpHandler so its response carries the secure header set.
// Existing headers on the handler's HttpResponseInit always win (so a route
// can opt into Cache-Control, Content-Type, etc. without being overridden).
export function withSecurity(handler: HttpHandler): HttpHandler {
  return async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    // Audit rows written while handling this request record its client address.
    const result = (await runWithClientAddress(request, () => handler(request, context))) ?? {};
    return {
      ...result,
      headers: {
        ...secureHeaders,
        ...(result.headers ?? {}),
      },
    };
  };
}
