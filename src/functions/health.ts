import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { withSecurity } from '../services/securityHeaders.js';

/**
 * GET /health
 *
 * Unauthenticated liveness endpoint. Returns a small JSON body so external
 * monitoring can distinguish a healthy API from a container that is only
 * serving the admin SPA's static index.html.
 *
 * `sha` is the deployed commit, injected at build time via the GIT_SHA
 * Dockerfile build-arg (see the deploy workflow). It is the empty string when the
 * image was built without that arg (e.g. a local `docker build`), which keeps
 * the endpoint useful even outside the CI deploy path.
 *
 * This route is registered as a literal path so it wins over serveAdmin's
 * `{*restOfPath}` SPA catch-all in Azure Functions route matching.
 */
async function health(
  _request: HttpRequest,
  _context: InvocationContext
): Promise<HttpResponseInit> {
  return {
    status: 200,
    jsonBody: { status: 'ok', sha: process.env.GIT_SHA ?? '' },
    headers: { 'Cache-Control': 'no-store' },
  };
}

app.http('health', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'health',
  handler: withSecurity(health),
});
