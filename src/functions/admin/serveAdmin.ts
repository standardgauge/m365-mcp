import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import * as fs from 'fs';
import * as path from 'path';
import { withSecurity, adminSpaHeaders } from '../../services/securityHeaders.js';
import { jsonForScript } from '../../services/scriptSafe.js';

// Compiled output: dist/functions/admin/serveAdmin.js
// Admin UI build: dist/admin/
const ADMIN_DIR = path.resolve(__dirname, '../../admin');

const MIME_TYPES: Record<string, string> = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.txt': 'text/plain; charset=utf-8',
  '.webmanifest': 'application/manifest+json',
};

/**
 * Injects runtime configuration into index.html so that env vars set on the
 * container at runtime take effect without requiring a rebuild.
 *
 * The config travels as a JSON data block, `<script type="application/json">`,
 * which the browser never executes and CSP's script-src does not govern. The
 * SPA reads it with JSON.parse (src/admin/runtimeConfig.ts). That is what lets
 * the admin CSP drop 'unsafe-inline' from script-src: the page has no inline
 * script left to allow (threat model 7.4). jsonForScript still escapes `<` so
 * a value cannot close the element early.
 */
export function injectRuntimeConfig(html: string): string {
  const config = {
    instanceName: process.env.MCP_INSTANCE_NAME ?? 'M365 MCP',
  };
  const injection = `<script type="application/json" id="runtime-config">${jsonForScript(config)}</script>`;
  return html.replace('</head>', `${injection}</head>`);
}

async function serveAdmin(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const rawPath = (request.params.restOfPath ?? '').replace(/\.\./g, '');

  // An unmatched /api/* path is an API call to a route that does not exist, not
  // an SPA deep link. Falling back to index.html there returns 200 text/html for
  // a dead endpoint, which makes every probe against this service unfalsifiable:
  // found the runbook checking `curl -f .../api/health`, a route that has
  // never existed, and reporting OK for months. Answer honestly instead.
  // Registered API routes are matched ahead of this catch-all and never land here.
  if (/^api(\/|$)/i.test(rawPath)) {
    return {
      status: 404,
      jsonBody: { error: 'Not found', path: `/${rawPath}` },
      headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' },
    };
  }

  const requestedFile = rawPath ? path.join(ADMIN_DIR, rawPath) : path.join(ADMIN_DIR, 'index.html');

  // Resolve the file to serve: exact match, then SPA fallback to index.html
  let resolvedPath = requestedFile;
  try {
    // eslint-disable-next-line security/detect-non-literal-fs-filename -- path is sanitised (../ stripped) and bounded to ADMIN_DIR above
    const stat = fs.statSync(resolvedPath);
    if (stat.isDirectory()) {
      resolvedPath = path.join(resolvedPath, 'index.html');
    }
  } catch {
    // File doesn't exist — fall back to SPA index.html
    resolvedPath = path.join(ADMIN_DIR, 'index.html');
  }

  // Safety check: ensure the resolved path stays inside ADMIN_DIR
  if (!resolvedPath.startsWith(ADMIN_DIR + path.sep) && resolvedPath !== path.join(ADMIN_DIR, 'index.html')) {
    return { status: 403, jsonBody: { error: 'Forbidden' } };
  }

  try {
    const ext = path.extname(resolvedPath).toLowerCase();
    // eslint-disable-next-line security/detect-object-injection -- ext is the result of path.extname(), not user input
    const contentType = MIME_TYPES[ext] ?? 'application/octet-stream';

    if (ext === '.html') {
      // eslint-disable-next-line security/detect-non-literal-fs-filename -- path is sanitised and bounded to ADMIN_DIR
      const html = fs.readFileSync(resolvedPath, 'utf-8');
      return {
        status: 200,
        body: injectRuntimeConfig(html),
        headers: { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-cache', ...adminSpaHeaders },
      };
    }

    // eslint-disable-next-line security/detect-non-literal-fs-filename -- path is sanitised and bounded to ADMIN_DIR
    const content = fs.readFileSync(resolvedPath);
    return {
      status: 200,
      body: content,
      headers: {
        'Content-Type': contentType,
        // Static assets built by Vite include content hashes — cache them aggressively
        'Cache-Control': rawPath.includes('/assets/') ? 'public, max-age=31536000, immutable' : 'no-cache',
      },
    };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    context.error('serveAdmin error:', message);
    return { status: 404, jsonBody: { error: 'Not found' } };
  }
}

app.http('serveAdmin', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: '{*restOfPath}',
  handler: withSecurity(serveAdmin),
});
