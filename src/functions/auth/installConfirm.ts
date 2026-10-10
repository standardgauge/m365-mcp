import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { timingSafeEqual } from 'crypto';
import {
  attachSessionToInstallNonce,
  deleteInstallHandoff,
  getInstallHandoff,
  recordInstallHandoffFailure,
} from '../../services/tableStorage.js';
import type { StoredInstallHandoff } from '../../services/tableStorage.js';
import { hashSessionToken } from '../../services/credentialCrypto.js';
import {
  confirmationCodeMatches,
  INSTALL_CONFIRM_PATH,
  INSTALL_HANDOFF_COOKIE,
} from '../../services/installConfirm.js';
import { withRateLimit } from '../../services/rateLimit.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { auditActor, logAccess } from '../../services/auditLog.js';

const NONCE_TTL_MS = 5 * 60 * 1000;
const HANDOFF_ID_FORMAT = /^[a-f0-9]{64}$/;

// One inline <style> block, no script. The form posts back to this route.
const PAGE_HEADERS = {
  'Content-Type': 'text/html; charset=utf-8',
  'Cache-Control': 'no-store',
  'Content-Security-Policy':
    "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; frame-ancestors 'none'",
} as const;

/**
 * GET|POST /api/auth/install-confirm
 *
 * Second half of an installer's sign-in. The OAuth callback records a pending
 * handoff and redirects here with its id in an HttpOnly cookie. GET shows who
 * is signed in, the device label the installer gave, and a box for the code the
 * installer printed. POST checks the code and, if it matches, attaches this
 * browser's session to the installer's nonce so install-poll can return it.
 *
 * Someone who was only sent a sign-in link has no installer and so no code: the
 * page tells them to close it, and "This wasn't me" discards the handoff. Five
 * wrong codes discard it too.
 *
 * The session handed over is the one in this request's mcp_session cookie, and
 * only if its keyed hash matches the one the callback recorded, so the page can
 * hand over nothing but the session this sign-in created.
 */
async function installConfirm(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const cookieHeader = request.headers.get('cookie') ?? '';
  const handoffId = readCookie(cookieHeader, INSTALL_HANDOFF_COOKIE);
  if (!handoffId || !HANDOFF_ID_FORMAT.test(handoffId)) {
    return page(410, gonePage());
  }

  const stored = await getInstallHandoff(handoffId);
  if (!stored) {
    return page(410, gonePage(), clearHandoffCookie());
  }

  if (request.method !== 'POST') {
    return page(200, formPage(stored));
  }

  const form = new URLSearchParams(await request.text());

  if (form.get('action') === 'cancel') {
    await deleteInstallHandoff(handoffId);
    context.log(`install handoff declined by the signed-in user (userId=${stored.record.userId})`);
    auditHandoff(stored, 'confirm', 'declined by the signed-in user');
    return page(
      200,
      messagePage(
        'Request discarded',
        'Nothing was connected. If you did not start this sign-in, someone may have sent you the link on purpose. Tell your administrator.'
      ),
      clearHandoffCookie()
    );
  }

  const sessionToken = readCookie(cookieHeader, 'mcp_session');
  if (!sessionToken || !sameHash(hashSessionToken(sessionToken), stored.record.sessionTokenHash)) {
    context.warn(`install handoff confirm without the session it was issued for (userId=${stored.record.userId})`);
    return page(403, messagePage('Sign in again', 'This browser no longer holds the session for this request. Run the installer again.'));
  }

  if (!confirmationCodeMatches(stored.record.challenge, form.get('code') ?? '')) {
    const remaining = await recordInstallHandoffFailure(handoffId, stored);
    if (remaining === 0) {
      context.warn(`install handoff discarded after repeated wrong codes (userId=${stored.record.userId})`);
      auditHandoff(stored, 'confirm', 'discarded after repeated wrong confirmation codes');
      return page(
        400,
        messagePage('Too many attempts', 'That code did not match and the request has been discarded. Run the installer again.'),
        clearHandoffCookie()
      );
    }
    const note = remaining === null
      ? 'That code did not match. Try again.'
      : `That code did not match. ${remaining} ${remaining === 1 ? 'attempt' : 'attempts'} left.`;
    return page(400, formPage(stored, note));
  }

  // One-time: only the request that deletes the version it read goes on to attach.
  if (!(await deleteInstallHandoff(handoffId, stored.etag))) {
    return page(410, gonePage(), clearHandoffCookie());
  }
  const attached = await attachSessionToInstallNonce(stored.record.challenge, {
    sessionToken,
    userId: stored.record.userId,
    email: stored.record.email,
    displayName: stored.record.displayName,
    deviceLabel: stored.record.deviceLabel,
    tenantId: stored.record.tenantId,
    expiresAt: Date.now() + NONCE_TTL_MS,
  });
  auditHandoff(stored, 'attach', attached ? undefined : 'session could not be attached to the install nonce');
  if (!attached) {
    context.error(`install handoff confirmed but could not be stored (userId=${stored.record.userId})`);
    return page(
      500,
      messagePage('Something went wrong', 'The installer could not be connected. Run it again.'),
      clearHandoffCookie()
    );
  }

  return page(
    200,
    messagePage('Connected', 'Go back to the installer; it will finish on its own. You can close this tab.'),
    clearHandoffCookie()
  );
}

/**
 * Audit row for a confirmation step. No reason means the session was attached;
 * a reason records why the handoff ended without one.
 */
function auditHandoff(stored: StoredInstallHandoff, resource: 'attach' | 'confirm', reason?: string): void {
  logAccess({
    ...auditActor(stored.record),
    operation: 'auth.install_handoff',
    resource,
    result: reason ? 'denied' : 'allowed',
    ...(reason ? { reason } : {}),
    source: 'http',
  });
}

function readCookie(header: string, name: string): string | null {
  for (const part of header.split(';')) {
    const eq = part.indexOf('=');
    if (eq > 0 && part.slice(0, eq).trim() === name) return part.slice(eq + 1).trim() || null;
  }
  return null;
}

function sameHash(a: string, b: string): boolean {
  return a.length === b.length && timingSafeEqual(Buffer.from(a), Buffer.from(b));
}

function clearHandoffCookie() {
  return [
    {
      name: INSTALL_HANDOFF_COOKIE,
      value: '',
      httpOnly: true,
      secure: true,
      sameSite: 'Lax' as const,
      path: INSTALL_CONFIRM_PATH,
      maxAge: 0,
    },
  ];
}

function page(
  status: number,
  body: string,
  cookies?: ReturnType<typeof clearHandoffCookie>
): HttpResponseInit {
  return { status, headers: { ...PAGE_HEADERS }, body, ...(cookies ? { cookies } : {}) };
}

const HTML_ESCAPES = new Map([
  ['&', '&amp;'],
  ['<', '&lt;'],
  ['>', '&gt;'],
  ['"', '&quot;'],
  ["'", '&#39;'],
]);

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (c) => HTML_ESCAPES.get(c) ?? c);
}

function instanceName(): string {
  return escapeHtml(process.env.MCP_INSTANCE_NAME ?? 'M365 MCP');
}

function shell(title: string, content: string): string {
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>${instanceName()} — ${escapeHtml(title)}</title>
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
  body { font: 16px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui, sans-serif;
         max-width: 560px; margin: 3em auto; padding: 0 1.5em; color: #1a1a1a; }
  h1 { font-size: 1.4em; }
  .who { background: #f4f4f6; border-radius: 6px; padding: 0.8em 1em; }
  .warn { color: #a4262c; }
  .note { color: #a4262c; font-weight: 600; }
  input[name=code] { font: 1.4em "SF Mono", Menlo, Monaco, Consolas, monospace;
                     letter-spacing: 0.15em; width: 9em; padding: 0.3em; text-transform: uppercase; }
  button { font-size: 1em; padding: 0.5em 1.4em; border-radius: 6px; border: 1px solid #0078d4; cursor: pointer; }
  .primary { background: #0078d4; color: #fff; }
  .secondary { background: #fff; color: #0078d4; }
  form { margin: 1.2em 0; }
</style>
</head>
<body>
<h1>${escapeHtml(title)}</h1>
${content}
</body>
</html>`;
}

function formPage(stored: StoredInstallHandoff, note?: string): string {
  const r = stored.record;
  const device = r.deviceLabel
    ? `an installer on a device that calls itself <strong>${escapeHtml(r.deviceLabel)}</strong>`
    : 'an installer';
  return shell(
    'Confirm this installer',
    `<p class="who">Signed in as <strong>${escapeHtml(r.displayName)}</strong> &lt;${escapeHtml(r.email)}&gt;</p>
<p>This sign-in was started by ${device}. Entering its code connects it to ${instanceName()} as you, with your access to mail, files and calendar.</p>
<p>Type the confirmation code the installer is showing.</p>
<p class="warn">If you did not just run an installer yourself, do not enter a code that someone sends you. Choose "This wasn't me".</p>
${note ? `<p class="note">${escapeHtml(note)}</p>` : ''}
<form method="post" action="${INSTALL_CONFIRM_PATH}">
  <input name="code" autocomplete="off" autocapitalize="characters" spellcheck="false" maxlength="12" placeholder="XXXX-XXXX" aria-label="Confirmation code" required autofocus>
  <button class="primary" type="submit">Connect</button>
</form>
<form method="post" action="${INSTALL_CONFIRM_PATH}">
  <input type="hidden" name="action" value="cancel">
  <button class="secondary" type="submit">This wasn't me</button>
</form>`
  );
}

function messagePage(title: string, text: string): string {
  return shell(title, `<p>${escapeHtml(text)}</p>`);
}

function gonePage(): string {
  return messagePage(
    'Request expired',
    'This sign-in request has expired or was already used. Run the installer again.'
  );
}

app.http('install-confirm', {
  methods: ['GET', 'POST'],
  authLevel: 'anonymous',
  route: 'api/auth/install-confirm',
  handler: withSecurity(withRateLimit('install-confirm', installConfirm)),
});
