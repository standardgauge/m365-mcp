import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { resolveDenySubject } from '../../services/mailboxOwner.js';
import { isPathDenied } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { resolveMailFolderName } from '../../services/containerResolver.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { searchMail, toMessageSummary, outcomeMeta, type MessageSummary } from '../../services/mailSearch.js';
import { resolveMaxResults, toEnvelope } from '../../services/resultEnvelope.js';

/**
 * GET /api/mail/search — the REST twin of the `search_mail` MCP tool.
 *
 * Query params: `q`, `participant`, `from`, `to`, `since`, `folderId`, `mailboxId`,
 * `maxResults`. At least one of `q` / `participant` / `from` / `to` is required.
 * Routing, matching and the honest-empty-result metadata are shared with the MCP
 * tool via services/mailSearch.ts: folder-scoped queries never use
 * `$search`, and address criteria in a folder run a newest-first scan so a
 * counterparty's address matches in Sent Items.
 */
async function searchMailHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const q = request.query.get('q') ?? undefined;
  const participant = request.query.get('participant') ?? undefined;
  const from = request.query.get('from') ?? undefined;
  const to = request.query.get('to') ?? undefined;
  const since = request.query.get('since') ?? undefined;
  const folderId = request.query.get('folderId');
  const mailboxId = request.query.get('mailboxId') ?? 'me';
  const maxResults = resolveMaxResults(request.query.get('maxResults') ?? undefined, 25, 50);

  if (![q, participant, from, to].some((v) => v && v.trim())) {
    return { status: 400, jsonBody: { error: 'At least one of q, participant, from, to is required' } };
  }
  if (mailboxId !== 'me') assertOpaqueId(mailboxId, 'mailboxId');
  if (folderId) assertOpaqueId(folderId, 'folderId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);
  const denySubject = await resolveDenySubject(graph, userId, mailboxId);

  const base = mailboxId === 'me' ? '/me' : `/users/${mailboxId}`;

  // Resolve folder display name for deny-list check (deny list stores display names, not IDs)
  if (folderId) {
    const tenantId = await getTenantId(userId);
    const folderName = await resolveMailFolderName(graph, folderId, base);
    if (folderName && await isPathDenied(tenantId, denySubject, 'mail', folderName)) {
      return { status: 403, jsonBody: { error: 'Access to this mail folder is restricted by the deny list' } };
    }
  }

  let outcome;
  try {
    outcome = await searchMail(graph, {
      base,
      folderId: folderId ?? undefined,
      maxResults,
      q,
      participant,
      from,
      to,
      since,
    });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    // Caller-side errors (bad `since`, no criteria) are 400; a Graph failure is 502.
    if (!msg.startsWith('Mail search failed')) {
      return { status: 400, jsonBody: { error: msg } };
    }
    context.error(`[searchMail] ${msg}`);
    return { status: 502, jsonBody: { error: msg } };
  }

  const tenantId = await getTenantId(userId);
  const allMessages = outcome.messages.map(toMessageSummary);

  // Post-filter: resolve folder display names and check against deny list
  let messages = allMessages;
  if (!folderId) {
    const filtered: MessageSummary[] = [];
    for (const msg of allMessages) {
      if (msg.folderId) {
        const folderName = await resolveMailFolderName(graph, msg.folderId, base);
        if (folderName && await isPathDenied(tenantId, denySubject, 'mail', folderName)) continue;
      }
      filtered.push(msg);
    }
    messages = filtered;
  }

  const { items, count, limit, truncated } = toEnvelope(messages, maxResults, outcome.moreAvailable);
  return { status: 200, jsonBody: { results: items, count, limit, truncated, ...outcomeMeta(outcome) } };
}

app.http('searchMail', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/mail/search',
  // Deny-list check is handled inside the handler because the deny list stores
  // folder display names, not Graph IDs — the handler resolves names via Graph.
  handler: withSecurity(withPolicyEnforcement('mail', searchMailHandler)),
});
