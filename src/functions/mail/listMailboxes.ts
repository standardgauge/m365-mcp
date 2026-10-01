import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface MailboxInfo {
  id: string;
  displayName: string;
  email: string;
  type: 'primary' | 'shared' | 'delegated';
}

async function listMailboxesHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Primary mailbox — always available via /me
  const me = await graph
    .api('/me')
    .select('id,displayName,mail,userPrincipalName')
    .get();

  const mailboxes: MailboxInfo[] = [
    {
      id: me.id,
      displayName: me.displayName,
      email: me.mail ?? me.userPrincipalName,
      type: 'primary',
    },
  ];

  // Shared / delegated mailboxes the user can send-as or has full-access to.
  // The /me/people endpoint gives us collaborators; for actual shared-mailbox
  // resolution we use findRooms / people with the mailbox hint.
  // A reliable approach for delegated access is to enumerate via Exchange
  // admin-scoped calls, but with delegated permissions (no app perms) we
  // query the user's own mailboxSettings to detect any auto-mapped mailboxes.
  try {
    // Look up shared mailboxes via the directory — requires Mail.ReadBasic
    const sharedResult = await graph
      .api('/me/mailboxSettings')
      .get();

    // mailboxSettings doesn't list shared mailboxes directly; it's a
    // configuration bag. We include it in the response for completeness but
    // the real shared-mailbox listing would need an Exchange-specific call or
    // tenant-scoped admin consent. Leave a well-typed placeholder so callers
    // know this is extensible.
    void sharedResult; // used for future extension
  } catch {
    // Delegated permissions may not include mailboxSettings — ignore
  }

  return {
    status: 200,
    jsonBody: {
      mailboxes,
      count: mailboxes.length,
      note: 'Shared/delegated mailboxes require Exchange admin consent or explicit delegation grants. Only the primary mailbox is listed via delegated permissions.',
    },
  };
}

app.http('listMailboxes', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/mail/mailboxes',
  handler: withSecurity(withPolicyEnforcement('mail', listMailboxesHandler)),
});
