import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedPaths } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface TeamItem {
  id: string;
  name: string;
  description: string | null;
}

async function listTeamsHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const graph = createGraphClient(accessToken);

  const result = await graph
    .api('/me/joinedTeams')
    .select('id,displayName,description')
    .get();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const teams: TeamItem[] = (result.value ?? []).map((t: any) => ({
    id: t.id,
    name: t.displayName,
    description: t.description ?? null,
  }));

  const allowed = await filterDeniedPaths(await getTenantId(userId), userId, 'teams', teams);

  return { status: 200, jsonBody: { teams: allowed, count: allowed.length } };
}

app.http('listTeams', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/teams/teams',
  handler: withSecurity(withPolicyEnforcement('teams', listTeamsHandler)),
});
