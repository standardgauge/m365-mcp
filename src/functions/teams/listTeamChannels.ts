import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedPaths } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { assertOpaqueId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface ChannelItem {
  id: string;
  displayName: string;
  membershipType: string;
}

async function listTeamChannelsHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const teamId = request.params.teamId;
  if (!teamId) {
    return { status: 400, jsonBody: { error: 'teamId path parameter is required' } };
  }
  assertOpaqueId(teamId, 'teamId');

  const graph = createGraphClient(accessToken);

  // Graph GET /teams/{teamId}/channels does not support $top — appending it fails the whole request.
  const result = await graph
    .api(`/teams/${teamId}/channels`)
    .select('id,displayName,membershipType')
    .get();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const allChannels = (result.value ?? []).map((c: any) => ({
    id: c.id,
    displayName: c.displayName,
    // Use channel ID as the path for deny-list filtering — the admin UI
    // stores channel IDs (not display names) as deny entries.
    path: c.id,
    membershipType: c.membershipType ?? 'standard',
  }));

  // Post-filter channels against the deny list
  const tenantId = await getTenantId(userId);
  const allowed = await filterDeniedPaths(tenantId, userId, 'teams', allChannels);
  // Strip internal path field from response
  const channels: ChannelItem[] = allowed.map((c) => ({
    id: c.id!,
    displayName: (c as typeof allChannels[number]).displayName,
    membershipType: (c as typeof allChannels[number]).membershipType,
  }));

  return { status: 200, jsonBody: { channels, count: channels.length } };
}

app.http('listTeamChannels', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/teams/teams/{teamId}/channels',
  handler: withSecurity(withPolicyEnforcement('teams', listTeamChannelsHandler, {
    getDenyListPaths: (req) => {
      const teamId = req.params.teamId;
      return teamId ? [teamId] : [];
    },
  })),
});
