import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { filterDeniedPaths } from '../../services/denyList.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

interface CalendarItem {
  id: string;
  name: string;
  color: string;
  isDefaultCalendar: boolean;
}

async function listCalendarsHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;
  const accessToken = await getValidAccessTokenForSession(auth.session);

  const graph = createGraphClient(accessToken);

  const result = await graph
    .api('/me/calendars')
    .select('id,name,color,isDefaultCalendar')
    .top(100)
    .get();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const calendars: CalendarItem[] = (result.value ?? []).map((c: any) => ({
    id: c.id,
    name: c.name,
    color: c.color ?? 'auto',
    isDefaultCalendar: c.isDefaultCalendar ?? false,
  }));

  const allowed = await filterDeniedPaths(await getTenantId(userId), userId, 'calendar', calendars);

  return { status: 200, jsonBody: { calendars: allowed, count: allowed.length } };
}

app.http('listCalendars', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/calendar/calendars',
  handler: withSecurity(withPolicyEnforcement('calendar', listCalendarsHandler)),
});
