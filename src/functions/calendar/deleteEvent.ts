import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { checkCalendarAccess } from '../../services/calendarAccess.js';
import { assertOpaqueId, encodeGraphId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function deleteEventHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const eventId = request.params['eventId'];
  if (!eventId) {
    return { status: 400, jsonBody: { error: 'eventId path param is required' } };
  }
  assertOpaqueId(eventId, 'eventId');

  const calendarId = request.query.get('calendarId');
  if (calendarId) assertOpaqueId(calendarId, 'calendarId');
  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Deny-list enforcement (by ID and name) for the explicit or default calendar.
  const tenantId = await getTenantId(userId);
  const violation = await checkCalendarAccess(graph, tenantId, userId, calendarId ?? undefined);
  if (violation) {
    return { status: violation.status, jsonBody: { error: violation.error } };
  }

  const encEventId = encodeGraphId(eventId, 'eventId');
  const apiPath = calendarId
    ? `/me/calendars/${encodeGraphId(calendarId, 'calendarId')}/events/${encEventId}`
    : `/me/events/${encEventId}`;

  await graph.api(apiPath).delete();

  return { status: 200, jsonBody: { status: 'deleted', eventId } };
}

app.http('deleteEvent', {
  methods: ['DELETE'],
  authLevel: 'anonymous',
  route: 'api/calendar/events/{eventId}',
  handler: withSecurity(withPolicyEnforcement('calendar', deleteEventHandler, {
    mutating: true,
    getDenyListPaths: (req) => {
      const calendarId = req.query.get('calendarId');
      return calendarId ? [calendarId] : [];
    },
  })),
});
