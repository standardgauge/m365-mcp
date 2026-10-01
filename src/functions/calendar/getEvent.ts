import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { checkCalendarAccess } from '../../services/calendarAccess.js';
import { assertOpaqueId, encodeGraphId } from '../../services/opaqueId.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function getEventHandler(
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
  const tenantId = await getTenantId(userId);

  // Deny-list enforcement (by ID and name) for the explicit or default calendar.
  const violation = await checkCalendarAccess(graph, tenantId, userId, calendarId ?? undefined);
  if (violation) {
    return { status: violation.status, jsonBody: { error: violation.error } };
  }

  const encEventId = encodeGraphId(eventId, 'eventId');
  const apiPath = calendarId
    ? `/me/calendars/${encodeGraphId(calendarId, 'calendarId')}/events/${encEventId}`
    : `/me/events/${encEventId}`;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const event: any = await graph
    .api(apiPath)
    .select('id,subject,start,end,location,organizer,attendees,isAllDay,body,webLink')
    .get();

  return {
    status: 200,
    jsonBody: {
      id: event.id,
      subject: event.subject,
      start: event.start,
      end: event.end,
      location: event.location?.displayName ?? null,
      organizer: event.organizer?.emailAddress?.address ?? null,
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      attendees: (event.attendees ?? []).map((a: any) => a.emailAddress?.address),
      isAllDay: event.isAllDay ?? false,
      body: event.body,
      webLink: event.webLink,
    },
  };
}

app.http('getEvent', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/calendar/events/{eventId}',
  handler: withSecurity(withPolicyEnforcement('calendar', getEventHandler, {
    getDenyListPaths: (req) => {
      // Only check explicit calendarId at the middleware level; default calendar
      // resolution + name-based deny are handled inside the handler.
      const calendarId = req.query.get('calendarId');
      return calendarId ? [calendarId] : [];
    },
  })),
});
