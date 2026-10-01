import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { checkCalendarAccess } from '../../services/calendarAccess.js';
import { assertOpaqueId, encodeGraphId } from '../../services/opaqueId.js';
import { resolveWindow } from '../../services/calendarWindow.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function listEventsHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const calendarId = request.query.get('calendarId');
  const startDateTime = request.query.get('startDateTime') ?? undefined;
  const endDateTime = request.query.get('endDateTime') ?? undefined;
  const timeZone = request.query.get('timeZone') ?? 'America/Los_Angeles';
  const maxResults = parseInt(request.query.get('maxResults') ?? '25', 10);

  if (calendarId) assertOpaqueId(calendarId, 'calendarId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Deny-list enforcement (by ID and name) for the explicit or default calendar.
  const tenantId = await getTenantId(userId);
  const violation = await checkCalendarAccess(graph, tenantId, userId, calendarId ?? undefined);
  if (violation) {
    return { status: violation.status, jsonBody: { error: violation.error } };
  }

  const select = 'id,subject,start,end,location,organizer,attendees,isAllDay,webLink';
  const top = Math.min(maxResults, 50);

  // With a date range, use calendarView so recurring series are expanded into
  // their concrete instances on those days. Filtering /events by start/dateTime
  // matches on the series MASTER's original start (often years past), so it
  // can never answer "what is on this specific day".
  const window = resolveWindow(startDateTime, endDateTime);
  let result;
  if (window) {
    const base = calendarId ? `/me/calendars/${encodeGraphId(calendarId, 'calendarId')}/calendarView` : '/me/calendarView';
    const qs = `startDateTime=${encodeURIComponent(window.start)}&endDateTime=${encodeURIComponent(window.end)}`;
    result = await graph
      .api(`${base}?${qs}`)
      .header('Prefer', `outlook.timezone="${timeZone}"`)
      .select(select)
      .top(top)
      .orderby('start/dateTime')
      .get();
  } else {
    const apiPath = calendarId ? `/me/calendars/${encodeGraphId(calendarId, 'calendarId')}/events` : '/me/events';
    result = await graph.api(apiPath).select(select).top(top).orderby('start/dateTime').get();
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const events = (result.value ?? []).map((e: any) => ({
    id: e.id,
    subject: e.subject,
    start: e.start,
    end: e.end,
    location: e.location?.displayName ?? null,
    organizer: e.organizer?.emailAddress?.address ?? null,
    attendees: (e.attendees ?? []).map((a: any) => a.emailAddress?.address),
    isAllDay: e.isAllDay ?? false,
    webLink: e.webLink,
  }));

  return { status: 200, jsonBody: { events, count: events.length } };
}

app.http('listEvents', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/calendar/events',
  handler: withSecurity(withPolicyEnforcement('calendar', listEventsHandler, {
    getDenyListPaths: (req) => {
      // Only check explicit calendarId at the middleware level; default calendar
      // resolution is handled inside the handler.
      const calendarId = req.query.get('calendarId');
      return calendarId ? [calendarId] : [];
    },
  })),
});
