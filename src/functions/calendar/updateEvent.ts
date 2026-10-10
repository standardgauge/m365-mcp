import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getValidAccessTokenForSession, getTenantId } from '../../services/tokenCache.js';
import { createGraphClient } from '../../services/graphClient.js';
import { withPolicyEnforcement } from '../../services/policyEnforcement.js';
import { checkCalendarAccess } from '../../services/calendarAccess.js';
import { assertOpaqueId, encodeGraphId } from '../../services/opaqueId.js';
import { resolveMailboxTimeZone } from '../../services/mailboxTimeZone.js';
import type { AuthResult } from '../../services/authMiddleware.js';
import { withSecurity } from '../../services/securityHeaders.js';
import { logAccess } from '../../services/auditLog.js';
import { auditClientAddress } from '../../services/clientAddress.js';
import {
  enforceEventUpdatePolicy,
  OutboundPolicyError,
} from '../../services/outboundPolicy.js';

// Graph's free/busy status enum for an event (event.showAs). Writable on both
// POST and PATCH; omitting it leaves the event's current status untouched.
type ShowAs = 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere' | 'unknown';

interface UpdateEventRequest {
  subject?: string;
  start?: string;
  end?: string;
  timeZone?: string;
  location?: string;
  body?: string;
  bodyType?: 'text' | 'html';
  attendees?: string[];
  isAllDay?: boolean;
  showAs?: ShowAs;
}

async function updateEventHandler(
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

  const reqBody = (await request.json()) as UpdateEventRequest;
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

  // Resolve a time zone only when start/end is actually being changed. Honor an
  // explicit timeZone, else default to the mailbox's own zone; never
  // guess a hardcoded one, and skip the mailbox lookup entirely when no time is
  // being rescheduled.
  const tz = (reqBody.start !== undefined || reqBody.end !== undefined)
    ? (reqBody.timeZone ?? (await resolveMailboxTimeZone(graph)))
    : reqBody.timeZone;
  const patch: Record<string, unknown> = {};

  if (reqBody.subject !== undefined) patch.subject = reqBody.subject;
  if (reqBody.isAllDay !== undefined) patch.isAllDay = reqBody.isAllDay;
  if (reqBody.start !== undefined) patch.start = { dateTime: reqBody.start, timeZone: tz };
  if (reqBody.end !== undefined) patch.end = { dateTime: reqBody.end, timeZone: tz };
  if (reqBody.location !== undefined) patch.location = { displayName: reqBody.location };
  if (reqBody.body !== undefined) {
    patch.body = {
      contentType: reqBody.bodyType === 'html' ? 'HTML' : 'Text',
      content: reqBody.body,
    };
  }
  if (reqBody.attendees !== undefined) {
    patch.attendees = reqBody.attendees.map((addr) => ({
      emailAddress: { address: addr },
      type: 'required',
    }));
  }
  // Free/busy status. Only patch it when supplied so an omitted showAs leaves
  // the event's current status untouched.
  if (reqBody.showAs !== undefined) patch.showAs = reqBody.showAs;

  const encEventId = encodeGraphId(eventId, 'eventId');
  const apiPath = calendarId
    ? `/me/calendars/${encodeGraphId(calendarId, 'calendarId')}/events/${encEventId}`
    : `/me/events/${encEventId}`;

  // Outbound policy: an organizer's edit goes to every attendee as an update.
  try {
    await enforceEventUpdatePolicy(graph, tenantId, userId, apiPath, patch, reqBody.attendees);
  } catch (err: unknown) {
    if (err instanceof OutboundPolicyError) return outboundDenied(auth, tenantId, request, err);
    throw err;
  }

  const result = await graph.api(apiPath).patch(patch);

  return {
    status: 200,
    jsonBody: {
      id: result.id,
      subject: result.subject,
      start: result.start,
      end: result.end,
      status: 'updated',
    },
  };
}

/** Refuse with 403 and record the refusal, as the policy wrapper does for its own checks. */
function outboundDenied(auth: AuthResult, tenantId: string, request: HttpRequest, err: OutboundPolicyError): HttpResponseInit {
  logAccess({
    tenantId,
    userId: auth.userId,
    userEmail: auth.session.email,
    deviceLabel: auth.session.deviceLabel,
    operation: 'calendar.patch',
    result: 'denied',
    reason: err.message,
    source: 'http',
    ip: auditClientAddress(request),
  });
  return { status: 403, jsonBody: { error: err.message } };
}

app.http('updateEvent', {
  methods: ['PATCH'],
  authLevel: 'anonymous',
  route: 'api/calendar/events/{eventId}',
  handler: withSecurity(withPolicyEnforcement('calendar', updateEventHandler, {
    mutating: true,
    getDenyListPaths: (req) => {
      const calendarId = req.query.get('calendarId');
      return calendarId ? [calendarId] : [];
    },
  })),
});
