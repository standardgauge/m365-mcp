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
  enforceOutboundPolicy,
  attendeeRecipients,
  OutboundPolicyError,
} from '../../services/outboundPolicy.js';

// Graph's free/busy status enum for an event (event.showAs). Writable on both
// POST and PATCH; omitting it leaves Graph's own default (busy for timed
// events, free for all-day).
type ShowAs = 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere' | 'unknown';

interface EventRequest {
  subject: string;
  start: string;
  end: string;
  timeZone?: string;
  location?: string;
  body?: string;
  bodyType?: 'text' | 'html';
  attendees?: string[];
  isAllDay?: boolean;
  showAs?: ShowAs;
  calendarId?: string;
}

async function createEventHandler(
  request: HttpRequest,
  context: InvocationContext,
  auth: AuthResult,
): Promise<HttpResponseInit> {
  const userId = auth.userId;

  const body = (await request.json()) as EventRequest;
  if (!body.subject || !body.start || !body.end) {
    return { status: 400, jsonBody: { error: 'subject, start, and end are required' } };
  }
  if (body.calendarId) assertOpaqueId(body.calendarId, 'calendarId');

  const accessToken = await getValidAccessTokenForSession(auth.session);
  const graph = createGraphClient(accessToken);

  // Deny-list enforcement (by ID and name) for the explicit or default calendar.
  // calendarId is body-level here, so it is not visible to the wrapper.
  const tenantId = await getTenantId(userId);
  const violation = await checkCalendarAccess(graph, tenantId, userId, body.calendarId);
  if (violation) {
    return { status: violation.status, jsonBody: { error: violation.error } };
  }

  // Honor an explicit timeZone; otherwise default to the mailbox's own
  // configured zone rather than a hardcoded one. Fails loud if the
  // zone can't be resolved — never guesses.
  const tz = body.timeZone ?? (await resolveMailboxTimeZone(graph));
  const event: Record<string, unknown> = {
    subject: body.subject,
    start: { dateTime: body.start, timeZone: tz },
    end: { dateTime: body.end, timeZone: tz },
    isAllDay: body.isAllDay ?? false,
  };

  if (body.location) {
    event.location = { displayName: body.location };
  }
  if (body.body) {
    event.body = {
      contentType: body.bodyType === 'html' ? 'HTML' : 'Text',
      content: body.body,
    };
  }
  if (body.attendees) {
    event.attendees = body.attendees.map((addr) => ({
      emailAddress: { address: addr },
      type: 'required',
    }));
  }
  // Free/busy status. Only set it when supplied so an omitted showAs leaves
  // Graph's default untouched.
  if (body.showAs) {
    event.showAs = body.showAs;
  }

  const apiPath = body.calendarId
    ? `/me/calendars/${encodeGraphId(body.calendarId, 'calendarId')}/events`
    : '/me/events';

  // Outbound policy: attendees get the invitation the moment the event is written.
  try {
    await enforceOutboundPolicy({
      graph, tenantId, userId, channel: 'calendarInvites',
      recipients: async () => attendeeRecipients(body.attendees),
    });
  } catch (err: unknown) {
    if (err instanceof OutboundPolicyError) return outboundDenied(auth, tenantId, request, err);
    throw err;
  }

  const result = await graph.api(apiPath).post(event);

  return {
    status: 201,
    jsonBody: {
      id: result.id,
      subject: result.subject,
      start: result.start,
      end: result.end,
      webLink: result.webLink,
      status: 'created',
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
    operation: 'calendar.post',
    result: 'denied',
    reason: err.message,
    source: 'http',
    ip: auditClientAddress(request),
  });
  return { status: 403, jsonBody: { error: err.message } };
}

app.http('createEvent', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'api/calendar/events',
  handler: withSecurity(withPolicyEnforcement('calendar', createEventHandler, { mutating: true })),
});
