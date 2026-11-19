/**
 * Calendar attendees: the typed input create_event and update_event accept, and
 * the shape get_event returns.
 *
 * Graph sorts an event's attendees into three types. `required` and `optional`
 * are people. `resource` is a room or equipment mailbox, and it is the type
 * that makes the mailbox's booking policy run: the resource accepts or
 * declines on its own, and a declined booking is visible in the attendee's
 * response. A resource added as `required` is sent an ordinary invitation and
 * nothing books it.
 *
 * Each attendee is either a bare address, which stays `required` so existing
 * callers keep working, or `{ email, type }`.
 */
import { ToolCallError } from './toolCallError.js';

export const ATTENDEE_TYPES = ['required', 'optional', 'resource'] as const;
export type AttendeeType = (typeof ATTENDEE_TYPES)[number];

/** JSON Schema for one attendee item, shared by create_event and update_event. */
export const ATTENDEE_ITEM_SCHEMA: Record<string, unknown> = {
  anyOf: [
    { type: 'string', description: 'Attendee email address, invited as required.' },
    {
      type: 'object',
      properties: {
        email: { type: 'string', description: 'Attendee email address.' },
        type: {
          type: 'string',
          enum: [...ATTENDEE_TYPES],
          description: 'required (default), optional, or resource for a room or equipment mailbox.',
        },
      },
      required: ['email'],
      additionalProperties: false,
    },
  ],
};

export const ATTENDEES_DESCRIPTION =
  'Attendees. Each is an email address (invited as required) or { email, type } with type required, optional, ' +
  'or resource. Book a room or equipment mailbox (a car, a projector) as type resource so its booking policy ' +
  'accepts or declines it; get_event then shows its response.';

export interface GraphAttendee {
  emailAddress: { address: string };
  type: AttendeeType;
}

function isAttendeeType(value: unknown): value is AttendeeType {
  return typeof value === 'string' && (ATTENDEE_TYPES as readonly string[]).includes(value);
}

/**
 * Convert the tool's attendees argument into Graph attendee objects. Throws a
 * ToolCallError naming the offending entry rather than dropping or guessing.
 */
export function toGraphAttendees(input: unknown): GraphAttendee[] {
  if (!Array.isArray(input)) throw new ToolCallError('attendees must be an array');
  const seen = new Set<string>();
  return input.map((entry, i) => {
    let email: unknown;
    let type: unknown = 'required';
    if (typeof entry === 'string') {
      email = entry;
    } else if (entry && typeof entry === 'object' && !Array.isArray(entry)) {
      const extra = Object.keys(entry).filter((k) => k !== 'email' && k !== 'type');
      if (extra.length > 0) {
        throw new ToolCallError(`attendees[${i}] has unsupported field(s): ${extra.join(', ')}. Use { email, type }.`);
      }
      ({ email } = entry as { email?: unknown });
      if ((entry as { type?: unknown }).type !== undefined) type = (entry as { type?: unknown }).type;
    } else {
      throw new ToolCallError(`attendees[${i}] must be an email address or { email, type }`);
    }
    if (typeof email !== 'string' || email.trim() === '') {
      throw new ToolCallError(`attendees[${i}] needs a non-empty email`);
    }
    if (!isAttendeeType(type)) {
      throw new ToolCallError(`attendees[${i}] has type ${JSON.stringify(type)}; use one of ${ATTENDEE_TYPES.join(', ')}`);
    }
    const address = email.trim();
    const key = address.toLowerCase();
    if (seen.has(key)) throw new ToolCallError(`attendees lists ${address} more than once`);
    seen.add(key);
    return { emailAddress: { address }, type };
  });
}

/** One attendee as get_event returns it. */
export interface AttendeeSummary {
  email: string | null;
  name: string | null;
  type: string | null;
  /** Graph's response: none, organizer, tentativelyAccepted, accepted, declined, notResponded. */
  response: string | null;
  respondedAt: string | null;
}

/** Graph reports "no response yet" with the minimum date rather than leaving the time out. */
const NO_RESPONSE_TIME = /^0001-01-01/;

export function shapeAttendee(a: unknown): AttendeeSummary {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const g = (a ?? {}) as any;
  const time = typeof g.status?.time === 'string' && !NO_RESPONSE_TIME.test(g.status.time) ? g.status.time : null;
  return {
    email: g.emailAddress?.address ?? null,
    name: g.emailAddress?.name ?? null,
    type: g.type ?? null,
    response: g.status?.response ?? null,
    respondedAt: time,
  };
}
