/**
 * Typed calendar attendees: the input create_event / update_event accept, the
 * recipients the outbound policy sees, and the shape get_event returns.
 */

import { toGraphAttendees, shapeAttendee } from '../services/calendarAttendees.js';
import { attendeeRecipients } from '../services/outboundPolicy.js';
import { ToolCallError } from '../services/toolCallError.js';

describe('toGraphAttendees', () => {
  it('keeps a bare address as a required attendee', () => {
    expect(toGraphAttendees(['a@example.com'])).toEqual([
      { emailAddress: { address: 'a@example.com' }, type: 'required' },
    ]);
  });

  it('honours the type on the object form, defaulting to required', () => {
    expect(toGraphAttendees([
      { email: 'car@example.com', type: 'resource' },
      { email: 'b@example.com', type: 'optional' },
      { email: 'c@example.com' },
      'd@example.com',
    ])).toEqual([
      { emailAddress: { address: 'car@example.com' }, type: 'resource' },
      { emailAddress: { address: 'b@example.com' }, type: 'optional' },
      { emailAddress: { address: 'c@example.com' }, type: 'required' },
      { emailAddress: { address: 'd@example.com' }, type: 'required' },
    ]);
  });

  it('trims addresses', () => {
    expect(toGraphAttendees([{ email: '  room@example.com ', type: 'resource' }])[0].emailAddress.address).toBe('room@example.com');
  });

  it('accepts an empty array, which clears attendees on update', () => {
    expect(toGraphAttendees([])).toEqual([]);
  });

  it.each([
    ['an unknown type', [{ email: 'a@example.com', type: 'chair' }], /type "chair"/],
    ['a missing email', [{ type: 'resource' }], /attendees\[0\] needs a non-empty email/],
    ['an empty string', [''], /needs a non-empty email/],
    ['a number', [42], /must be an email address or \{ email, type \}/],
    ['a nested array', [['a@example.com']], /must be an email address/],
    ['an unsupported field', [{ email: 'a@example.com', role: 'resource' }], /unsupported field\(s\): role/],
    ['a duplicate address', ['a@example.com', { email: 'A@example.com', type: 'optional' }], /more than once/],
    ['a non-array', 'a@example.com', /must be an array/],
  ])('refuses %s with a ToolCallError', (_label, input, message) => {
    expect(() => toGraphAttendees(input)).toThrow(ToolCallError);
    expect(() => toGraphAttendees(input)).toThrow(message);
  });
});

describe('attendeeRecipients', () => {
  it('reads the address from both attendee forms, so the outbound policy sees every recipient', () => {
    expect(attendeeRecipients(['a@example.com', { email: 'car@example.com', type: 'resource' }])).toEqual([
      { address: 'a@example.com' },
      { address: 'car@example.com' },
    ]);
  });

  it('yields a null address for an entry it cannot read, which the policy treats as external', () => {
    expect(attendeeRecipients([{ type: 'resource' }])).toEqual([{ address: null }]);
  });
});

describe('shapeAttendee', () => {
  it('returns email, name, type and response', () => {
    expect(shapeAttendee({
      type: 'resource',
      status: { response: 'accepted', time: '2026-07-01T09:00:00Z' },
      emailAddress: { name: 'Pool car', address: 'car@example.com' },
    })).toEqual({ email: 'car@example.com', name: 'Pool car', type: 'resource', response: 'accepted', respondedAt: '2026-07-01T09:00:00Z' });
  });

  it("maps Graph's minimum-date placeholder to no response time", () => {
    expect(shapeAttendee({
      type: 'required', status: { response: 'none', time: '0001-01-01T00:00:00Z' }, emailAddress: { address: 'a@example.com' },
    })).toMatchObject({ response: 'none', respondedAt: null, name: null });
  });

  it('tolerates a sparse attendee', () => {
    expect(shapeAttendee({})).toEqual({ email: null, name: null, type: null, response: null, respondedAt: null });
  });
});
