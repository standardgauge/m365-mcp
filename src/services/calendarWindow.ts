/**
 * Helpers for building well-formed Microsoft Graph calendar time windows.
 *
 * Graph's calendarView, getSchedule, and findMeetingTimes all reject or
 * silently mis-handle loosely-formed datetime inputs:
 *
 *   - A bare calendar date (`2026-09-08`) has no time component, so the window
 *     duration is undefined. getSchedule answers this with the opaque error
 *     "The time duration specified for FreeBusyViewOptions.TimeWindow is
 *     invalid"; calendarView/findMeetingTimes quietly return nothing useful.
 *   - A datetime that carries an absolute-time marker (a trailing `Z` or a
 *     numeric `±HH:MM` offset) contradicts the separate `timeZone` field these
 *     APIs also take, and can make Graph reject the window the same way.
 *
 * These helpers normalize an input to the naive-local form Graph expects,
 * which is then paired with an explicit `timeZone` (in the body for
 * getSchedule / findMeetingTimes, or via the `Prefer: outlook.timezone`
 * header for calendarView).
 *
 * See: the M365 connector's three calendar-read verbs were
 * unusable for a specific-day availability check because of these input-shape
 * failures (compounded, for list_events, by querying /events — recurring
 * series masters — instead of /calendarView — expanded instances).
 */

const ABSOLUTE_MARKER = /(Z|[+-]\d{2}:?\d{2})$/i;
const DATE_ONLY = /^\d{4}-\d{2}-\d{2}$/;

/**
 * Normalize a caller-supplied datetime to the naive-local ISO 8601 form Graph
 * calendar APIs expect alongside an explicit timeZone.
 *
 * - Strips a trailing `Z` or numeric offset (contradicts the timeZone field).
 * - Expands a bare date to the correct edge of that day: start → `T00:00:00`,
 *   end → `T23:59:59`.
 */
export function normalizeGraphDateTime(value: string, edge: 'start' | 'end'): string {
  let v = String(value).trim();
  v = v.replace(ABSOLUTE_MARKER, '');
  if (DATE_ONLY.test(v)) v += edge === 'start' ? 'T00:00:00' : 'T23:59:59';
  return v;
}

/**
 * Shift a naive-local ISO datetime by whole days. Used to default the missing
 * edge of a one-sided window. Returns the input unchanged if it can't be
 * parsed.
 */
export function shiftDays(isoLocal: string, days: number): string {
  const d = new Date(`${isoLocal}Z`);
  if (Number.isNaN(d.getTime())) return isoLocal;
  d.setUTCDate(d.getUTCDate() + days);
  // Drop the milliseconds + `Z` to return the same naive-local shape as the input.
  return d.toISOString().replace(/\.\d{3}Z$/, '');
}

/**
 * Resolve a calendar window from optional start/end bounds.
 *
 * calendarView and a findMeetingTimes timeConstraint both require *both* edges,
 * so when only one bound is supplied the other is defaulted `defaultSpanDays`
 * away — a single-sided request still yields a bounded window. Returns null
 * when neither bound is supplied (the caller decides what an open request
 * means).
 */
export function resolveWindow(
  start: string | undefined,
  end: string | undefined,
  defaultSpanDays = 7,
): { start: string; end: string } | null {
  let s = start ? normalizeGraphDateTime(start, 'start') : null;
  let e = end ? normalizeGraphDateTime(end, 'end') : null;
  if (!s && !e) return null;
  if (s && !e) e = shiftDays(s, defaultSpanDays);
  if (e && !s) s = shiftDays(e, -defaultSpanDays);
  return { start: s as string, end: e as string };
}
