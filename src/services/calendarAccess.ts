/**
 * Calendar access-control helper.
 *
 * Consolidates the calendar deny-list enforcement that every calendar
 * operation must apply, so the HTTP routes and the MCP tool handlers cannot
 * drift apart and open a bypass. A single check covers:
 *
 *   1. The explicit `calendarId` when one is supplied, OR the user's default
 *      calendar (reached via `/me/events`) when it is omitted.
 *   2. Deny by calendar ID — the opaque Graph identifier.
 *   3. Deny by calendar NAME — admins add deny-list entries for shared or
 *      sensitive calendars by their human-readable name (e.g. "HR",
 *      "Executive"), so we resolve the effective calendar's name and check it
 *      too. Without this, a name-based deny entry would only block get_event
 *      while list/create/update/delete stayed open.
 *
 * The caller MUST validate any user-supplied `calendarId` with assertOpaqueId
 * (HTTP) or assertOpaqueIds (MCP) before calling this — the value flows into
 * resolveCalendarName, which interpolates it into a Graph path.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { isPathDenied } from './denyList.js';
import { resolveDefaultCalendarId, resolveCalendarName } from './containerResolver.js';
import type { PolicyViolation } from './policyEnforcement.js';

/**
 * Returns a 403 PolicyViolation when the effective calendar is blocked by the
 * deny list (by ID or by name), or null when access is allowed.
 */
export async function checkCalendarAccess(
  graph: Client,
  tenantId: string,
  userId: string,
  calendarId?: string,
): Promise<PolicyViolation | null> {
  const effectiveCalId = calendarId || (await resolveDefaultCalendarId(graph, userId));
  if (!effectiveCalId) return null;

  // The default-calendar message names the fact that the block is on the
  // default calendar; explicit-id access uses the generic message so a caller
  // learns only that access was refused, not the deny-list contents.
  const error = calendarId
    ? 'Access restricted by deny list'
    : 'Access to the default calendar is restricted by the deny list';

  if (await isPathDenied(tenantId, userId, 'calendar', effectiveCalId)) {
    return { status: 403, error };
  }

  const calName = await resolveCalendarName(graph, effectiveCalId);
  if (calName && (await isPathDenied(tenantId, userId, 'calendar', calName))) {
    return { status: 403, error };
  }

  return null;
}
