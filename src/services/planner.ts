/**
 * Microsoft Planner helpers for the planner MCP tools.
 *
 * Planner differs from the rest of Graph in three ways the tools have to absorb:
 *
 *   - Every update and delete needs the item's current ETag in `If-Match`, and a
 *     task's description and checklist live on a separate `details` resource with
 *     its own ETag. The tools read the ETag immediately before writing rather than
 *     making the agent carry one between calls. A write that races another editor
 *     still fails with Graph's 412, which reaches the agent as an actionable error.
 *   - Status and priority are numbers (`percentComplete` 0/50/100, `priority`
 *     0-10). The tools speak the names the Planner UI shows and map them here.
 *   - Assignments are a map keyed by user ID, and an assignee is removed by
 *     PATCHing their key to null.
 *
 * The deny list keys Planner on plans: an entry naming a plan's ID or its title
 * blocks every bucket and task in it. Buckets and tasks are resolved to their
 * plan before anything is read or written.
 */
import type { Client } from '@microsoft/microsoft-graph-client';
import { filterDeniedPaths, type DenySubject } from './denyList.js';
import { collectPage } from './graphPaging.js';
import { encodeGraphId } from './opaqueId.js';
import { ToolCallError } from './toolCallError.js';

/** Planner UI status names and the percentComplete each one writes. */
export const PLANNER_STATUS_VALUES = ['notStarted', 'inProgress', 'completed'] as const;
export type PlannerStatus = (typeof PLANNER_STATUS_VALUES)[number];
const STATUS_PERCENT = new Map<string, number>([['notStarted', 0], ['inProgress', 50], ['completed', 100]]);

/**
 * Planner UI priority names and the numeric priority each one writes. Graph
 * buckets 0-1 as urgent, 2-4 important, 5-7 medium and 8-10 low; the values
 * here are the ones the Planner UI itself writes.
 */
export const PLANNER_PRIORITY_VALUES = ['urgent', 'important', 'medium', 'low'] as const;
const PRIORITY_NUMBER = new Map<string, number>([['urgent', 1], ['important', 3], ['medium', 5], ['low', 9]]);

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type GraphObject = Record<string, any>;

/** Order hint that places a new bucket or assignment first. */
const FIRST_ORDER_HINT = ' !';

export function statusFromPercent(percent: unknown): PlannerStatus {
  const n = typeof percent === 'number' ? percent : 0;
  if (n >= 100) return 'completed';
  if (n > 0) return 'inProgress';
  return 'notStarted';
}

export function priorityLabel(priority: unknown): string | null {
  if (typeof priority !== 'number') return null;
  if (priority <= 1) return 'urgent';
  if (priority <= 4) return 'important';
  if (priority <= 7) return 'medium';
  return 'low';
}

function statusToPercent(status: unknown): number {
  const n = typeof status === 'string' ? STATUS_PERCENT.get(status) : undefined;
  if (n === undefined) throw new ToolCallError(`status must be one of ${PLANNER_STATUS_VALUES.join(', ')}`);
  return n;
}

function priorityToNumber(priority: unknown): number {
  const n = typeof priority === 'string' ? PRIORITY_NUMBER.get(priority) : undefined;
  if (n === undefined) throw new ToolCallError(`priority must be one of ${PLANNER_PRIORITY_VALUES.join(', ')}`);
  return n;
}

/**
 * A caller-supplied start or due date as Graph wants it, or null to clear it.
 *
 * Accepted shapes are ISO 8601 only: a bare date (`2026-11-10`), or a date-time
 * with minutes and optional seconds and fraction (`2026-11-10T09:30`,
 * `2026-11-10T09:30:00.000Z`, `2026-11-10T09:30:00-08:00`). A date-time with no
 * offset is read as UTC rather than the server's local zone. The calendar date
 * and the time are range-checked, so an impossible date such as `2026-02-31` is
 * refused instead of rolling over into March the way `Date.parse` would.
 *
 * A bare date is pinned to 12:00 UTC. Planner shows dates in the viewer's local
 * time, so midnight UTC would display as the previous day for anyone west of
 * Greenwich; noon keeps the calendar date from UTC-11 to UTC+11. An empty string
 * clears the date.
 */
const ISO_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})$/;
// eslint-disable-next-line security/detect-unsafe-regex -- anchored, every group is fixed-width or bounded (\d{1,7}) and no two quantifiers can match the same characters, so matching is linear
const ISO_TIME_RE = /^(\d{2}):(\d{2})(:\d{2})?(\.\d{1,7})?(Z|[+-]\d{2}:\d{2})?$/;

function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

export function plannerDate(value: unknown, field: string): string | null {
  if (value === '') return null;
  const invalid = () => new ToolCallError(`${field} must be an ISO 8601 date (YYYY-MM-DD) or date-time (YYYY-MM-DDTHH:MM[:SS][Z|±HH:MM])`);
  if (typeof value !== 'string') throw invalid();
  const [datePart, timePart, ...rest] = value.split('T');
  const m = ISO_DATE_RE.exec(datePart);
  const t = timePart === undefined ? null : ISO_TIME_RE.exec(timePart);
  if (!m || rest.length > 0 || (timePart !== undefined && !t)) throw invalid();
  const [, y, mo, d] = m;
  const [, hh, mi, sec, frac, offset] = t ?? [];
  const ss = sec?.slice(1);
  const year = Number(y);
  const month = Number(mo);
  const day = Number(d);
  if (month < 1 || month > 12 || day < 1 || day > daysInMonth(year, month)) throw invalid();
  if (hh === undefined) return new Date(Date.UTC(year, month - 1, day, 12)).toISOString();
  if (Number(hh) > 23 || Number(mi) > 59 || (ss !== undefined && Number(ss) > 59)) throw invalid();
  if (offset && offset !== 'Z' && (Number(offset.slice(1, 3)) > 23 || Number(offset.slice(4, 6)) > 59)) throw invalid();
  const ms = Date.parse(`${y}-${mo}-${d}T${hh}:${mi}:${ss ?? '00'}${frac ?? ''}${offset ?? 'Z'}`);
  if (Number.isNaN(ms)) throw invalid();
  return new Date(ms).toISOString();
}

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function assigneeList(value: unknown): string[] {
  if (!Array.isArray(value) || value.some((v) => typeof v !== 'string' || !GUID_RE.test(v.trim()))) {
    throw new ToolCallError('assigneeIds must be an array of Entra user object IDs or email addresses');
  }
  return [...new Set((value as string[]).map((v) => v.trim().toLowerCase()))];
}

/**
 * Turn the caller's assignees into Entra object IDs, which is what Planner keys
 * assignments on. An agent usually knows people by address, so an entry with an
 * `@` is looked up through `/users/{upn}`; anything else must already be an ID.
 */
export async function resolveAssignees(graph: Client, value: unknown): Promise<string[]> {
  if (!Array.isArray(value) || value.some((v) => typeof v !== 'string' || !v.trim())) {
    throw new ToolCallError('assigneeIds must be an array of Entra user object IDs or email addresses');
  }
  const ids: string[] = [];
  for (const raw of value as string[]) {
    const entry = raw.trim();
    if (!entry.includes('@')) {
      ids.push(entry);
      continue;
    }
    let user: GraphObject;
    try {
      user = await graph.api(`/users/${encodeURIComponent(entry)}`).select('id').get();
    } catch (err) {
      throw new ToolCallError(`No user found for assignee ${entry}`, { cause: err });
    }
    if (typeof user?.id !== 'string') throw new ToolCallError(`No user found for assignee ${entry}`);
    ids.push(user.id);
  }
  return assigneeList(ids);
}

function assignment() {
  return { '@odata.type': '#microsoft.graph.plannerAssignment', orderHint: FIRST_ORDER_HINT };
}

/** Task fields the create and update tools accept, before mapping. */
export interface PlannerTaskInput {
  title?: unknown;
  bucketId?: unknown;
  startDateTime?: unknown;
  dueDateTime?: unknown;
  status?: unknown;
  priority?: unknown;
  assigneeIds?: unknown;
}

/**
 * The Graph body for a task create or update. `currentAssignees` is the task's
 * assignee set on update: anyone in it and not in `assigneeIds` is removed by
 * setting their key to null, which is how Planner unassigns.
 */
export function plannerTaskBody(input: PlannerTaskInput, currentAssignees: string[] = []): Record<string, unknown> {
  const body: Record<string, unknown> = {};
  if (input.title !== undefined) {
    if (typeof input.title !== 'string' || !input.title.trim()) throw new ToolCallError('title must be a non-empty string');
    body.title = input.title;
  }
  if (input.bucketId !== undefined) body.bucketId = input.bucketId;
  if (input.startDateTime !== undefined) body.startDateTime = plannerDate(input.startDateTime, 'startDateTime');
  if (input.dueDateTime !== undefined) body.dueDateTime = plannerDate(input.dueDateTime, 'dueDateTime');
  if (input.status !== undefined) body.percentComplete = statusToPercent(input.status);
  if (input.priority !== undefined) body.priority = priorityToNumber(input.priority);
  if (input.assigneeIds !== undefined) {
    const wanted = assigneeList(input.assigneeIds);
    const assignments = new Map<string, unknown>();
    const current = new Set(currentAssignees.map((id) => id.toLowerCase()));
    for (const id of currentAssignees) if (!wanted.includes(id.toLowerCase())) assignments.set(id, null);
    for (const id of wanted) if (!current.has(id)) assignments.set(id, assignment());
    if (assignments.size > 0) body.assignments = Object.fromEntries(assignments);
  }
  return body;
}

export function shapePlan(p: GraphObject) {
  return {
    id: p.id,
    title: p.title,
    containerId: p.container?.containerId ?? p.owner ?? null,
    containerType: p.container?.type ?? (p.owner ? 'group' : null),
    createdDateTime: p.createdDateTime ?? null,
  };
}

export function shapeBucket(b: GraphObject) {
  return { id: b.id, name: b.name, planId: b.planId, orderHint: b.orderHint ?? null };
}

export function shapeTask(t: GraphObject) {
  return {
    id: t.id,
    title: t.title,
    planId: t.planId,
    bucketId: t.bucketId ?? null,
    status: statusFromPercent(t.percentComplete),
    percentComplete: t.percentComplete ?? 0,
    priority: priorityLabel(t.priority),
    startDateTime: t.startDateTime ?? null,
    dueDateTime: t.dueDateTime ?? null,
    completedDateTime: t.completedDateTime ?? null,
    createdDateTime: t.createdDateTime ?? null,
    assigneeIds: Object.keys(t.assignments ?? {}),
    hasDescription: Boolean(t.hasDescription),
    checklistItemCount: t.checklistItemCount ?? 0,
  };
}

export function shapeTaskDetails(d: GraphObject) {
  const checklist = Object.entries((d.checklist ?? {}) as Record<string, GraphObject>).map(([id, item]) => ({
    id,
    title: item?.title ?? '',
    isChecked: Boolean(item?.isChecked),
  }));
  return { description: d.description ?? '', checklist };
}

/** The plan a bucket or task belongs to, with the item's ETag for a later write. */
export interface PlannerItem {
  item: GraphObject;
  etag: string;
}

function etagOf(item: GraphObject, what: string): string {
  const etag = item['@odata.etag'];
  if (typeof etag !== 'string' || !etag) {
    throw new ToolCallError(`Microsoft Graph returned the ${what} without an ETag, so it cannot be changed safely`);
  }
  return etag;
}

/** Read a task with its ETag. */
export async function getPlannerTask(graph: Client, taskId: string): Promise<PlannerItem> {
  const item: GraphObject = await graph.api(`/planner/tasks/${encodeGraphId(taskId, 'taskId')}`).get();
  return { item, etag: etagOf(item, 'task') };
}

/** Read a task's details (description, checklist) with their ETag. */
export async function getPlannerTaskDetails(graph: Client, taskId: string): Promise<PlannerItem> {
  const item: GraphObject = await graph.api(`/planner/tasks/${encodeGraphId(taskId, 'taskId')}/details`).get();
  return { item, etag: etagOf(item, 'task details') };
}

/** Read a bucket with its ETag. */
export async function getPlannerBucket(graph: Client, bucketId: string): Promise<PlannerItem> {
  const item: GraphObject = await graph.api(`/planner/buckets/${encodeGraphId(bucketId, 'bucketId')}`).get();
  return { item, etag: etagOf(item, 'bucket') };
}

/**
 * Set a task's description. Planner keeps it on the details resource, which
 * has its own ETag, so this is a read followed by a conditional PATCH.
 */
export async function setPlannerTaskDescription(graph: Client, taskId: string, description: string): Promise<void> {
  const { etag } = await getPlannerTaskDetails(graph, taskId);
  await graph
    .api(`/planner/tasks/${encodeGraphId(taskId, 'taskId')}/details`)
    .header('If-Match', etag)
    .patch({ description });
}

/**
 * Whether the deny list blocks a plan, matching the entry against the plan's ID
 * and its title. Fails closed: filterDeniedPaths returns nothing when storage is
 * down, which reads here as denied.
 */
async function planDenied(tenantId: string, subject: DenySubject, plan: { id: string; title?: string }): Promise<boolean> {
  const kept = await filterDeniedPaths(tenantId, subject, 'planner', [{ id: plan.id, name: plan.title }]);
  return kept.length === 0;
}

/**
 * Throw unless the plan is reachable under the deny list. Reads the plan's title
 * so an entry recorded by name is honoured as well as one recorded by ID.
 */
export async function assertPlanAllowed(graph: Client, tenantId: string, subject: DenySubject, planId: string): Promise<void> {
  const plan: GraphObject = await graph.api(`/planner/plans/${encodeGraphId(planId, 'planId')}`).get();
  if (await planDenied(tenantId, subject, { id: planId, title: plan?.title })) {
    throw new ToolCallError('Access restricted by deny list — the plan is blocked');
  }
}

/**
 * Drop plans the deny list blocks.
 */
export async function filterDeniedPlans<T extends { id: string; title?: string }>(
  tenantId: string,
  subject: DenySubject,
  plans: T[],
): Promise<T[]> {
  const kept = await filterDeniedPaths(tenantId, subject, 'planner', plans.map((p) => ({ id: p.id, name: p.title, plan: p })));
  return kept.map((k) => k.plan);
}

/**
 * How many plans the cross-plan task listing will read one at a time when the
 * user's own plan collection doesn't include them. Past this, the remaining
 * plans count as unreadable and their tasks are dropped.
 */
export const MAX_PLAN_LOOKUPS = 10;

/** How many plans the cross-plan task listing reads from `/me/planner/plans`. */
const PLAN_INDEX_LIMIT = 1000;

/**
 * Drop tasks whose plan the deny list blocks, for the cross-plan "assigned to
 * me" listing.
 *
 * Plan titles come from one paged read of `/me/planner/plans`, joined by plan
 * ID, so the Graph request count doesn't grow with the number of tasks. A plan
 * missing from that collection is read on its own, up to MAX_PLAN_LOOKUPS of
 * them. A plan that still has no title (unreadable, or past the cap) is treated
 * as blocked, so its tasks are dropped rather than returned unchecked. The deny
 * list itself is consulted once for every plan together.
 */
export async function filterTasksByAllowedPlan<T extends { planId?: string }>(
  graph: Client,
  tenantId: string,
  subject: DenySubject,
  tasks: T[],
): Promise<T[]> {
  const planIds = [...new Set(tasks.map((t) => t.planId).filter((id): id is string => typeof id === 'string' && id.length > 0))];
  if (planIds.length === 0) return [];

  const titles = new Map<string, string | undefined>();
  try {
    const page = await collectPage<GraphObject>(
      () => graph.api('/me/planner/plans').get(),
      (link) => graph.api(link).get(),
      0,
      PLAN_INDEX_LIMIT,
    );
    for (const plan of page.items) if (typeof plan?.id === 'string') titles.set(plan.id, plan.title);
  } catch {
    // Fall through to the bounded per-plan reads below.
  }

  let lookups = 0;
  for (const planId of planIds) {
    if (titles.has(planId)) continue;
    if (lookups >= MAX_PLAN_LOOKUPS) break;
    lookups++;
    try {
      const plan: GraphObject = await graph.api(`/planner/plans/${encodeGraphId(planId, 'planId')}`).get();
      titles.set(planId, plan?.title);
    } catch {
      // Unreadable: left out of `titles`, so its tasks are dropped.
    }
  }

  const known = planIds.filter((id) => titles.has(id)).map((id) => ({ id, title: titles.get(id) }));
  const allowed = new Set((await filterDeniedPlans(tenantId, subject, known)).map((p) => p.id));
  return tasks.filter((t) => typeof t.planId === 'string' && allowed.has(t.planId));
}

/** Body for a new bucket, placed first in the plan. */
export function plannerBucketBody(planId: string, name: unknown): Record<string, unknown> {
  if (typeof name !== 'string' || !name.trim()) throw new ToolCallError('name must be a non-empty string');
  return { name, planId, orderHint: FIRST_ORDER_HINT };
}
