/**
 * Outbound policy: admin control over the write paths that put agent-written
 * content in front of other people without a draft step.
 *
 * Enforced draft mode (userEmailSettings.ts) holds `send_mail` and `send_draft`.
 * Three other channels deliver at once and are not covered by it:
 *
 *   calendarInvites  `create_event` / `update_event` / `move_event` on a meeting
 *                    the user organizes. Graph sends the invitation, or the
 *                    update, to every attendee as soon as the event is written,
 *                    carrying whatever subject and body the agent supplied.
 *   eventResponses   `respond_to_event` with a `comment`. The comment goes to the
 *                    organizer with the RSVP. A response without a comment is a
 *                    fixed accept/tentative/decline and is not gated.
 *   teamsMessages    `send_chat_message` and `send_channel_message`.
 *
 * Each channel takes one of three modes:
 *
 *   allow     no restriction (the default, and the behaviour before this policy).
 *   internal  refused when any recipient is outside the tenant: an address whose
 *             domain is not one of the tenant's verified domains, or a Teams
 *             member whose home tenant is another one.
 *   block     refused whenever anyone other than the user would be notified.
 *
 * There is no "hold" mode. Graph has no draft state for a meeting invitation or
 * a chat message, so the only way to put a person in the loop is to refuse the
 * tool call and have the user act in Outlook or Teams, which is what `block`
 * does. That is also how enforced draft mode treats `send_draft`.
 *
 * Table: OutboundPolicy
 *   PartitionKey: tenantId
 *   RowKey:       TENANT_POLICY_ROW_KEY for the tenant-wide policy, or userId
 *   calendarInvites / eventResponses / teamsMessages: OutboundMode
 *   updatedAt / updatedBy: audit trail of the last change
 *
 * The effective mode for a user is the stricter of the tenant row and the
 * user's row, per channel, so a per-user row can tighten the tenant policy but
 * never loosen it.
 *
 * Fail closed: a policy read that fails for any reason other than "no row",
 * a tenant-domain lookup that fails or returns nothing, and a recipient list
 * that cannot be read all refuse the call rather than let it through.
 */

import { TableClient, TableEntity } from '@azure/data-tables';
import type { Client } from '@microsoft/microsoft-graph-client';

const TABLE = 'OutboundPolicy';
/** RowKey of the tenant-wide row; same value and reasoning as the draft-mode policy. */
const TENANT_POLICY_ROW_KEY = '__tenant__';
const AZURITE_CONNECTION_STRING = 'UseDevelopmentStorage=true';

export type OutboundChannel = 'calendarInvites' | 'eventResponses' | 'teamsMessages';
export type OutboundMode = 'allow' | 'internal' | 'block';

export const OUTBOUND_CHANNELS: readonly OutboundChannel[] = ['calendarInvites', 'eventResponses', 'teamsMessages'];
export const OUTBOUND_MODES: readonly OutboundMode[] = ['allow', 'internal', 'block'];

/** Strictness order, used to resolve the tenant and user rows. */
const RANK: Record<OutboundMode, number> = { allow: 0, internal: 1, block: 2 };

export type OutboundModes = Record<OutboundChannel, OutboundMode>;

export interface OutboundPolicy extends OutboundModes {
  /** ISO-8601 timestamp of the last change */
  updatedAt?: string;
  /** userId of the admin who last changed the policy */
  updatedBy?: string;
}

export type OutboundPolicyScope =
  | { scope: 'tenant' }
  | { scope: 'user'; userId: string };

/** Which row set the effective mode; null when it is 'allow'. */
export type OutboundEnforcedBy = 'tenant' | 'user' | null;

export interface EffectiveOutboundMode {
  mode: OutboundMode;
  enforcedBy: OutboundEnforcedBy;
}

/**
 * Text every refusal starts with. The MCP dispatcher matches on it to log the
 * refusal as denied, the way it does for the enforced draft mode marker.
 */
export const OUTBOUND_POLICY_MARKER = "Blocked by your administrator's outbound policy";

/** A refusal under the outbound policy. Maps to 403 on REST routes. */
export class OutboundPolicyError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'OutboundPolicyError';
  }
}

const CHANNEL_LABEL: Record<OutboundChannel, string> = {
  calendarInvites: 'calendar invitations',
  eventResponses: 'comments on meeting responses',
  teamsMessages: 'Teams messages',
};

const CHANNEL_HUMAN_PATH: Record<OutboundChannel, string> = {
  calendarInvites: 'Create or update the meeting in Outlook instead.',
  eventResponses: 'Respond without a comment, or add the comment from Outlook.',
  teamsMessages: 'Send the message from Teams instead.',
};

function defaultModes(): OutboundModes {
  return { calendarInvites: 'allow', eventResponses: 'allow', teamsMessages: 'allow' };
}

function isMode(value: unknown): value is OutboundMode {
  return typeof value === 'string' && (OUTBOUND_MODES as readonly string[]).includes(value);
}

function getTableClient(): TableClient {
  const conn = process.env.AZURE_STORAGE_CONNECTION_STRING ?? AZURITE_CONNECTION_STRING;
  const allowInsecureConnection =
    conn === 'UseDevelopmentStorage=true' || conn.includes('DefaultEndpointsProtocol=http;');
  return TableClient.fromConnectionString(conn, TABLE, { allowInsecureConnection });
}

async function ensureTable(): Promise<void> {
  try {
    await getTableClient().createTable();
  } catch (err: unknown) {
    if ((err as { statusCode?: number }).statusCode !== 409) throw err;
  }
}

function rowKeyFor(target: OutboundPolicyScope): string {
  return target.scope === 'tenant' ? TENANT_POLICY_ROW_KEY : target.userId;
}

/**
 * Read one policy row. A missing row means every channel is 'allow'; any other
 * failure propagates so the caller fails closed. A stored value that is not a
 * known mode reads as 'block': an unreadable restriction is not a permission.
 */
async function readRow(tenantId: string, rowKey: string): Promise<OutboundPolicy> {
  try {
    const entity = await getTableClient().getEntity<Record<string, unknown>>(tenantId, rowKey);
    const policy: OutboundPolicy = defaultModes();
    for (const channel of OUTBOUND_CHANNELS) {
      const value = entity[channel];
      if (value === undefined || value === null) continue;
      policy[channel] = isMode(value) ? value : 'block';
    }
    if (typeof entity.updatedAt === 'string') policy.updatedAt = entity.updatedAt;
    if (typeof entity.updatedBy === 'string') policy.updatedBy = entity.updatedBy;
    return policy;
  } catch (err: unknown) {
    if ((err as { statusCode?: number }).statusCode === 404) return defaultModes();
    throw err;
  }
}

/** Read one policy as stored, without resolving precedence. */
export async function getOutboundPolicy(tenantId: string, target: OutboundPolicyScope): Promise<OutboundPolicy> {
  await ensureTable();
  return readRow(tenantId, rowKeyFor(target));
}

/**
 * Set some or all channels of the tenant or user policy. Channels not named in
 * `modes` keep their stored value. Writes the whole row so the audit fields
 * always describe the last change.
 */
export async function setOutboundPolicy(
  tenantId: string,
  target: OutboundPolicyScope,
  modes: Partial<OutboundModes>,
  setBy: string,
): Promise<OutboundPolicy> {
  await ensureTable();
  const current = await readRow(tenantId, rowKeyFor(target));
  const policy: OutboundPolicy = {
    calendarInvites: modes.calendarInvites ?? current.calendarInvites,
    eventResponses: modes.eventResponses ?? current.eventResponses,
    teamsMessages: modes.teamsMessages ?? current.teamsMessages,
    updatedAt: new Date().toISOString(),
    updatedBy: setBy,
  };
  const entity: TableEntity = { partitionKey: tenantId, rowKey: rowKeyFor(target), ...policy };
  await getTableClient().upsertEntity(entity, 'Replace');
  return policy;
}

/** The tenant row, the user's row, and the stricter of the two per channel. */
export async function getOutboundEnforcement(
  tenantId: string,
  userId: string,
): Promise<{ tenant: OutboundPolicy; user: OutboundPolicy; effective: Record<OutboundChannel, EffectiveOutboundMode> }> {
  await ensureTable();
  const [tenant, user] = await Promise.all([
    readRow(tenantId, TENANT_POLICY_ROW_KEY),
    readRow(tenantId, userId),
  ]);
  const effective = {} as Record<OutboundChannel, EffectiveOutboundMode>;
  for (const channel of OUTBOUND_CHANNELS) {
    const t = tenant[channel];
    const u = user[channel];
    // Tenant wins a tie, so an admin reading the result sees the tenant-wide
    // setting named as the reason rather than an incidental per-user copy.
    if (RANK[t] >= RANK[u]) effective[channel] = { mode: t, enforcedBy: t === 'allow' ? null : 'tenant' };
    else effective[channel] = { mode: u, enforcedBy: 'user' };
  }
  return { tenant, user, effective };
}

// ── Tenant domains ─────────────────────────────────────────────

const DOMAIN_TTL_MS = 10 * 60 * 1000;
const domainCache = new Map<string, { domains: Set<string>; expires: number }>();

/** Test hook: forget cached tenant domains. */
export function clearTenantDomainCache(): void {
  domainCache.clear();
}

/**
 * The tenant's verified domains, read with the caller's own token from
 * `/organization` (covered by User.Read). Cached per tenant for ten minutes.
 * Throws when the list cannot be read or is empty, so 'internal' fails closed.
 */
export async function getTenantDomains(graph: Client, tenantId: string): Promise<Set<string>> {
  const cached = domainCache.get(tenantId);
  if (cached && cached.expires > Date.now()) return cached.domains;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const res: any = await graph.api('/organization').select('id,verifiedDomains').get();
  const domains = new Set<string>();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  for (const org of (res?.value ?? []) as any[]) {
    for (const d of org?.verifiedDomains ?? []) {
      if (typeof d?.name === 'string' && d.name.trim()) domains.add(d.name.trim().toLowerCase());
    }
  }
  if (domains.size === 0) {
    throw new OutboundPolicyError(
      `${OUTBOUND_POLICY_MARKER}: the organization's domains could not be read, so recipients cannot be checked.`,
    );
  }
  domainCache.set(tenantId, { domains, expires: Date.now() + DOMAIN_TTL_MS });
  return domains;
}

/** Lowercased domain of an address, or null when it is not an address. */
export function addressDomain(address: unknown): string | null {
  if (typeof address !== 'string') return null;
  const at = address.trim().lastIndexOf('@');
  if (at <= 0 || at === address.trim().length - 1) return null;
  return address.trim().slice(at + 1).toLowerCase();
}

/**
 * Someone who would be notified. `address` is an SMTP address; `tenantId` is
 * the home tenant Teams reports for a chat or channel member.
 */
export interface OutboundRecipient {
  address?: string | null;
  tenantId?: string | null;
}

/**
 * Internal means: the domain of the address is one of the tenant's verified
 * domains, and, where a home tenant is known, it is this tenant. A recipient
 * with neither an address nor a home tenant is treated as external. Domain
 * matching is exact: a subdomain counts only if it is itself verified.
 */
export function isInternalRecipient(r: OutboundRecipient, tenantId: string, domains: Set<string>): boolean {
  if (r.tenantId && r.tenantId !== tenantId) return false;
  const domain = addressDomain(r.address);
  if (domain) return domains.has(domain);
  // No address: a member of this tenant (by home tenant) with no mail is
  // internal; anything else cannot be shown to be.
  return Boolean(r.tenantId) && r.tenantId === tenantId;
}

export interface OutboundCheck {
  graph: Client;
  tenantId: string;
  userId: string;
  channel: OutboundChannel;
  /**
   * Who would be notified. Called only when the policy is not 'allow', so the
   * Graph reads it may need cost nothing for an unrestricted user. Returning
   * an empty list means nothing goes out and the call proceeds.
   */
  recipients: () => Promise<OutboundRecipient[]>;
  /**
   * Set when the call always reaches someone (a Teams post), so 'block' can
   * refuse without reading the recipient list first.
   */
  alwaysNotifies?: boolean;
}

/**
 * Refuse the call when the outbound policy forbids it. Resolves when the call
 * may proceed; throws OutboundPolicyError otherwise.
 */
export async function enforceOutboundPolicy(check: OutboundCheck): Promise<void> {
  const { effective } = await getOutboundEnforcement(check.tenantId, check.userId);
  const { mode, enforcedBy } = effective[check.channel];
  if (mode === 'allow') return;
  const scope = enforcedBy === 'user' ? 'for your account' : 'for this organization';
  const blocked = () => new OutboundPolicyError(
    `${OUTBOUND_POLICY_MARKER}: ${CHANNEL_LABEL[check.channel]} cannot be sent by this tool ${scope}. ` +
    CHANNEL_HUMAN_PATH[check.channel],
  );
  if (mode === 'block' && check.alwaysNotifies) throw blocked();

  let recipients: OutboundRecipient[];
  try {
    recipients = await check.recipients();
  } catch (err: unknown) {
    if (err instanceof OutboundPolicyError) throw err;
    const detail = err instanceof Error ? err.message : String(err);
    throw new OutboundPolicyError(
      `${OUTBOUND_POLICY_MARKER}: the recipients of these ${CHANNEL_LABEL[check.channel]} could not be read (${detail}), so the call was refused.`,
    );
  }
  if (recipients.length === 0) return;

  if (mode === 'block') throw blocked();

  const domains = await getTenantDomains(check.graph, check.tenantId);
  const external = recipients
    .filter((r) => !isInternalRecipient(r, check.tenantId, domains))
    .map((r) => r.address || (r.tenantId ? `a member of tenant ${r.tenantId}` : 'an unidentified member'));
  if (external.length > 0) {
    const shown = external.slice(0, 5).join(', ') + (external.length > 5 ? `, and ${external.length - 5} more` : '');
    throw new OutboundPolicyError(
      `${OUTBOUND_POLICY_MARKER}: ${CHANNEL_LABEL[check.channel]} can only reach people inside the organization ${scope}, ` +
      `and this would reach ${shown}. ${CHANNEL_HUMAN_PATH[check.channel]}`,
    );
  }
}

// ── Calendar helpers shared by the MCP tools and the REST routes ──

/** Event fields whose change makes Graph send an update to the attendees. */
const NOTIFYING_EVENT_FIELDS = ['subject', 'body', 'start', 'end', 'location', 'isAllDay', 'attendees'];

/** True when an event PATCH would send an update to the event's attendees. */
export function eventPatchNotifies(patch: Record<string, unknown>): boolean {
  return NOTIFYING_EVENT_FIELDS.some((f) => patch[f] !== undefined);
}

/** Map `create_event` style attendee strings to recipients. */
export function attendeeRecipients(addresses: unknown): OutboundRecipient[] {
  return Array.isArray(addresses) ? addresses.map((a) => ({ address: typeof a === 'string' ? a : null })) : [];
}

/** Map Graph `attendees` objects to recipients. */
export function graphAttendeeRecipients(attendees: unknown): OutboundRecipient[] {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  return Array.isArray(attendees) ? attendees.map((a: any) => ({ address: a?.emailAddress?.address ?? null })) : [];
}

/**
 * Check an event update. Graph notifies attendees only when the user organizes
 * the event and the patch touches a notifying field; everyone on the current
 * list and on the new one is counted, since a removed attendee still gets a
 * cancellation.
 */
export async function enforceEventUpdatePolicy(
  graph: Client,
  tenantId: string,
  userId: string,
  eventPath: string,
  patch: Record<string, unknown>,
  newAttendees: unknown,
): Promise<void> {
  if (!eventPatchNotifies(patch)) return;
  await enforceOutboundPolicy({
    graph,
    tenantId,
    userId,
    channel: 'calendarInvites',
    recipients: async () => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const current: any = await graph.api(eventPath).select('attendees,isOrganizer').get();
      if (current?.isOrganizer === false) return [];
      return [...graphAttendeeRecipients(current?.attendees), ...attendeeRecipients(newAttendees)];
    },
  });
}

// ── Teams helpers ──────────────────────────────────────────────

// Bound on member pages followed; a chat or channel this large is refused
// under 'internal' rather than partly checked.
const MAX_MEMBER_PAGES = 20;

/**
 * Read every member of a chat or channel as recipients. Reading members needs
 * a permission the Teams send tools do not (ChatMember.Read for a chat,
 * ChannelMember.Read.All for a channel); `permission` is named in the error so
 * an administrator who set 'internal' without granting it sees why every send
 * is refused.
 */
export async function teamsMemberRecipients(
  graph: Client,
  membersPath: string,
  permission?: string,
): Promise<OutboundRecipient[]> {
  const out: OutboundRecipient[] = [];
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let page: any = await graph.api(membersPath).get();
    for (let pages = 1; ; pages++) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const m of (page?.value ?? []) as any[]) {
        out.push({ address: m?.email ?? null, tenantId: m?.tenantId ?? null });
      }
      const next = page?.['@odata.nextLink'];
      if (!next) break;
      if (pages >= MAX_MEMBER_PAGES) throw new Error('member list too long to check');
      page = await graph.api(next).get();
    }
  } catch (err: unknown) {
    const status = (err as { statusCode?: number }).statusCode;
    if (permission && (status === 401 || status === 403)) {
      throw new Error(`reading members needs the ${permission} permission on the app registration`);
    }
    throw err;
  }
  return out;
}
