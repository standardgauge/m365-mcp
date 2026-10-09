/**
 * Unit tests for the outbound policy service.
 *
 * Covers:
 *   - effective mode is the stricter of the tenant and user rows, per channel
 *   - a missing row reads as 'allow'; an unknown stored value reads as 'block'
 *   - a policy read that fails for anything but 404 propagates (fail closed)
 *   - setOutboundPolicy merges with the stored row and stamps the audit fields
 *   - enforceOutboundPolicy: allow never reads recipients; block refuses when
 *     anyone is notified and passes when nobody is; internal compares against the
 *     tenant's verified domains and the Teams home tenant; unreadable recipients
 *     and an empty domain list refuse
 *   - enforceEventUpdatePolicy: only notifying fields on an organizer's event count
 *   - teamsMemberRecipients follows nextLink and refuses an unbounded chain
 */

import { jest } from '@jest/globals';
import type { Client } from '@microsoft/microsoft-graph-client';

const mockGetEntity = jest.fn<(partitionKey: string, rowKey: string) => Promise<unknown>>();
const mockUpsertEntity = jest.fn<(entity: Record<string, unknown>, mode: string) => Promise<void>>();

jest.mock('@azure/data-tables', () => ({
  TableClient: {
    fromConnectionString: () => ({
      getEntity: (pk: string, rk: string) => mockGetEntity(pk, rk),
      upsertEntity: (e: Record<string, unknown>, m: string) => mockUpsertEntity(e, m),
      createTable: () => Promise.resolve(),
    }),
  },
}));

import {
  getOutboundEnforcement,
  setOutboundPolicy,
  enforceOutboundPolicy,
  enforceEventUpdatePolicy,
  teamsMemberRecipients,
  isInternalRecipient,
  clearTenantDomainCache,
  OutboundPolicyError,
  OUTBOUND_POLICY_MARKER,
} from '../services/outboundPolicy.js';

const TENANT = 'tenant-abc';
const OTHER_TENANT = 'tenant-other';
const USER = 'user-xyz';
const NOT_FOUND = { statusCode: 404 };

function rows(tenantRow?: Record<string, unknown>, userRow?: Record<string, unknown>) {
  mockGetEntity.mockImplementation(async (_pk, rk) => {
    const row = rk === '__tenant__' ? tenantRow : userRow;
    if (!row) throw NOT_FOUND;
    return row;
  });
}

/** A Graph client whose GETs are answered by path. */
function fakeGraph(responses: Record<string, unknown>): { graph: Client; gets: string[] } {
  const gets: string[] = [];
  const graph = {
    api: (path: string) => {
      const chain = {
        select: () => chain,
        get: async () => {
          gets.push(path);
          const r = responses[path];
          if (r instanceof Error) throw r;
          if (r === undefined) throw new Error(`unexpected GET ${path}`);
          return r;
        },
      };
      return chain;
    },
  } as unknown as Client;
  return { graph, gets };
}

const ORG = { value: [{ verifiedDomains: [{ name: 'contoso.com' }, { name: 'contoso.onmicrosoft.com' }] }] };

beforeEach(() => {
  jest.clearAllMocks();
  clearTenantDomainCache();
  rows();
});

describe('getOutboundEnforcement', () => {
  it('defaults every channel to allow when there are no rows', async () => {
    const { effective } = await getOutboundEnforcement(TENANT, USER);
    expect(effective.calendarInvites).toEqual({ mode: 'allow', enforcedBy: null });
    expect(effective.teamsMessages).toEqual({ mode: 'allow', enforcedBy: null });
  });

  it('takes the stricter row per channel and names its source', async () => {
    rows(
      { calendarInvites: 'internal', teamsMessages: 'block' },
      { calendarInvites: 'block', teamsMessages: 'allow', eventResponses: 'internal' },
    );
    const { effective } = await getOutboundEnforcement(TENANT, USER);
    expect(effective.calendarInvites).toEqual({ mode: 'block', enforcedBy: 'user' });
    expect(effective.teamsMessages).toEqual({ mode: 'block', enforcedBy: 'tenant' });
    expect(effective.eventResponses).toEqual({ mode: 'internal', enforcedBy: 'user' });
  });

  it('reads an unknown stored value as block', async () => {
    rows({ calendarInvites: 'sometimes' });
    const { effective } = await getOutboundEnforcement(TENANT, USER);
    expect(effective.calendarInvites.mode).toBe('block');
  });

  it('propagates a read failure other than 404', async () => {
    mockGetEntity.mockRejectedValue({ statusCode: 503 });
    await expect(getOutboundEnforcement(TENANT, USER)).rejects.toEqual({ statusCode: 503 });
  });
});

describe('setOutboundPolicy', () => {
  it('keeps channels it was not given and stamps the admin', async () => {
    rows({ calendarInvites: 'block', eventResponses: 'internal', teamsMessages: 'allow' });
    const policy = await setOutboundPolicy(TENANT, { scope: 'tenant' }, { teamsMessages: 'internal' }, 'admin-1');
    expect(policy).toMatchObject({ calendarInvites: 'block', eventResponses: 'internal', teamsMessages: 'internal', updatedBy: 'admin-1' });
    const [entity, mode] = mockUpsertEntity.mock.calls[0];
    expect(mode).toBe('Replace');
    expect(entity).toMatchObject({ partitionKey: TENANT, rowKey: '__tenant__', teamsMessages: 'internal' });
  });

  it('writes a user row under the user id', async () => {
    await setOutboundPolicy(TENANT, { scope: 'user', userId: USER }, { calendarInvites: 'block' }, 'admin-1');
    expect(mockUpsertEntity.mock.calls[0][0]).toMatchObject({ rowKey: USER, calendarInvites: 'block', teamsMessages: 'allow' });
  });
});

describe('isInternalRecipient', () => {
  const domains = new Set(['contoso.com']);
  it('matches verified domains exactly, case-insensitively', () => {
    expect(isInternalRecipient({ address: 'Pat@Contoso.com' }, TENANT, domains)).toBe(true);
    expect(isInternalRecipient({ address: 'pat@mail.contoso.com' }, TENANT, domains)).toBe(false);
    expect(isInternalRecipient({ address: 'pat@fabrikam.com' }, TENANT, domains)).toBe(false);
  });
  it('treats another home tenant as external even on a verified domain', () => {
    expect(isInternalRecipient({ address: 'pat@contoso.com', tenantId: OTHER_TENANT }, TENANT, domains)).toBe(false);
  });
  it('treats a recipient with no address and no tenant as external', () => {
    expect(isInternalRecipient({ address: 'not-an-address' }, TENANT, domains)).toBe(false);
    expect(isInternalRecipient({}, TENANT, domains)).toBe(false);
    expect(isInternalRecipient({ tenantId: TENANT }, TENANT, domains)).toBe(true);
  });
});

describe('enforceOutboundPolicy', () => {
  it('allow: proceeds without reading recipients or domains', async () => {
    const recipients = jest.fn(async () => [{ address: 'x@fabrikam.com' }]);
    const { graph, gets } = fakeGraph({});
    await enforceOutboundPolicy({ graph, tenantId: TENANT, userId: USER, channel: 'calendarInvites', recipients });
    expect(recipients).not.toHaveBeenCalled();
    expect(gets).toEqual([]);
  });

  it('block: refuses when anyone would be notified, with the marker', async () => {
    rows({ calendarInvites: 'block' });
    const { graph } = fakeGraph({});
    const p = enforceOutboundPolicy({
      graph, tenantId: TENANT, userId: USER, channel: 'calendarInvites',
      recipients: async () => [{ address: 'pat@contoso.com' }],
    });
    await expect(p).rejects.toBeInstanceOf(OutboundPolicyError);
    await expect(p).rejects.toThrow(OUTBOUND_POLICY_MARKER);
    await expect(p).rejects.toThrow('for this organization');
  });

  it('block: proceeds when nobody would be notified', async () => {
    rows(undefined, { calendarInvites: 'block' });
    const { graph } = fakeGraph({});
    await expect(enforceOutboundPolicy({
      graph, tenantId: TENANT, userId: USER, channel: 'calendarInvites', recipients: async () => [],
    })).resolves.toBeUndefined();
  });

  it('block with alwaysNotifies: refuses without reading recipients', async () => {
    rows(undefined, { teamsMessages: 'block' });
    const recipients = jest.fn(async () => []);
    const { graph } = fakeGraph({});
    await expect(enforceOutboundPolicy({
      graph, tenantId: TENANT, userId: USER, channel: 'teamsMessages', alwaysNotifies: true, recipients,
    })).rejects.toThrow('for your account');
    expect(recipients).not.toHaveBeenCalled();
  });

  it('internal: passes when every recipient is on a verified domain', async () => {
    rows({ calendarInvites: 'internal' });
    const { graph, gets } = fakeGraph({ '/organization': ORG });
    await enforceOutboundPolicy({
      graph, tenantId: TENANT, userId: USER, channel: 'calendarInvites',
      recipients: async () => [{ address: 'a@contoso.com' }, { address: 'b@contoso.onmicrosoft.com' }],
    });
    expect(gets).toEqual(['/organization']);
  });

  it('internal: refuses and names the external recipients', async () => {
    rows({ calendarInvites: 'internal' });
    const { graph } = fakeGraph({ '/organization': ORG });
    await expect(enforceOutboundPolicy({
      graph, tenantId: TENANT, userId: USER, channel: 'calendarInvites',
      recipients: async () => [{ address: 'a@contoso.com' }, { address: 'eve@fabrikam.com' }],
    })).rejects.toThrow(/inside the organization.*eve@fabrikam\.com/);
  });

  it('internal: caches the domain list per tenant', async () => {
    rows({ calendarInvites: 'internal' });
    const { graph, gets } = fakeGraph({ '/organization': ORG });
    const check = { graph, tenantId: TENANT, userId: USER, channel: 'calendarInvites' as const, recipients: async () => [{ address: 'a@contoso.com' }] };
    await enforceOutboundPolicy(check);
    await enforceOutboundPolicy(check);
    expect(gets).toEqual(['/organization']);
  });

  it('internal: refuses when the domain list is empty or unreadable', async () => {
    rows({ calendarInvites: 'internal' });
    const check = (graph: Client) => enforceOutboundPolicy({
      graph, tenantId: TENANT, userId: USER, channel: 'calendarInvites', recipients: async () => [{ address: 'a@contoso.com' }],
    });
    await expect(check(fakeGraph({ '/organization': { value: [] } }).graph)).rejects.toThrow(OUTBOUND_POLICY_MARKER);
    await expect(check(fakeGraph({ '/organization': new Error('403 Forbidden') }).graph)).rejects.toThrow('403 Forbidden');
  });

  it('refuses when the recipients cannot be read', async () => {
    rows({ teamsMessages: 'internal' });
    const { graph } = fakeGraph({});
    const p = enforceOutboundPolicy({
      graph, tenantId: TENANT, userId: USER, channel: 'teamsMessages',
      recipients: async () => { throw new Error('Forbidden'); },
    });
    await expect(p).rejects.toBeInstanceOf(OutboundPolicyError);
    await expect(p).rejects.toThrow(/could not be read \(Forbidden\)/);
  });

  it('fails closed when the policy cannot be read', async () => {
    mockGetEntity.mockRejectedValue(new Error('storage down'));
    const { graph } = fakeGraph({});
    await expect(enforceOutboundPolicy({
      graph, tenantId: TENANT, userId: USER, channel: 'calendarInvites', recipients: async () => [],
    })).rejects.toThrow('storage down');
  });
});

describe('enforceEventUpdatePolicy', () => {
  const PATH = '/me/events/ev1';

  it('ignores a patch that only changes free/busy', async () => {
    rows({ calendarInvites: 'block' });
    const { graph, gets } = fakeGraph({});
    await enforceEventUpdatePolicy(graph, TENANT, USER, PATH, { showAs: 'busy' }, undefined);
    expect(gets).toEqual([]);
  });

  it('refuses an organizer edit to a meeting with attendees', async () => {
    rows({ calendarInvites: 'block' });
    const { graph } = fakeGraph({ [PATH]: { isOrganizer: true, attendees: [{ emailAddress: { address: 'a@contoso.com' } }] } });
    await expect(enforceEventUpdatePolicy(graph, TENANT, USER, PATH, { subject: 'x' }, undefined)).rejects.toThrow(OUTBOUND_POLICY_MARKER);
  });

  it('allows an edit to an event the user does not organize', async () => {
    rows({ calendarInvites: 'block' });
    const { graph } = fakeGraph({ [PATH]: { isOrganizer: false, attendees: [{ emailAddress: { address: 'a@contoso.com' } }] } });
    await expect(enforceEventUpdatePolicy(graph, TENANT, USER, PATH, { subject: 'x' }, undefined)).resolves.toBeUndefined();
  });

  it('counts both current and new attendees under internal', async () => {
    rows({ calendarInvites: 'internal' });
    const { graph } = fakeGraph({
      [PATH]: { isOrganizer: true, attendees: [{ emailAddress: { address: 'eve@fabrikam.com' } }] },
      '/organization': ORG,
    });
    // The patch replaces the external attendee with an internal one; the removed
    // attendee still gets a cancellation, so the call is refused.
    await expect(enforceEventUpdatePolicy(graph, TENANT, USER, PATH, { attendees: [] }, ['a@contoso.com']))
      .rejects.toThrow('eve@fabrikam.com');
  });
});

describe('teamsMemberRecipients', () => {
  it('follows nextLink and maps email and home tenant', async () => {
    const { graph } = fakeGraph({
      '/chats/c1/members': { value: [{ email: 'a@contoso.com', tenantId: TENANT }], '@odata.nextLink': 'https://graph/next' },
      'https://graph/next': { value: [{ email: null, tenantId: OTHER_TENANT }] },
    });
    expect(await teamsMemberRecipients(graph, '/chats/c1/members')).toEqual([
      { address: 'a@contoso.com', tenantId: TENANT },
      { address: null, tenantId: OTHER_TENANT },
    ]);
  });

  it('names the missing permission on a 403', async () => {
    const forbidden = Object.assign(new Error('Forbidden'), { statusCode: 403 });
    const { graph } = fakeGraph({ '/chats/c1/members': forbidden });
    await expect(teamsMemberRecipients(graph, '/chats/c1/members', 'ChatMember.Read')).rejects.toThrow('ChatMember.Read permission');
  });

  it('refuses a member list that never ends', async () => {
    const { graph } = fakeGraph({
      '/chats/c1/members': { value: [], '@odata.nextLink': 'https://graph/loop' },
      'https://graph/loop': { value: [], '@odata.nextLink': 'https://graph/loop' },
    });
    await expect(teamsMemberRecipients(graph, '/chats/c1/members')).rejects.toThrow('too long');
  });
});
