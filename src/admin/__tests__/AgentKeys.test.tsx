import React from 'react';
import { render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, expect, it } from 'vitest';
import AgentKeys from '../components/AgentKeys';
import type { KeyPolicy } from '../components/KeyScopeFields';
import { stubFetch } from './fetchStub';
import type { Call } from './fetchStub';

const SECRET = 'mcpk_live_0123456789abcdef';

interface TenantPolicy { allowKeys: boolean; disabledServices: string[]; readOnlyServices: string[]; maxKeyLifetimeDays: number; ownersMayWiden: boolean }

const TENANT: TenantPolicy = { allowKeys: true, disabledServices: [], readOnlyServices: [], maxKeyLifetimeDays: 365, ownersMayWiden: true };

function key(keyId: string, label: string, policy: KeyPolicy) {
  return {
    keyId, label, userId: 'u1', userEmail: 'owner@fabrikam.com', policy,
    createdAt: Date.UTC(2026, 0, 1), expiresAt: Date.UTC(2026, 3, 1), lastUsedAt: null, revokedAt: null, revokedBy: null,
  };
}

/** A fake server for the owner's key routes. `keys` is mutable so a create or revoke shows on the reload. */
function server(opts: { tenant?: Partial<TenantPolicy>; keys?: ReturnType<typeof key>[] } = {}): Call[] {
  const keys = [...(opts.keys ?? [])];
  return stubFetch(({ method, url, body }) => {
    if (url === '/api/manage/agent-key-policy' && method === 'GET') return { body: { ...TENANT, ...opts.tenant } };
    if (url === '/api/manage/agent-keys' && method === 'GET') return { body: { keys } };
    if (url === '/api/manage/agent-keys' && method === 'POST') {
      const b = body as { label: string; services: string[]; readOnlyServices: string[] };
      keys.push(key('k-new', b.label, { services: b.services, readOnlyServices: b.readOnlyServices, allowedSites: [] }));
      return { body: { key: SECRET, keyId: 'k-new', label: b.label, reminder: { eventId: 'evt-1' } } };
    }
    if (url.startsWith('/api/manage/agent-keys?keyId=') && method === 'DELETE') return { body: { revoked: true } };
    if (url.startsWith('/api/manage/agent-keys?keyId=') && method === 'PATCH') return { body: { widened: false } };
    return undefined;
  });
}

function renderKeys(grantable = ['mail', 'calendar', 'sharepoint'], allowedSites?: { id: string; name: string }[]) {
  return render(<AgentKeys userId="u1" grantableServices={grantable} allowedSites={allowedSites} />);
}

async function openCreateForm(user: ReturnType<typeof userEvent.setup>) {
  await user.click(await screen.findByRole('button', { name: 'New agent key' }));
}

describe('AgentKeys: minting', () => {
  it('shows the key value once, with the exact usage text, and never again after Done', async () => {
    const user = userEvent.setup();
    const calls = server();
    renderKeys();
    await openCreateForm(user);
    await user.type(screen.getByPlaceholderText('e.g. nightly-report'), 'nightly-report');
    await user.click(screen.getByRole('checkbox', { name: 'Mail' }));
    await user.click(screen.getByRole('button', { name: 'Create key' }));

    const dialog = await screen.findByRole('dialog', { name: 'New agent key' });
    expect(within(dialog).getByText(SECRET)).toBeInTheDocument();
    expect(within(dialog).getByText(/Key “nightly-report” created\./)).toBeInTheDocument();
    expect(within(dialog).getByText('A reminder is on your calendar a few days before it expires.')).toBeInTheDocument();

    // The helper text, as rendered: the escaped <key> placeholder must reach the reader as literal text.
    const helper = within(dialog).getByText((_, el) => el?.tagName === 'P' && el.textContent?.startsWith('Send it as') === true);
    expect(helper.textContent).toBe(
      'Send it as Authorization: Bearer <key> to /api/mcp, or hand it to the shim with --store-token like a session token. Keep it in a secret store, never in a repo.',
    );
    expect([...helper.querySelectorAll('code')].map((c) => c.textContent)).toEqual(['Authorization: Bearer <key>', '/api/mcp', '--store-token']);

    const post = calls.find((c) => c.method === 'POST');
    expect(post?.body).toEqual({ label: 'nightly-report', services: ['mail'], readOnlyServices: [], expiresInDays: 90, reminder: true });

    await user.click(within(dialog).getByRole('button', { name: 'Done' }));
    expect(screen.queryByRole('dialog', { name: 'New agent key' })).not.toBeInTheDocument();
    // The new key is listed, by label and id only.
    expect(await screen.findByText('k-new')).toBeInTheDocument();
    expect(document.body.textContent).not.toContain(SECRET);
  });

  it('blocks a lifetime outside 1 to the tenant maximum, or not a whole number of days', async () => {
    const user = userEvent.setup();
    server({ tenant: { maxKeyLifetimeDays: 30 } });
    renderKeys();
    await openCreateForm(user);
    await user.type(screen.getByPlaceholderText('e.g. nightly-report'), 'report');
    await user.click(screen.getByRole('checkbox', { name: 'Mail' }));
    const days = screen.getByRole('spinbutton');
    const create = screen.getByRole('button', { name: 'Create key' });
    expect(screen.getByText(/1 to 30\. The maximum is set by your admin\./)).toBeInTheDocument();

    for (const bad of ['31', '0', '1.5', '']) {
      await user.clear(days);
      if (bad) await user.type(days, bad);
      expect(create).toBeDisabled();
    }
    await user.clear(days);
    await user.type(days, '30');
    expect(create).toBeEnabled();
  });

  it('blocks a key with no service or no label', async () => {
    const user = userEvent.setup();
    server();
    renderKeys();
    await openCreateForm(user);
    const create = screen.getByRole('button', { name: 'Create key' });
    await user.type(screen.getByPlaceholderText('e.g. nightly-report'), 'report');
    expect(create).toBeDisabled();
    await user.click(screen.getByRole('checkbox', { name: 'Mail' }));
    expect(create).toBeEnabled();
    await user.clear(screen.getByPlaceholderText('e.g. nightly-report'));
    await user.type(screen.getByPlaceholderText('e.g. nightly-report'), '   ');
    expect(create).toBeDisabled();
  });

  it('does not offer a service the tenant keeps out of every key', async () => {
    const user = userEvent.setup();
    server({ tenant: { disabledServices: ['calendar'] } });
    renderKeys(['mail', 'calendar']);
    await openCreateForm(user);
    expect(screen.getByRole('checkbox', { name: 'Mail' })).toBeInTheDocument();
    expect(screen.queryByRole('checkbox', { name: /^Calendar(?! reminder)/ })).not.toBeInTheDocument();
  });

  it('sends picked sites only while SharePoint is on the key', async () => {
    const user = userEvent.setup();
    const calls = server();
    renderKeys(['mail', 'sharepoint'], [{ id: 's1', name: 'Finance' }]);
    await openCreateForm(user);
    await user.type(screen.getByPlaceholderText('e.g. nightly-report'), 'docs');
    await user.click(screen.getByRole('checkbox', { name: 'SharePoint' }));
    await user.click(screen.getByRole('checkbox', { name: 'Finance' }));
    await user.click(screen.getByRole('button', { name: 'Create key' }));
    await screen.findByRole('dialog', { name: 'New agent key' });
    expect(calls.find((c) => c.method === 'POST')?.body).toMatchObject({ services: ['sharepoint'], allowedSites: [{ id: 's1', name: 'Finance' }] });
  });

  it('offers no create form when the tenant has turned keys off', async () => {
    server({ tenant: { allowKeys: false } });
    renderKeys();
    expect(await screen.findByText(/Your administrator has turned agent keys off/)).toBeInTheDocument();
    expect(screen.queryByRole('button', { name: 'New agent key' })).not.toBeInTheDocument();
  });
});

describe('AgentKeys: editing', () => {
  const mailOnly = key('k1', 'reporter', { services: ['mail'], readOnlyServices: [], allowedSites: [] });

  it('warns and relabels the button only when the edit widens the key', async () => {
    const user = userEvent.setup();
    server({ keys: [mailOnly] });
    renderKeys();
    await user.click(await screen.findByRole('button', { name: 'Edit' }));
    expect(screen.getByRole('button', { name: 'Save' })).toBeInTheDocument();
    expect(screen.queryByText(/This change/)).not.toBeInTheDocument();

    // Narrowing: make Mail read-only. No warning.
    const mailRow = screen.getByText('Mail', { selector: '.service-name' }).closest('.service-row') as HTMLElement;
    await user.click(within(mailRow).getByRole('checkbox', { name: /read-only/ }));
    expect(screen.queryByText(/This change/)).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Save' })).toBeEnabled();

    // Widening: add Calendar.
    await user.click(screen.getByRole('checkbox', { name: 'Calendar' }));
    expect(screen.getByText((_, el) => el?.tagName === 'P' && el.textContent === 'This change widens the key. Whatever holds it gains the new access on its next call.')).toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Save and widen' })).toBeEnabled();

    // Undo the widening; the warning goes with it.
    await user.click(screen.getByRole('checkbox', { name: 'Calendar' }));
    expect(screen.queryByText(/This change/)).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Save' })).toBeEnabled();
  });

  it('locks widening controls with their reasons when the tenant is narrow-only', async () => {
    const user = userEvent.setup();
    const roMail = key('k2', 'reader', { services: ['mail'], readOnlyServices: ['mail'], allowedSites: [] });
    server({ tenant: { ownersMayWiden: false }, keys: [roMail] });
    renderKeys();
    await user.click(await screen.findByRole('button', { name: 'Edit' }));
    expect(screen.getByText((_, el) => el?.tagName === 'P' && el.textContent?.startsWith('Your organization allows a key to be narrowed only: you can remove services') === true)).toBeInTheDocument();
    const calendar = screen.getByRole('checkbox', { name: /^Calendar/ });
    expect(calendar).toBeDisabled();
    const calendarRow = screen.getByText('Calendar', { selector: '.service-name' }).closest('.service-row') as HTMLElement;
    expect(within(calendarRow).getByText('Not on this key; adding it would widen the key, which is not allowed here.')).toBeInTheDocument();
    expect(screen.getByRole('checkbox', { name: 'read-only (already read-only; clearing it would widen the key)' })).toBeDisabled();
    expect(screen.queryByText(/This change/)).not.toBeInTheDocument();
  });

  it('blocks saving a narrow-only edit that removes every site', async () => {
    const user = userEvent.setup();
    const sites = [{ id: 's1', name: 'Finance' }];
    const spKey = key('k3', 'docs', { services: ['sharepoint'], readOnlyServices: [], allowedSites: sites });
    server({ tenant: { ownersMayWiden: false }, keys: [spKey] });
    renderKeys(['sharepoint'], sites);
    await user.click(await screen.findByRole('button', { name: 'Edit' }));
    expect(screen.getByRole('button', { name: 'Save' })).toBeEnabled();
    await user.click(screen.getByRole('checkbox', { name: 'Finance' }));
    expect(screen.getByText(/Removing every site would fall back/)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /^Save/ })).toBeDisabled();
  });
});

describe('AgentKeys: revoking', () => {
  it('asks first, sends nothing on Cancel, and revokes on confirm', async () => {
    const user = userEvent.setup();
    const calls = server({ keys: [key('k1', 'reporter', { services: ['mail'], readOnlyServices: [], allowedSites: [] })] });
    renderKeys();
    await user.click(await screen.findByRole('button', { name: 'Revoke' }));
    let dialog = screen.getByRole('dialog', { name: 'Revoke agent key' });
    expect(dialog).toHaveTextContent('Revoke “reporter”? Anything using it stops working within about half a minute. Its calendar reminder is removed.');
    await user.click(within(dialog).getByRole('button', { name: 'Cancel' }));
    expect(screen.queryByRole('dialog', { name: 'Revoke agent key' })).not.toBeInTheDocument();
    expect(calls.some((c) => c.method === 'DELETE')).toBe(false);

    await user.click(screen.getByRole('button', { name: 'Revoke' }));
    dialog = screen.getByRole('dialog', { name: 'Revoke agent key' });
    await user.click(within(dialog).getByRole('button', { name: 'Revoke' }));
    await waitFor(() => expect(calls.filter((c) => c.method === 'DELETE').map((c) => c.url)).toEqual(['/api/manage/agent-keys?keyId=k1']));
    expect(await screen.findByText('Revoked “reporter”.')).toBeInTheDocument();
  });
});
