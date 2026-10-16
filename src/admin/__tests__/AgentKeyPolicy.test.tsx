import React from 'react';
import { render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, expect, it } from 'vitest';
import AgentKeyPolicy from '../components/AgentKeyPolicy';
import type { KeyPolicy, SitePick } from '../components/KeyScopeFields';
import { stubFetch } from './fetchStub';
import type { Call } from './fetchStub';

interface TenantPolicy { allowKeys: boolean; disabledServices: string[]; readOnlyServices: string[]; maxKeyLifetimeDays: number; ownersMayWiden: boolean }

const TENANT: TenantPolicy = { allowKeys: true, disabledServices: [], readOnlyServices: [], maxKeyLifetimeDays: 90, ownersMayWiden: true };
const ALL_SERVICES = ['mail', 'sharepoint', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'];
const FINANCE: SitePick = { id: 's1', name: 'Finance' };
const LEGAL: SitePick = { id: 's2', name: 'Legal' };

function key(keyId: string, label: string, userEmail: string, policy: KeyPolicy) {
  return {
    keyId, label, userId: `id-${userEmail}`, userEmail, policy,
    createdAt: Date.UTC(2026, 0, 1), expiresAt: Date.UTC(2026, 3, 1), lastUsedAt: null, revokedAt: null, revokedBy: null,
  };
}

const OTHERS_KEY = key('k1', 'their-agent', 'someone@fabrikam.com', { services: ['mail', 'sharepoint'], readOnlyServices: ['mail'], allowedSites: [FINANCE, LEGAL] });

function server(opts: { tenant?: Partial<TenantPolicy>; keys?: ReturnType<typeof key>[] } = {}): Call[] {
  let tenant: TenantPolicy = { ...TENANT, ...opts.tenant };
  return stubFetch(({ method, url, body }) => {
    if (url === '/api/manage/agent-key-policy' && method === 'GET') return { body: tenant };
    if (url === '/api/manage/agent-key-policy' && method === 'POST') {
      tenant = { ...tenant, ...(body as Partial<TenantPolicy>) };
      return { body: tenant };
    }
    if (url === '/api/manage/agent-keys?all=true') return { body: { keys: opts.keys ?? [] } };
    if (url.startsWith('/api/manage/agent-keys?keyId=') && (method === 'PATCH' || method === 'DELETE')) return { body: {} };
    return undefined;
  });
}

/** The admin's narrow form, open for the key named `label`. */
function narrowForm(label: string): HTMLElement {
  return screen.getByText(new RegExp(`^Narrow “${label}”`)).closest('form') as HTMLElement;
}

function renderPolicy(allowedSites?: SitePick[]) {
  return render(<AgentKeyPolicy userId="admin" enabledServices={ALL_SERVICES} allowedSites={allowedSites} />);
}

describe('AgentKeyPolicy: tenant settings', () => {
  it('blocks a key lifetime cap outside 1 to 365, fractional, or unchanged', async () => {
    const user = userEvent.setup();
    const calls = server();
    renderPolicy();
    const days = await screen.findByRole('spinbutton');
    const save = within(days.closest('p') as HTMLElement).getByRole('button', { name: 'Save' });
    expect(days).toHaveValue(90);
    expect(save).toBeDisabled();

    for (const bad of ['0', '366', '2.5', '']) {
      await user.clear(days);
      if (bad) await user.type(days, bad);
      expect(save).toBeDisabled();
    }
    await user.clear(days);
    await user.type(days, '30');
    expect(save).toBeEnabled();
    await user.click(save);
    await waitFor(() => expect(calls.find((c) => c.method === 'POST')?.body).toEqual({ maxKeyLifetimeDays: 30 }));
  });

  it('switches owners to narrow-only', async () => {
    const user = userEvent.setup();
    const calls = server();
    renderPolicy();
    await user.click(await screen.findByRole('checkbox', { name: /^Owners may widen their own keys/ }));
    await waitFor(() => expect(calls.find((c) => c.method === 'POST')?.body).toEqual({ ownersMayWiden: false }));
    expect(await screen.findByText('Narrow-only')).toBeInTheDocument();
  });
});

describe("AgentKeyPolicy: another user's key", () => {
  it('offers only narrowing: no service the key lacks, and its read-only flags locked', async () => {
    const user = userEvent.setup();
    server({ keys: [OTHERS_KEY] });
    renderPolicy([FINANCE, LEGAL, { id: 's3', name: 'Marketing' }]);
    await user.click(await screen.findByRole('button', { name: 'Narrow' }));

    const form = narrowForm('their-agent');
    expect(within(form).getByText('someone@fabrikam.com', { exact: false })).toBeInTheDocument();
    const offered = within(form).getAllByText(/.+/, { selector: '.services-list .service-name' }).map((el) => el.textContent);
    // Only the key's own services and sites appear; nothing the key lacks can be ticked on.
    expect(offered).toEqual(['Mail', 'SharePoint', 'Finance', 'Legal']);
    expect(within(form).getByRole('checkbox', { name: 'read-only (already read-only; clearing it would widen the key)' })).toBeDisabled();
  });

  it('sends the narrowed policy', async () => {
    const user = userEvent.setup();
    const calls = server({ keys: [OTHERS_KEY] });
    renderPolicy();
    await user.click(await screen.findByRole('button', { name: 'Narrow' }));
    await user.click(screen.getByRole('checkbox', { name: 'Legal' }));
    await user.click(within(narrowForm('their-agent')).getByRole('button', { name: 'Save' }));
    await waitFor(() => expect(calls.find((c) => c.method === 'PATCH')).toEqual({
      method: 'PATCH',
      url: '/api/manage/agent-keys?keyId=k1',
      body: { services: ['mail', 'sharepoint'], readOnlyServices: ['mail'], allowedSites: [FINANCE] },
    }));
    expect(await screen.findByText("Narrowed “their-agent” (someone@fabrikam.com). The owner's app has the new limits from its next call.")).toBeInTheDocument();
  });

  it('blocks saving when removing every site would widen the key', async () => {
    const user = userEvent.setup();
    const calls = server({ keys: [OTHERS_KEY] });
    renderPolicy();
    await user.click(await screen.findByRole('button', { name: 'Narrow' }));
    await user.click(screen.getByRole('checkbox', { name: 'Finance' }));
    await user.click(screen.getByRole('checkbox', { name: 'Legal' }));
    expect(screen.getByText(/Removing every site would fall back/)).toBeInTheDocument();
    const save = within(narrowForm('their-agent')).getByRole('button', { name: 'Save' });
    expect(save).toBeDisabled();
    // Dropping SharePoint as well is a narrowing again.
    await user.click(screen.getByRole('checkbox', { name: 'SharePoint' }));
    expect(save).toBeEnabled();
    expect(calls.some((c) => c.method === 'PATCH')).toBe(false);
  });

  it('asks before revoking, naming the owner, and revokes on confirm', async () => {
    const user = userEvent.setup();
    const calls = server({ keys: [OTHERS_KEY] });
    renderPolicy();
    await user.click(await screen.findByRole('button', { name: 'Revoke' }));
    let dialog = screen.getByRole('dialog', { name: 'Revoke agent key' });
    expect(dialog).toHaveTextContent('Revoke “their-agent” belonging to someone@fabrikam.com? Anything using it stops working within about half a minute. Their calendar reminder for it is removed.');
    await user.click(within(dialog).getByRole('button', { name: 'Cancel' }));
    expect(calls.some((c) => c.method === 'DELETE')).toBe(false);

    await user.click(screen.getByRole('button', { name: 'Revoke' }));
    dialog = screen.getByRole('dialog', { name: 'Revoke agent key' });
    await user.click(within(dialog).getByRole('button', { name: 'Revoke' }));
    await waitFor(() => expect(calls.filter((c) => c.method === 'DELETE').map((c) => c.url)).toEqual(['/api/manage/agent-keys?keyId=k1']));
    expect(await screen.findByText('Revoked “their-agent” (someone@fabrikam.com).')).toBeInTheDocument();
  });
});
