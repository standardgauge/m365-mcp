import React, { useState } from 'react';
import { render, screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, expect, it, vi } from 'vitest';
import KeyScopeFields, { isNarrowingOf, serviceNames } from '../components/KeyScopeFields';
import type { KeyPolicy, SitePick } from '../components/KeyScopeFields';

const SITE_A: SitePick = { id: 'site-a', name: 'Finance' };
const SITE_B: SitePick = { id: 'site-b', name: 'Legal' };
const SITE_C: SitePick = { id: 'site-c', name: 'Marketing' };

function policy(services: string[], readOnlyServices: string[] = [], allowedSites: SitePick[] = []): KeyPolicy {
  return { services, readOnlyServices, allowedSites };
}

/** KeyScopeFields is controlled; hold its value the way its parents do and expose the latest one. */
function Harness(props: Omit<React.ComponentProps<typeof KeyScopeFields>, 'value' | 'onChange'> & { initial: KeyPolicy; onValue?: (p: KeyPolicy) => void }) {
  const { initial, onValue, ...rest } = props;
  const [value, setValue] = useState(initial);
  return <KeyScopeFields {...rest} value={value} onChange={(next) => { setValue(next); onValue?.(next); }} />;
}

function serviceRow(name: string): HTMLElement {
  const row = screen.getByText(name, { selector: '.service-name' }).closest('.service-row');
  if (!(row instanceof HTMLElement)) throw new Error(`no row for ${name}`);
  return row;
}

describe('isNarrowingOf', () => {
  it.each([
    ['same policy', policy(['mail']), policy(['mail']), true],
    ['drops a service', policy(['mail']), policy(['mail', 'calendar']), true],
    ['adds a service', policy(['mail', 'calendar']), policy(['mail']), false],
    ['adds read-only', policy(['mail'], ['mail']), policy(['mail']), true],
    ['clears read-only', policy(['mail']), policy(['mail'], ['mail']), false],
    ['names sites where it had none', policy(['sharepoint'], [], [SITE_A]), policy(['sharepoint']), true],
    ['removes one of its sites', policy(['sharepoint'], [], [SITE_A]), policy(['sharepoint'], [], [SITE_A, SITE_B]), true],
    ['adds a site it did not name', policy(['sharepoint'], [], [SITE_A, SITE_C]), policy(['sharepoint'], [], [SITE_A, SITE_B]), false],
    ['removes every site, keeps SharePoint', policy(['sharepoint']), policy(['sharepoint'], [], [SITE_A]), false],
    ['removes every site and SharePoint', policy(['mail']), policy(['mail', 'sharepoint'], [], [SITE_A]), true],
  ])('%s', (_name, next, prev, expected) => {
    expect(isNarrowingOf(next, prev)).toBe(expected);
  });
});

describe('serviceNames', () => {
  it('labels known services in order and falls back to the raw name', () => {
    expect(serviceNames(['mail', 'sharepoint', 'unknown'])).toBe('Mail, SharePoint, unknown');
    expect(serviceNames([])).toBe('—');
  });
});

describe('KeyScopeFields', () => {
  it('offers only grantable services', () => {
    render(<Harness initial={policy([])} grantable={['mail', 'calendar']} tenantReadOnly={[]} />);
    expect(screen.getByRole('checkbox', { name: 'Mail' })).toBeEnabled();
    expect(screen.getByRole('checkbox', { name: 'Calendar' })).toBeEnabled();
    expect(screen.queryByText('Teams')).not.toBeInTheDocument();
    expect(screen.queryByText('SharePoint')).not.toBeInTheDocument();
  });

  it('locks a tenant read-only service as read-only and says why', async () => {
    const user = userEvent.setup();
    render(<Harness initial={policy([])} grantable={['mail']} tenantReadOnly={['mail']} />);
    await user.click(screen.getByRole('checkbox', { name: 'Mail' }));
    const ro = within(serviceRow('Mail')).getByRole('checkbox', { name: /read-only/ });
    expect(ro).toBeChecked();
    expect(ro).toBeDisabled();
    expect(ro).toHaveAccessibleName('read-only (set by admin for every key)');
  });

  it('clears the chosen sites when SharePoint is unticked', async () => {
    const user = userEvent.setup();
    const onValue = vi.fn();
    render(<Harness initial={policy(['sharepoint'])} grantable={['sharepoint']} tenantReadOnly={[]} siteOptions={[SITE_A, SITE_B]} onValue={onValue} />);
    expect(screen.getByText('SharePoint sites (leave empty for every site the tenant allows)')).toBeInTheDocument();
    await user.click(screen.getByRole('checkbox', { name: 'Finance' }));
    expect(onValue).toHaveBeenLastCalledWith(policy(['sharepoint'], [], [SITE_A]));
    await user.click(screen.getByRole('checkbox', { name: 'SharePoint' }));
    expect(onValue).toHaveBeenLastCalledWith(policy([], [], []));
    expect(screen.queryByRole('checkbox', { name: 'Finance' })).not.toBeInTheDocument();
  });

  describe('narrow-only', () => {
    const base = policy(['mail', 'sharepoint'], ['mail'], [SITE_A, SITE_B]);

    it('locks a service the key does not have, with the reason inline', () => {
      render(<Harness initial={base} grantable={['mail', 'calendar', 'sharepoint']} tenantReadOnly={[]} siteOptions={[SITE_A, SITE_B, SITE_C]} narrowOnlyFrom={base} />);
      const calendar = screen.getByRole('checkbox', { name: /^Calendar/ });
      expect(calendar).not.toBeChecked();
      expect(calendar).toBeDisabled();
      expect(within(serviceRow('Calendar')).getByText('Not on this key; adding it would widen the key, which is not allowed here.')).toBeInTheDocument();
    });

    it('locks a read-only flag the key already has, with the reason inline', () => {
      render(<Harness initial={base} grantable={['mail', 'sharepoint']} tenantReadOnly={[]} siteOptions={[SITE_A, SITE_B]} narrowOnlyFrom={base} />);
      const ro = within(serviceRow('Mail')).getByRole('checkbox', { name: /read-only/ });
      expect(ro).toBeChecked();
      expect(ro).toBeDisabled();
      expect(ro).toHaveAccessibleName('read-only (already read-only; clearing it would widen the key)');
      // SharePoint is not read-only on the key, so making it read-only (narrowing) stays open.
      expect(within(serviceRow('SharePoint')).getByRole('checkbox', { name: /read-only/ })).toBeEnabled();
    });

    it('offers only the sites the key names, and warns when every site is removed', async () => {
      const user = userEvent.setup();
      render(<Harness initial={base} grantable={['mail', 'sharepoint']} tenantReadOnly={[]} siteOptions={[SITE_A, SITE_B, SITE_C]} narrowOnlyFrom={base} />);
      expect(screen.getByText('SharePoint sites (the key names these sites; you can only remove)')).toBeInTheDocument();
      expect(screen.queryByRole('checkbox', { name: 'Marketing' })).not.toBeInTheDocument();
      await user.click(screen.getByRole('checkbox', { name: 'Finance' }));
      expect(screen.queryByText(/Removing every site/)).not.toBeInTheDocument();
      await user.click(screen.getByRole('checkbox', { name: 'Legal' }));
      expect(screen.getByText("Removing every site would fall back to the tenant's whole allow-list, which widens the key. Keep at least one, or drop SharePoint from the key.")).toBeInTheDocument();
      // A removed site can be put back; one the key never named cannot appear.
      expect(screen.getByRole('checkbox', { name: 'Finance' })).toBeEnabled();
    });

    it('lets a service be removed, and restores its read-only flag when it is put back', async () => {
      const user = userEvent.setup();
      const onValue = vi.fn();
      render(<Harness initial={base} grantable={['mail', 'sharepoint']} tenantReadOnly={[]} siteOptions={[SITE_A, SITE_B]} narrowOnlyFrom={base} onValue={onValue} />);
      await user.click(screen.getByRole('checkbox', { name: 'Mail' }));
      expect(onValue).toHaveBeenLastCalledWith(policy(['sharepoint'], [], [SITE_A, SITE_B]));
      // Mail is on the key, so putting it back is allowed, and it comes back read-only
      // as the key had it rather than writable behind a locked read-only box.
      await user.click(screen.getByRole('checkbox', { name: 'Mail' }));
      expect(onValue).toHaveBeenLastCalledWith(policy(['sharepoint', 'mail'], ['mail'], [SITE_A, SITE_B]));
      expect(within(serviceRow('Mail')).getByRole('checkbox', { name: /read-only/ })).toBeChecked();
    });
  });
});
