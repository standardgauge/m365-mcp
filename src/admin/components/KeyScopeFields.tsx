import React from 'react';

// A Map, not an object: service names arrive from the API, and a name like
// "constructor" must not resolve through the prototype chain.
const SERVICE_LABELS = new Map<string, string>([
  ['mail', 'Mail'], ['sharepoint', 'SharePoint'], ['calendar', 'Calendar'], ['onedrive', 'OneDrive'],
  ['onenote', 'OneNote'], ['contacts', 'Contacts'], ['teams', 'Teams'],
]);
export function serviceLabel(s: string): string {
  return SERVICE_LABELS.get(s) ?? s;
}
export const SERVICE_ORDER = ['mail', 'sharepoint', 'calendar', 'onedrive', 'onenote', 'contacts', 'teams'];

export interface SitePick { id: string; name: string }
export interface KeyPolicy { services: string[]; readOnlyServices: string[]; allowedSites: SitePick[] }

export function serviceNames(list: string[]): string {
  return list.length ? list.map(serviceLabel).join(', ') : '—';
}

/** True when `next` grants no more than `prev`. Mirrors isNarrowingOf on the server. */
export function isNarrowingOf(next: KeyPolicy, prev: KeyPolicy): boolean {
  if (next.services.some((s) => !prev.services.includes(s))) return false;
  for (const s of next.services) {
    if (prev.readOnlyServices.includes(s) && !next.readOnlyServices.includes(s)) return false;
  }
  if (prev.allowedSites.length === 0) return true;
  if (next.allowedSites.length === 0) return !next.services.includes('sharepoint');
  const prevIds = new Set(prev.allowedSites.map((s) => s.id));
  return next.allowedSites.every((s) => prevIds.has(s.id));
}

interface Props {
  value: KeyPolicy;
  onChange: (next: KeyPolicy) => void;
  /** Services the editor may offer at all. */
  grantable: string[];
  /** Services the tenant forces read-only for every key; shown checked and locked. */
  tenantReadOnly: string[];
  /** Site options to pick from, when the tenant has an allow-list. */
  siteOptions?: SitePick[];
  /** When set, the editor refuses anything wider than this policy and says why. */
  narrowOnlyFrom?: KeyPolicy;
  disabled?: boolean;
}

/**
 * The services / read-only / sites part of a key's policy. Used by the owner's
 * create and edit forms and by the administrator's edit form. In narrow-only
 * mode a service the key does not already have cannot be added, a read-only
 * flag it already has cannot be cleared, and a site it does not already name
 * cannot be picked. Each locked control says so inline.
 */
export default function KeyScopeFields({ value, onChange, grantable, tenantReadOnly, siteOptions, narrowOnlyFrom, disabled }: Props) {
  const toggle = (list: string[], v: string, on: boolean) => (on ? [...new Set([...list, v])] : list.filter((x) => x !== v));
  const baseHasSites = (narrowOnlyFrom?.allowedSites.length ?? 0) > 0;
  const baseSiteIds = new Set((narrowOnlyFrom?.allowedSites ?? []).map((s) => s.id));

  return (
    <>
      <div className="service-name" style={{ marginBottom: 4 }}>Services the key may use</div>
      <div className="services-list">
        {SERVICE_ORDER.filter((s) => grantable.includes(s) || value.services.includes(s)).map((s) => {
          const on = value.services.includes(s);
          const canAdd = !narrowOnlyFrom || narrowOnlyFrom.services.includes(s);
          const roForced = tenantReadOnly.includes(s);
          const roLockedByBase = Boolean(narrowOnlyFrom?.readOnlyServices.includes(s));
          const ro = on && (roForced || value.readOnlyServices.includes(s));
          return (
            <div key={s} className={`service-row${!grantable.includes(s) ? ' service-row--tenant-disabled' : ''}`}>
              <label className="service-toggle">
                <input type="checkbox" checked={on} disabled={disabled || (!on && !canAdd) || (!on && !grantable.includes(s))}
                  onChange={(e) => {
                    const services = toggle(value.services, s, e.target.checked);
                    // Re-adding a service the key holds read-only brings the flag back with it; its
                    // read-only box is locked in narrow-only mode, so it could not be set afterwards.
                    const readOnlyServices = !e.target.checked
                      ? value.readOnlyServices.filter((x) => x !== s)
                      : roLockedByBase ? toggle(value.readOnlyServices, s, true) : value.readOnlyServices;
                    const allowedSites = s === 'sharepoint' && !e.target.checked ? [] : value.allowedSites;
                    onChange({ services, readOnlyServices, allowedSites });
                  }} />
                <span className="service-info">
                  <span className="service-name">{serviceLabel(s)}</span>
                  {!on && !canAdd && <span className="service-desc">Not on this key; adding it would widen the key, which is not allowed here.</span>}
                  {!on && canAdd && !grantable.includes(s) && <span className="service-desc">Not available to you, so not grantable.</span>}
                </span>
              </label>
              {on && (
                <label className="service-toggle" style={{ marginLeft: 'auto' }}>
                  <input type="checkbox" checked={ro} disabled={disabled || roForced || roLockedByBase}
                    onChange={(e) => onChange({ ...value, readOnlyServices: toggle(value.readOnlyServices, s, e.target.checked) })} />
                  <span className="service-desc">
                    read-only{roForced ? ' (set by admin for every key)' : roLockedByBase ? ' (already read-only; clearing it would widen the key)' : ''}
                  </span>
                </label>
              )}
            </div>
          );
        })}
      </div>

      {value.services.includes('sharepoint') && siteOptions && siteOptions.length > 0 && (
        <>
          <div className="service-name" style={{ margin: '10px 0 4px' }}>
            SharePoint sites {baseHasSites ? '(the key names these sites; you can only remove)' : '(leave empty for every site the tenant allows)'}
          </div>
          <div className="services-list">
            {siteOptions.filter((site) => !baseHasSites || baseSiteIds.has(site.id) || value.allowedSites.some((v) => v.id === site.id)).map((site) => {
              const picked = value.allowedSites.some((v) => v.id === site.id);
              const canPick = !baseHasSites || baseSiteIds.has(site.id);
              return (
                <div key={site.id} className="service-row">
                  <label className="service-toggle">
                    <input type="checkbox" checked={picked} disabled={disabled || (!picked && !canPick)}
                      onChange={(e) => onChange({
                        ...value,
                        allowedSites: e.target.checked
                          ? [...value.allowedSites, { id: site.id, name: site.name }]
                          : value.allowedSites.filter((v) => v.id !== site.id),
                      })} />
                    <span className="service-info"><span className="service-name">{site.name}</span></span>
                  </label>
                </div>
              );
            })}
          </div>
          {baseHasSites && value.allowedSites.length === 0 && (
            <p className="error-msg">Removing every site would fall back to the tenant's whole allow-list, which widens the key. Keep at least one, or drop SharePoint from the key.</p>
          )}
        </>
      )}
    </>
  );
}
