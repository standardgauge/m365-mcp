import React, { useState, useEffect } from 'react';

interface Props {
  /** Unique key for localStorage persistence */
  storageKey: string;
  /** Section heading text */
  title: string;
  /** Optional count badge displayed next to the title */
  count?: number;
  /** Default collapsed if entries exist, expanded if empty (when not overridden by localStorage) */
  defaultCollapsed?: boolean;
  children: React.ReactNode;
}

export default function CollapsibleSection({ storageKey, title, count, defaultCollapsed, children }: Props) {
  const lsKey = `collapsible_${storageKey}`;

  const [collapsed, setCollapsed] = useState(() => {
    const stored = localStorage.getItem(lsKey);
    if (stored !== null) return stored === '1';
    return defaultCollapsed ?? false;
  });

  useEffect(() => {
    localStorage.setItem(lsKey, collapsed ? '1' : '0');
  }, [collapsed, lsKey]);

  return (
    <div className="card">
      <h2
        className="collapsible-header"
        onClick={() => setCollapsed((c) => !c)}
      >
        <span className="collapsible-arrow">{collapsed ? '\u25b8' : '\u25be'}</span>
        {title}
        {count !== undefined && <span className="badge-count">{count}</span>}
      </h2>
      {!collapsed && children}
    </div>
  );
}
