import React, { useEffect, useState } from 'react';

interface CalendarItem {
  id: string;
  name: string;
  color: string;
  isDefaultCalendar: boolean;
}

interface Props {
  userId: string;
  accessToken?: string;
  onSelect?: (id: string, name: string) => void;
}

async function apiFetch(path: string, userId: string, accessToken?: string) {
  const headers: Record<string, string> = { 'x-user-id': userId };
  if (accessToken) headers['Authorization'] = `Bearer ${accessToken}`;
  const res = await fetch(`/api/${path}`, { headers });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  return res.json();
}

const COLOR_SWATCHES: Record<string, string> = {
  auto: '#0078d4',
  lightBlue: '#56a0d3',
  lightGreen: '#56c56e',
  lightOrange: '#e67e22',
  lightGray: '#999',
  lightYellow: '#f4c542',
  lightTeal: '#17a589',
  lightPink: '#e91e8c',
  lightBrown: '#8d6e63',
  lightRed: '#e74c3c',
  maxBlue: '#0051a2',
  maxGreen: '#1a7a45',
  maxOrange: '#c0392b',
  maxGray: '#555',
  maxYellow: '#b8860b',
  maxTeal: '#0e7a6c',
  maxPink: '#c2185b',
  maxBrown: '#5d4037',
  maxRed: '#a93226',
};

export default function CalendarBrowser({ userId, accessToken, onSelect }: Props) {
  const [items, setItems] = useState<CalendarItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setLoading(true);
    apiFetch('calendar/calendars', userId, accessToken)
      .then((data) => setItems(data.calendars ?? []))
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [userId, accessToken]);

  return (
    <div className="card">
      <h2>Calendar Browser</h2>
      {loading && <p style={{ color: '#666', fontSize: 13 }}>Loading…</p>}
      {error && <p className="error-msg">{error}</p>}
      {!loading && !error && items.length === 0 && (
        <p className="empty-state">No calendars found.</p>
      )}
      {!loading && items.length > 0 && (
        <div style={{ maxHeight: 400, overflowY: 'auto' }}>
          {items.map((item) => (
            <div key={item.id} className="tree-node">
              <span className="toggle" style={{ visibility: 'hidden' }}>▸</span>
              <span className="node-label">
                <span
                  style={{
                    display: 'inline-block',
                    width: 10,
                    height: 10,
                    borderRadius: '50%',
                    background: COLOR_SWATCHES[item.color] ?? '#0078d4',
                    marginRight: 6,
                    verticalAlign: 'middle',
                  }}
                />
                {item.name}
                {item.isDefaultCalendar && (
                  <span style={{ color: '#999', fontSize: 11, marginLeft: 6 }}>(default)</span>
                )}
              </span>
              {onSelect && (
                <button
                  className="remove-btn"
                  style={{ color: '#0078d4', borderColor: '#0078d4' }}
                  onClick={() => onSelect(item.id, item.name)}
                >
                  Block
                </button>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
