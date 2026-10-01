import React, { useEffect, useState } from 'react';

interface TaskListItem {
  id: string;
  name: string;
  isOwner: boolean;
  isShared: boolean;
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

export default function TasksBrowser({ userId, accessToken, onSelect }: Props) {
  const [items, setItems] = useState<TaskListItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!accessToken) return;
    setLoading(true);
    apiFetch('tasks/lists', userId, accessToken)
      .then((data) => setItems(data.lists ?? []))
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [userId, accessToken]);

  return (
    <div className="card">
      <h2>Tasks Browser</h2>
      {loading && <p style={{ color: '#666', fontSize: 13 }}>Loading…</p>}
      {error && <p className="error-msg">{error}</p>}
      {!loading && !error && items.length === 0 && (
        <p className="empty-state">No task lists found.</p>
      )}
      {!loading && items.length > 0 && (
        <div style={{ maxHeight: 400, overflowY: 'auto' }}>
          {items.map((item) => (
            <div key={item.id} className="tree-node">
              <span className="toggle" style={{ visibility: 'hidden' }}>▸</span>
              <span className="node-label">
                ✅ {item.name}
                {item.isShared && (
                  <span style={{ color: '#999', fontSize: 11, marginLeft: 6 }}>shared</span>
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
