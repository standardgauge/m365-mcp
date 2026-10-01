import React, { useEffect, useState } from 'react';

interface ContactFolderItem {
  id: string;
  name: string;
  parentFolderId: string | null;
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

export default function ContactsBrowser({ userId, accessToken, onSelect }: Props) {
  const [items, setItems] = useState<ContactFolderItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setLoading(true);
    apiFetch('contacts/folders', userId, accessToken)
      .then((data) => setItems(data.folders ?? []))
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [userId, accessToken]);

  return (
    <div className="card">
      <h2>Contacts Browser</h2>
      {loading && <p style={{ color: '#666', fontSize: 13 }}>Loading…</p>}
      {error && <p className="error-msg">{error}</p>}
      {!loading && !error && items.length === 0 && (
        <p className="empty-state">No contact folders found.</p>
      )}
      {!loading && items.length > 0 && (
        <div style={{ maxHeight: 400, overflowY: 'auto' }}>
          {items.map((item) => (
            <div key={item.id} className="tree-node">
              <span className="toggle" style={{ visibility: 'hidden' }}>▸</span>
              <span className="node-label">
                👥 {item.name}
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
