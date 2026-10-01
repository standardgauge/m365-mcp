import React, { useCallback, useEffect, useState } from 'react';

interface NotebookItem {
  id: string;
  name: string;
  createdDateTime: string;
}

interface SectionItem {
  id: string;
  displayName: string;
  pagesUrl: string;
}

interface NotebookNode {
  notebook: NotebookItem;
  sections: SectionItem[] | null;
  sectionError: string | null;
  expanded: boolean;
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

export default function OneNoteBrowser({ userId, accessToken, onSelect }: Props) {
  const [nodes, setNodes] = useState<NotebookNode[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setLoading(true);
    apiFetch('onenote/notebooks', userId, accessToken)
      .then((data) =>
        setNodes(
          (data.notebooks ?? []).map((n: NotebookItem) => ({
            notebook: n,
            sections: null,
            sectionError: null,
            expanded: false,
          }))
        )
      )
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [userId, accessToken]);

  const fetchSections = useCallback(
    async (notebookId: string): Promise<SectionItem[]> => {
      const headers: Record<string, string> = { 'x-user-id': userId };
      if (accessToken) headers['Authorization'] = `Bearer ${accessToken}`;
      const res = await fetch(`/api/onenote/notebooks/${notebookId}/sections`, { headers });
      if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
      const data = await res.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (data.sections ?? []).map((s: any) => ({
        id: s.id,
        displayName: s.displayName,
        pagesUrl: s.pagesUrl,
      }));
    },
    [userId, accessToken]
  );

  const toggleNode = async (index: number) => {
    const newNodes = [...nodes];
    const node = { ...newNodes[index] };

    if (!node.expanded && node.sections === null) {
      try {
        node.sections = await fetchSections(node.notebook.id);
        node.sectionError = null;
      } catch (e: unknown) {
        node.sections = [];
        node.sectionError = e instanceof Error ? e.message : 'Failed to load sections';
      }
    }
    node.expanded = !node.expanded;
    newNodes[index] = node;
    setNodes(newNodes);
  };

  return (
    <div className="card">
      <h2>OneNote Browser</h2>
      {loading && <p style={{ color: '#666', fontSize: 13 }}>Loading…</p>}
      {error && <p className="error-msg">{error}</p>}
      {!loading && !error && nodes.length === 0 && (
        <p className="empty-state">No notebooks found.</p>
      )}
      {!loading && nodes.length > 0 && (
        <div style={{ maxHeight: 400, overflowY: 'auto' }}>
          {nodes.map((node, i) => (
            <div key={node.notebook.id}>
              <div className="tree-node">
                <span
                  className="toggle"
                  onClick={() => toggleNode(i)}
                >
                  {node.expanded ? '▾' : '▸'}
                </span>
                <span className="node-label" onClick={() => toggleNode(i)}>
                  📓 {node.notebook.name}
                </span>
                {onSelect && (
                  <button
                    className="remove-btn"
                    style={{ color: '#0078d4', borderColor: '#0078d4' }}
                    onClick={() => onSelect(node.notebook.id, node.notebook.name)}
                  >
                    Block
                  </button>
                )}
              </div>
              {node.expanded && node.sectionError && (
                <div className="tree-children">
                  <p className="error-msg" style={{ margin: '4px 0' }}>{node.sectionError}</p>
                </div>
              )}
              {node.expanded && !node.sectionError && node.sections && node.sections.length > 0 && (
                <div className="tree-children">
                  {node.sections.map((sec) => (
                    <div key={sec.id} className="tree-node">
                      <span className="toggle" style={{ visibility: 'hidden' }}>▸</span>
                      <span className="node-label">
                        📄 {sec.displayName}
                      </span>
                      {onSelect && (
                        <button
                          className="remove-btn"
                          style={{ color: '#0078d4', borderColor: '#0078d4' }}
                          onClick={() => onSelect(sec.id, `${node.notebook.name} / ${sec.displayName}`)}
                        >
                          Block
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
