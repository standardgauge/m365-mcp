import React, { useCallback, useEffect, useState } from 'react';

interface OneDriveFolder {
  id: string;
  name: string;
  path: string;
  childCount: number;
}

interface TreeNode {
  folder: OneDriveFolder;
  children: TreeNode[] | null;
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

export default function OneDriveBrowser({ userId, accessToken, onSelect }: Props) {
  const [tree, setTree] = useState<TreeNode[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const loadFolders = useCallback(
    async (parentId?: string): Promise<TreeNode[]> => {
      const qs = parentId ? `parentId=${parentId}` : '';
      const data = await apiFetch(`onedrive/folders${qs ? `?${qs}` : ''}`, userId, accessToken);
      const items = data.items ?? data.folders ?? [];
      return items
        .filter((f: OneDriveFolder & { type?: string }) => f.type === 'folder' || f.childCount !== undefined)
        .map((f: OneDriveFolder) => ({
          folder: f,
          children: null,
          expanded: false,
        }));
    },
    [userId, accessToken]
  );

  useEffect(() => {
    setLoading(true);
    loadFolders()
      .then(setTree)
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [loadFolders]);

  const toggleNode = async (path: number[]) => {
    const newTree: TreeNode[] = JSON.parse(JSON.stringify(tree));
    let node = newTree[path[0]];
    for (let i = 1; i < path.length; i++) node = node.children![path[i]];

    if (!node.expanded && node.children === null && node.folder.childCount > 0) {
      try {
        node.children = await loadFolders(node.folder.id);
      } catch {
        node.children = [];
      }
    }
    node.expanded = !node.expanded;
    setTree(newTree);
  };

  const renderTree = (nodes: TreeNode[], pathPrefix: number[] = []) =>
    nodes.map((node, i) => {
      const nodePath = [...pathPrefix, i];
      return (
        <div key={node.folder.id}>
          <div className="tree-node">
            <span
              className="toggle"
              onClick={() => toggleNode(nodePath)}
              style={{ visibility: node.folder.childCount > 0 ? 'visible' : 'hidden' }}
            >
              {node.expanded ? '▾' : '▸'}
            </span>
            <span className="node-label" onClick={() => toggleNode(nodePath)}>
              📁 {node.folder.name}
            </span>
            {onSelect && (
              <button
                className="remove-btn"
                style={{ color: '#0078d4', borderColor: '#0078d4' }}
                onClick={() => onSelect(node.folder.id, node.folder.name)}
              >
                Block
              </button>
            )}
          </div>
          {node.expanded && node.children && (
            <div className="tree-children">
              {renderTree(node.children, nodePath)}
            </div>
          )}
        </div>
      );
    });

  return (
    <div className="card">
      <h2>OneDrive Browser</h2>
      {loading && <p style={{ color: '#666', fontSize: 13 }}>Loading…</p>}
      {error && <p className="error-msg">{error}</p>}
      {!loading && !error && tree.length === 0 && (
        <p className="empty-state">No folders found.</p>
      )}
      {!loading && tree.length > 0 && (
        <div style={{ maxHeight: 400, overflowY: 'auto' }}>
          {renderTree(tree)}
        </div>
      )}
    </div>
  );
}
