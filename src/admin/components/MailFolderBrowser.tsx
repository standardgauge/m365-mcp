import React, { useCallback, useEffect, useState } from 'react';

interface MailFolder {
  id: string;
  name: string;
  path: string;
  totalItemCount: number;
  childFolderCount: number;
  wellKnownName: string | null;
}

interface MailTreeNode {
  folder: MailFolder;
  mailboxId: string;
  children: MailTreeNode[] | null;
  expanded: boolean;
  denied: boolean;
}

interface Props {
  userId: string;
  accessToken?: string;
  onSelect?: (path: string) => void;
}

async function apiFetch(path: string, userId: string, accessToken?: string) {
  const headers: Record<string, string> = { 'x-user-id': userId };
  if (accessToken) headers['Authorization'] = `Bearer ${accessToken}`;
  const res = await fetch(`/api/${path}`, { headers });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  return res.json();
}

// Well-known folder icons
const FOLDER_ICONS: Record<string, string> = {
  inbox: '📥',
  sentItems: '📤',
  drafts: '📝',
  deletedItems: '🗑️',
  junkEmail: '🚫',
  archive: '📦',
};

function folderIcon(wellKnownName: string | null, name: string): string {
  if (wellKnownName && FOLDER_ICONS[wellKnownName]) return FOLDER_ICONS[wellKnownName];
  return '📁';
}

export default function MailFolderBrowser({ userId, accessToken, onSelect }: Props) {
  const [tree, setTree] = useState<MailTreeNode[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const loadFolders = useCallback(
    async (mailboxId = 'me', parentFolderId?: string): Promise<MailTreeNode[]> => {
      const qs = parentFolderId
        ? `mailboxId=${mailboxId}&parentFolderId=${parentFolderId}`
        : `mailboxId=${mailboxId}`;
      const data = await apiFetch(`mail/folders?${qs}`, userId, accessToken);
      return (data.folders ?? []).map((f: MailFolder) => ({
        folder: f,
        mailboxId,
        children: null,
        expanded: false,
        denied: false,
      }));
    },
    [userId, accessToken]
  );

  // Load top-level folders once accessToken is ready
  useEffect(() => {
    setLoading(true);
    loadFolders()
      .then(setTree)
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [loadFolders]);

  const toggleNode = async (path: number[]) => {
    const newTree: MailTreeNode[] = JSON.parse(JSON.stringify(tree));
    let node = newTree[path[0]];
    for (let i = 1; i < path.length; i++) node = node.children![path[i]];

    if (!node.expanded && node.children === null && node.folder.childFolderCount > 0) {
      try {
        node.children = await loadFolders(node.mailboxId, node.folder.id);
      } catch {
        node.children = [];
      }
    }
    node.expanded = !node.expanded;
    setTree(newTree);
  };

  const toggleDeny = async (node: MailTreeNode) => {
    const method = node.denied ? 'DELETE' : 'POST';
    await fetch('/api/manage/deny-list/global', {
      method,
      headers: { 'Content-Type': 'application/json', 'x-user-id': userId },
      body: JSON.stringify({ type: 'mail', path: node.folder.path }),
    });
    // Reload from the top to reflect updated state
    setLoading(true);
    loadFolders()
      .then(setTree)
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  };

  const renderTree = (nodes: MailTreeNode[], pathPrefix: number[] = []) =>
    nodes.map((node, i) => {
      const nodePath = [...pathPrefix, i];
      const icon = folderIcon(node.folder.wellKnownName, node.folder.name);
      return (
        <div key={node.folder.id}>
          <div className="tree-node">
            <span
              className="toggle"
              onClick={() => toggleNode(nodePath)}
              style={{ visibility: node.folder.childFolderCount > 0 ? 'visible' : 'hidden' }}
            >
              {node.expanded ? '▾' : '▸'}
            </span>
            <span className="node-label" onClick={() => toggleNode(nodePath)}>
              {icon} {node.folder.name}
              <span style={{ color: '#999', fontSize: 11, marginLeft: 6 }}>
                ({node.folder.totalItemCount})
              </span>
            </span>
            {node.denied && <span className="denied-badge">blocked</span>}
            {onSelect && (
              <button
                className="remove-btn"
                style={{ color: '#0078d4', borderColor: '#0078d4' }}
                onClick={() => onSelect(node.folder.path)}
              >
                Block
              </button>
            )}
            {!onSelect && (
              <button
                className="remove-btn"
                style={node.denied ? {} : { color: '#0078d4', borderColor: '#0078d4' }}
                onClick={() => toggleDeny(node)}
              >
                {node.denied ? 'Unblock' : 'Block'}
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
      <h2>Mail Folder Browser</h2>
      {loading && <p style={{ color: '#666', fontSize: 13 }}>Loading…</p>}
      {error && <p className="error-msg">{error}</p>}
      {!loading && !error && tree.length === 0 && (
        <p className="empty-state">No mail folders found.</p>
      )}
      {!loading && tree.length > 0 && (
        <div style={{ maxHeight: 400, overflowY: 'auto' }}>
          {renderTree(tree)}
        </div>
      )}
    </div>
  );
}
