import React, { useCallback, useEffect, useState } from 'react';

interface Site {
  id: string;
  displayName: string;
  webUrl: string;
}

interface Folder {
  id: string;
  name: string;
  path: string;
  childCount: number;
}

interface TreeNode {
  folder: Folder;
  siteId: string;
  children: TreeNode[] | null; // null = not yet loaded
  expanded: boolean;
  denied: boolean;
}

interface Props {
  userId: string;
  accessToken?: string;
  onSelect?: (path: string) => void;
  allowedSites?: Array<{ id: string; name: string }>;
  addedByName?: string;
  onDenyChange?: () => void;
}

async function apiFetch(path: string, userId: string, accessToken?: string) {
  const headers: Record<string, string> = { 'x-user-id': userId };
  if (accessToken) headers['Authorization'] = `Bearer ${accessToken}`;
  const res = await fetch(`/api/${path}`, { headers });
  if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
  return res.json();
}

export default function FolderBrowser({ userId, accessToken, onSelect, allowedSites, addedByName, onDenyChange }: Props) {
  const [sites, setSites] = useState<Site[]>([]);
  const [selectedSite, setSelectedSite] = useState<Site | null>(null);
  const [tree, setTree] = useState<TreeNode[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  // Load sites once both userId and accessToken are ready
  useEffect(() => {
    setLoading(true);
    apiFetch('sharepoint/sites', userId, accessToken)
      .then((data) => setSites(data.sites ?? []))
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [userId, accessToken]);

  const visibleSites =
    allowedSites !== undefined && allowedSites.length > 0
      ? sites.filter((s) => allowedSites.some((a) => a.id === s.id))
      : allowedSites !== undefined
        ? []
        : sites;

  const loadChildren = useCallback(
    async (siteId: string, parentId?: string): Promise<TreeNode[]> => {
      const qs = parentId
        ? `siteId=${siteId}&parentId=${parentId}`
        : `siteId=${siteId}`;
      const data = await apiFetch(`sharepoint/folders?${qs}`, userId, accessToken);
      return (data.folders ?? []).map((f: Folder) => ({
        folder: f,
        siteId,
        children: null,
        expanded: false,
        denied: false,
      }));
    },
    [userId, accessToken]
  );

  const selectSite = async (site: Site) => {
    setSelectedSite(site);
    setError(null);
    setLoading(true);
    try {
      const nodes = await loadChildren(site.id);
      setTree(nodes);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to load folders');
    } finally {
      setLoading(false);
    }
  };

  const toggleNode = async (path: number[]) => {
    const newTree = JSON.parse(JSON.stringify(tree)) as TreeNode[];
    let node = newTree[path[0]];
    for (let i = 1; i < path.length; i++) node = node.children![path[i]];

    if (!node.expanded && node.children === null) {
      // Load children on first expand
      try {
        node.children = await loadChildren(node.siteId, node.folder.id);
      } catch {
        node.children = [];
      }
    }
    node.expanded = !node.expanded;
    setTree(newTree);
  };

  const toggleDeny = async (node: TreeNode) => {
    const method = node.denied ? 'DELETE' : 'POST';
    await fetch('/api/manage/deny-list/global', {
      method,
      headers: { 'Content-Type': 'application/json', 'x-user-id': userId },
      body: JSON.stringify({ type: 'sharepoint', path: node.folder.path, ...(addedByName ? { addedByName } : {}) }),
    });
    onDenyChange?.();
    // Refetch tree from scratch to reflect updated deny status
    if (selectedSite) await selectSite(selectedSite);
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
      <h2>SharePoint Folder Browser</h2>

      <div style={{ marginBottom: 12 }}>
        <label style={{ fontSize: 12, display: 'block', marginBottom: 4 }}>Site</label>
        {allowedSites !== undefined && allowedSites.length === 0 ? (
          <p className="empty-state">No sites have been enabled by your administrator.</p>
        ) : (
          <select
            style={{ width: '100%', padding: '6px 8px', fontSize: 13, borderRadius: 4, border: '1px solid #ccc' }}
            value={selectedSite?.id ?? ''}
            onChange={(e) => {
              const site = visibleSites.find((s) => s.id === e.target.value);
              if (site) selectSite(site);
            }}
          >
            <option value="">— select a site —</option>
            {visibleSites.map((s) => (
              <option key={s.id} value={s.id}>
                {s.displayName}
              </option>
            ))}
          </select>
        )}
      </div>

      {loading && <p style={{ color: '#666', fontSize: 13 }}>Loading…</p>}
      {error && <p className="error-msg">{error}</p>}
      {!loading && !error && tree.length === 0 && selectedSite && (
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
