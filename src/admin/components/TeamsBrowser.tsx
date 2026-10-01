import React, { useCallback, useEffect, useState } from 'react';

interface TeamItem {
  id: string;
  name: string;
  description: string | null;
}

interface ChannelItem {
  id: string;
  displayName: string;
  membershipType: string;
}

interface TeamNode {
  team: TeamItem;
  channels: ChannelItem[] | null;
  channelError: string | null;
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

export default function TeamsBrowser({ userId, accessToken, onSelect }: Props) {
  const [nodes, setNodes] = useState<TeamNode[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    setLoading(true);
    apiFetch('teams/teams', userId, accessToken)
      .then((data) =>
        setNodes(
          (data.teams ?? []).map((t: TeamItem) => ({
            team: t,
            channels: null,
            channelError: null,
            expanded: false,
          }))
        )
      )
      .catch((e) => setError(e.message))
      .finally(() => setLoading(false));
  }, [userId, accessToken]);

  const fetchChannels = useCallback(
    async (teamId: string): Promise<ChannelItem[]> => {
      const headers: Record<string, string> = { 'x-user-id': userId };
      if (accessToken) headers['Authorization'] = `Bearer ${accessToken}`;
      const res = await fetch(`/api/teams/teams/${teamId}/channels`, { headers });
      if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
      const data = await res.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      return (data.channels ?? []).map((c: any) => ({
        id: c.id,
        displayName: c.displayName,
        membershipType: c.membershipType ?? 'standard',
      }));
    },
    [userId, accessToken]
  );

  const toggleNode = async (index: number) => {
    const newNodes = [...nodes];
    const node = { ...newNodes[index] };

    if (!node.expanded && node.channels === null) {
      try {
        node.channels = await fetchChannels(node.team.id);
        node.channelError = null;
      } catch (e: unknown) {
        node.channels = [];
        node.channelError = e instanceof Error ? e.message : 'Failed to load channels';
      }
    }
    node.expanded = !node.expanded;
    newNodes[index] = node;
    setNodes(newNodes);
  };

  return (
    <div className="card">
      <h2>Teams Browser</h2>
      {loading && <p style={{ color: '#666', fontSize: 13 }}>Loading…</p>}
      {error && <p className="error-msg">{error}</p>}
      {!loading && !error && nodes.length === 0 && (
        <p className="empty-state">No teams found.</p>
      )}
      {!loading && nodes.length > 0 && (
        <div style={{ maxHeight: 400, overflowY: 'auto' }}>
          {nodes.map((node, i) => (
            <div key={node.team.id}>
              <div className="tree-node">
                <span
                  className="toggle"
                  onClick={() => toggleNode(i)}
                >
                  {node.expanded ? '▾' : '▸'}
                </span>
                <span className="node-label" onClick={() => toggleNode(i)}>
                  👥 {node.team.name}
                  {node.team.description && (
                    <span style={{ color: '#999', fontSize: 11, marginLeft: 6 }}>
                      {node.team.description}
                    </span>
                  )}
                </span>
                {onSelect && (
                  <button
                    className="remove-btn"
                    style={{ color: '#0078d4', borderColor: '#0078d4' }}
                    onClick={() => onSelect(node.team.id, node.team.name)}
                  >
                    Block
                  </button>
                )}
              </div>
              {node.expanded && node.channelError && (
                <div className="tree-children">
                  <p className="error-msg" style={{ margin: '4px 0' }}>{node.channelError}</p>
                </div>
              )}
              {node.expanded && !node.channelError && node.channels && node.channels.length > 0 && (
                <div className="tree-children">
                  {node.channels.map((ch) => (
                    <div key={ch.id} className="tree-node">
                      <span className="toggle" style={{ visibility: 'hidden' }}>▸</span>
                      <span className="node-label">
                        # {ch.displayName}
                        {ch.membershipType === 'private' && (
                          <span style={{ color: '#999', fontSize: 11, marginLeft: 6 }}>private</span>
                        )}
                      </span>
                      {onSelect && (
                        <button
                          className="remove-btn"
                          style={{ color: '#0078d4', borderColor: '#0078d4' }}
                          onClick={() => onSelect(ch.id, `${node.team.name} / ${ch.displayName}`)}
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
