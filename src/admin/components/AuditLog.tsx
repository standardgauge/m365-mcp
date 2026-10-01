import React, { useState, useEffect, useRef, useCallback } from 'react';

interface AuditEntry {
  timestamp: string;
  userEmail: string;
  deviceLabel?: string;
  operation: string;
  resource?: string;
  result: 'allowed' | 'denied';
  reason?: string;
  source: 'http' | 'mcp';
  ip?: string;
}

export default function AuditLog() {
  const [entries, setEntries] = useState<AuditEntry[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [total, setTotal] = useState(0);

  // Filter state
  const [userEmail, setUserEmail] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [operation, setOperation] = useState('');
  const [result, setResult] = useState('');

  const debounceRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const fetchEntries = useCallback(async (filters: {
    userEmail: string;
    startDate: string;
    endDate: string;
    operation: string;
    result: string;
  }) => {
    setLoading(true);
    setError(null);
    try {
      const params = new URLSearchParams();
      if (filters.userEmail) params.set('userEmail', filters.userEmail);
      if (filters.startDate) params.set('startDate', filters.startDate);
      if (filters.endDate) params.set('endDate', filters.endDate);
      if (filters.operation) params.set('operation', filters.operation);
      if (filters.result) params.set('result', filters.result);

      const res = await fetch(`/api/manage/audit-log?${params.toString()}`);
      if (!res.ok) {
        const data = await res.json() as { error?: string };
        throw new Error(data.error ?? `HTTP ${res.status}`);
      }
      const data = await res.json() as { entries: AuditEntry[]; total: number };
      setEntries(data.entries ?? []);
      setTotal(data.total ?? 0);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }, []);

  // Fetch on mount
  useEffect(() => {
    fetchEntries({ userEmail, startDate, endDate, operation, result });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Debounced re-fetch when filters change
  const triggerFetch = useCallback((filters: {
    userEmail: string;
    startDate: string;
    endDate: string;
    operation: string;
    result: string;
  }) => {
    if (debounceRef.current) clearTimeout(debounceRef.current);
    debounceRef.current = setTimeout(() => {
      fetchEntries(filters);
    }, 300);
  }, [fetchEntries]);

  const currentFilters = { userEmail, startDate, endDate, operation, result };

  const handleSearch = () => {
    if (debounceRef.current) clearTimeout(debounceRef.current);
    fetchEntries(currentFilters);
  };

  const handleExportCsv = async () => {
    const params = new URLSearchParams();
    if (userEmail) params.set('userEmail', userEmail);
    if (startDate) params.set('startDate', startDate);
    if (endDate) params.set('endDate', endDate);
    if (operation) params.set('operation', operation);
    if (result) params.set('result', result);
    params.set('format', 'csv');
    params.set('limit', '1000');

    try {
      const res = await fetch(`/api/manage/audit-log?${params.toString()}`);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'audit-log.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: unknown) {
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  return (
    <div style={{ fontFamily: 'inherit' }}>
      {/* Filter controls */}
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginBottom: '12px', alignItems: 'flex-end' }}>
        <label style={{ display: 'flex', flexDirection: 'column', gap: '2px', fontSize: '13px' }}>
          User Email
          <input
            type="text"
            placeholder="user@example.com"
            value={userEmail}
            onChange={e => { setUserEmail(e.target.value); triggerFetch({ ...currentFilters, userEmail: e.target.value }); }}
            style={{ padding: '4px 8px', border: '1px solid #ccc', borderRadius: '4px', fontSize: '13px' }}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: '2px', fontSize: '13px' }}>
          Start Date
          <input
            type="date"
            value={startDate}
            onChange={e => { setStartDate(e.target.value); triggerFetch({ ...currentFilters, startDate: e.target.value }); }}
            style={{ padding: '4px 8px', border: '1px solid #ccc', borderRadius: '4px', fontSize: '13px' }}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: '2px', fontSize: '13px' }}>
          End Date
          <input
            type="date"
            value={endDate}
            onChange={e => { setEndDate(e.target.value); triggerFetch({ ...currentFilters, endDate: e.target.value }); }}
            style={{ padding: '4px 8px', border: '1px solid #ccc', borderRadius: '4px', fontSize: '13px' }}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: '2px', fontSize: '13px' }}>
          Operation
          <input
            type="text"
            placeholder="e.g. read_file"
            value={operation}
            onChange={e => { setOperation(e.target.value); triggerFetch({ ...currentFilters, operation: e.target.value }); }}
            style={{ padding: '4px 8px', border: '1px solid #ccc', borderRadius: '4px', fontSize: '13px' }}
          />
        </label>
        <label style={{ display: 'flex', flexDirection: 'column', gap: '2px', fontSize: '13px' }}>
          Result
          <select
            value={result}
            onChange={e => { setResult(e.target.value); triggerFetch({ ...currentFilters, result: e.target.value }); }}
            style={{ padding: '4px 8px', border: '1px solid #ccc', borderRadius: '4px', fontSize: '13px' }}
          >
            <option value="">All</option>
            <option value="allowed">Allowed</option>
            <option value="denied">Denied</option>
          </select>
        </label>
        <div style={{ display: 'flex', gap: '8px', alignItems: 'flex-end' }}>
          <button
            onClick={handleSearch}
            style={{ padding: '5px 14px', background: '#0078d4', color: '#fff', border: 'none', borderRadius: '4px', cursor: 'pointer', fontSize: '13px' }}
          >
            Search
          </button>
          <button
            onClick={handleExportCsv}
            style={{ padding: '5px 14px', background: '#fff', color: '#333', border: '1px solid #ccc', borderRadius: '4px', cursor: 'pointer', fontSize: '13px' }}
          >
            Export CSV
          </button>
        </div>
      </div>

      {/* Status */}
      {loading && <p style={{ color: '#666', fontSize: '13px' }}>Loading...</p>}
      {error && <p style={{ color: '#c00', fontSize: '13px' }}>Error: {error}</p>}
      {!loading && !error && (
        <p style={{ fontSize: '12px', color: '#666', marginBottom: '8px' }}>{total} entries</p>
      )}

      {/* Table */}
      {!loading && entries.length > 0 && (
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px' }}>
            <thead>
              <tr style={{ background: '#f5f5f5', textAlign: 'left' }}>
                {['Timestamp', 'User', 'Device', 'Operation', 'Resource', 'Result', 'Reason'].map(col => (
                  <th key={col} style={{ padding: '6px 10px', borderBottom: '1px solid #ddd', fontWeight: 600, whiteSpace: 'nowrap' }}>{col}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {entries.map((entry, i) => (
                <tr key={i} style={{ background: i % 2 === 0 ? '#fff' : '#fafafa' }}>
                  <td style={{ padding: '5px 10px', borderBottom: '1px solid #eee', whiteSpace: 'nowrap', color: '#555' }}>
                    {new Date(entry.timestamp).toLocaleString()}
                  </td>
                  <td style={{ padding: '5px 10px', borderBottom: '1px solid #eee' }}>{entry.userEmail}</td>
                  <td style={{ padding: '5px 10px', borderBottom: '1px solid #eee', color: '#777' }}>{entry.deviceLabel ?? ''}</td>
                  <td style={{ padding: '5px 10px', borderBottom: '1px solid #eee', fontFamily: 'monospace' }}>{entry.operation}</td>
                  <td style={{ padding: '5px 10px', borderBottom: '1px solid #eee', color: '#777', maxWidth: '200px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{entry.resource ?? ''}</td>
                  <td style={{ padding: '5px 10px', borderBottom: '1px solid #eee' }}>
                    <span style={{
                      display: 'inline-block',
                      padding: '1px 8px',
                      borderRadius: '10px',
                      fontSize: '11px',
                      fontWeight: 600,
                      background: entry.result === 'allowed' ? '#e6f4ea' : '#fce8e6',
                      color: entry.result === 'allowed' ? '#1e7e34' : '#c0392b',
                    }}>
                      {entry.result}
                    </span>
                  </td>
                  <td style={{ padding: '5px 10px', borderBottom: '1px solid #eee', color: '#777' }}>{entry.reason ?? ''}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {!loading && !error && entries.length === 0 && (
        <p style={{ color: '#999', fontSize: '13px' }}>No audit entries found.</p>
      )}
    </div>
  );
}
