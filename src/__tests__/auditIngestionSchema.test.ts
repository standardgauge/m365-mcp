/**
 * Holds the audit ingestion contract together across infra/ and src/.
 *
 * The Logs Ingestion API drops a column the stream declaration doesn't name,
 * and the DCR drops one the table doesn't have, both without an error the
 * server would see. So the column list in infra/audit-ingestion.bicep must
 * cover every field toLogAnalyticsRecord can emit, and both top-level
 * templates must hand the app the rule's coordinates and the publisher role.
 */

import { readFileSync } from 'fs';
import { join } from 'path';
import { DEFAULT_AUDIT_STREAM, toLogAnalyticsRecord } from '../services/auditLogAnalytics.js';

const infra = (f: string) => readFileSync(join(__dirname, '..', '..', 'infra', f), 'utf8');

const moduleSrc = infra('audit-ingestion.bicep');

function bicepColumns(): Map<string, string> {
  const block = moduleSrc.match(/var columns = \[([\s\S]*?)\n\]/);
  if (!block) throw new Error('columns block not found in audit-ingestion.bicep');
  const cols = new Map<string, string>();
  for (const m of block[1].matchAll(/\{ name: '(\w+)', type: '(\w+)' \}/g)) cols.set(m[1], m[2]);
  return cols;
}

describe('audit ingestion schema', () => {
  const full = toLogAnalyticsRecord(
    {
      tenantId: 't', userId: 'u', userEmail: 'adele@fabrikam.com', deviceLabel: 'd',
      operation: 'o', resource: 'r', result: 'denied', reason: 'x', source: 'mcp', ip: '203.0.113.7',
      before: '[]', after: '["x"]',
    },
    '2026-01-02T03:04:05.678Z',
    'e',
  );

  it('declares exactly the columns the server emits', () => {
    expect([...bicepColumns().keys()].sort()).toEqual(Object.keys(full).sort());
  });

  it('types TimeGenerated as datetime and everything else as string', () => {
    for (const [name, type] of bicepColumns()) {
      expect([name, type]).toEqual([name, name === 'TimeGenerated' ? 'datetime' : 'string']);
    }
  });

  it('uses the stream name the server defaults to, and a _CL table', () => {
    expect(moduleSrc).toContain(`var streamName = '${DEFAULT_AUDIT_STREAM}'`);
    expect(moduleSrc).toMatch(/var tableName = '\w+_CL'/);
    expect(moduleSrc).toContain("kind: 'Direct'");
  });

  it.each(['main.bicep', 'container-app.bicep'])('%s wires the app to the rule with only the publisher role', (f) => {
    const src = infra(f);
    for (const v of ['AUDIT_LOGS_INGESTION_ENDPOINT', 'AUDIT_DCR_IMMUTABLE_ID', 'AUDIT_DCR_STREAM_NAME']) {
      expect(src).toContain(`name: '${v}'`);
    }
    // Monitoring Metrics Publisher, scoped to the audit DCR and nothing wider.
    expect(src).toContain("var monitoringMetricsPublisherRoleId = '3913510d-42f4-4e42-8a64-420c390055eb'");
    const assignment = src.match(/resource auditPublisherRoleAssignment[\s\S]*?\n\}/);
    expect(assignment?.[0]).toContain('scope: auditDcrRef');
    expect(assignment?.[0]).toContain('principalId: containerApp.identity.principalId');
  });
});
