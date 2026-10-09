import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { authenticateConsoleRequest, checkGlobalAdmin } from '../../services/authMiddleware.js';
import { queryAuditLog } from '../../services/auditLog.js';
import { getTenantIdFromSession } from '../../services/tokenCache.js';
import { withSecurity } from '../../services/securityHeaders.js';

async function getAuditLog(
  request: HttpRequest,
  _context: InvocationContext
): Promise<HttpResponseInit> {
  const auth = await authenticateConsoleRequest(request);
  if (!auth) {
    return { status: 401, jsonBody: { error: 'Authentication required' } };
  }

  const isAdmin = await checkGlobalAdmin(auth.userId);
  if (!isAdmin) {
    return { status: 403, jsonBody: { error: 'Global Administrator role required' } };
  }

  const tenantId = getTenantIdFromSession(auth.session);
  const userEmail = request.query.get('userEmail') ?? undefined;
  const startDate = request.query.get('startDate') ?? undefined;
  const endDate = request.query.get('endDate') ?? undefined;
  const operation = request.query.get('operation') ?? undefined;
  const resultFilter = (request.query.get('result') ?? undefined) as 'allowed' | 'denied' | undefined;
  const limit = Math.min(parseInt(request.query.get('limit') ?? '200', 10), 1000);

  const entries = await queryAuditLog(tenantId, {
    userEmail,
    startDate,
    endDate,
    operation,
    result: resultFilter,
    limit,
  });

  // CSV export
  if (request.query.get('format') === 'csv') {
    const headers = ['timestamp', 'userEmail', 'deviceLabel', 'operation', 'resource', 'result', 'reason', 'source', 'ip'];
    const rows = entries.map(e =>
      headers.map(h => {
        const v = (e as unknown as Record<string, unknown>)[h];
        const s = v == null ? '' : String(v);
        return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
      }).join(',')
    );
    const csv = [headers.join(','), ...rows].join('\n');
    return {
      status: 200,
      body: csv,
      headers: {
        'Content-Type': 'text/csv',
        'Content-Disposition': 'attachment; filename="audit-log.csv"',
      },
    };
  }

  return { status: 200, jsonBody: { entries, total: entries.length } };
}

app.http('getAuditLog', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'api/manage/audit-log',
  handler: withSecurity(getAuditLog),
});
