// audit-ingestion.bicep — the audit trail's destination in Log Analytics.
//
// Creates the M365McpAudit_CL custom table in the instance's workspace and a
// data collection rule (kind Direct, so it carries its own logs ingestion
// endpoint and needs no separate data collection endpoint) that accepts the
// Custom-M365McpAudit stream and routes it to that table.
//
// The server uploads one row per logAccess event (src/services/auditLogAnalytics.ts).
// The column list below, the stream declaration, and AuditLogAnalyticsRecord in
// that file must stay in step; src/__tests__/auditIngestionSchema.test.ts holds
// them together.
//
// Granting the Container App's identity permission to send is done by the
// caller, after the app exists: see the Monitoring Metrics Publisher role
// assignment scoped to this rule in main.bicep and container-app.bicep.

@description('Azure region; must match the workspace region.')
param location string

@description('Name of the existing Log Analytics workspace the table lives in.')
param workspaceName string

@description('Name for the data collection rule.')
param dataCollectionRuleName string

// Log Analytics accepts only certain values for these two, and anything else
// fails the whole deployment with InvalidParameter rather than being rounded:
//   retentionInDays       -1, or 4 through 730.
//   totalRetentionInDays  -1, 4 through 730, or a whole number of years past
//                         that: 1095, 1460, 1826, 2191, 2556, 2922, 3288, 3653,
//                         4018, 4383. Five years is 1826, not 1825.
// totalRetentionInDays must also be at least retentionInDays. The @maxValue
// bounds catch values past the top of each range at compile time; the gaps
// between the year values, and anything below 4 other than -1, only Azure
// checks. There is no @minValue because -1 sits below the resource schema's
// minimum of 4 and Bicep would warn on every build.
// The two -1 defaults differ: retentionInDays -1 takes the workspace retention,
// while totalRetentionInDays -1 takes the table's retentionInDays, so leaving
// total at -1 means no separate long-term retention.

@description('Days the audit table keeps rows interactively queryable: 4 through 730. -1 inherits the workspace default.')
@maxValue(730)
param retentionInDays int = -1

@description('Total days kept including long-term retention: 4 through 730, or one of 1095, 1460, 1826, 2191, 2556, 2922, 3288, 3653, 4018, 4383 (3 to 12 years; five years is 1826). -1 uses retentionInDays, with no separate long-term retention.')
@maxValue(4383)
param totalRetentionInDays int = -1

var tableName = 'M365McpAudit_CL'
var streamName = 'Custom-M365McpAudit'

// TenantId is reserved in every Log Analytics table (it holds the workspace
// id), so the Entra tenant is EntraTenantId.
var columns = [
  { name: 'TimeGenerated', type: 'datetime' }
  { name: 'EventId', type: 'string' }
  { name: 'EntraTenantId', type: 'string' }
  { name: 'UserId', type: 'string' }
  { name: 'UserEmail', type: 'string' }
  { name: 'DeviceLabel', type: 'string' }
  { name: 'Operation', type: 'string' }
  { name: 'TargetResource', type: 'string' }
  { name: 'Result', type: 'string' }
  { name: 'Reason', type: 'string' }
  { name: 'Source', type: 'string' }
  { name: 'ClientIp', type: 'string' }
  { name: 'Before', type: 'string' }
  { name: 'After', type: 'string' }
]

resource workspace 'Microsoft.OperationalInsights/workspaces@2022-10-01' existing = {
  name: workspaceName
}

resource auditTable 'Microsoft.OperationalInsights/workspaces/tables@2022-10-01' = {
  parent: workspace
  name: tableName
  properties: {
    plan: 'Analytics'
    retentionInDays: retentionInDays
    totalRetentionInDays: totalRetentionInDays
    schema: {
      name: tableName
      columns: columns
    }
  }
}

resource auditDcr 'Microsoft.Insights/dataCollectionRules@2023-03-11' = {
  name: dataCollectionRuleName
  location: location
  kind: 'Direct'
  dependsOn: [
    auditTable  // the output stream must exist before the rule can route to it
  ]
  properties: {
    streamDeclarations: {
      '${streamName}': {
        columns: columns
      }
    }
    destinations: {
      logAnalytics: [
        {
          name: 'auditWorkspace'
          workspaceResourceId: workspace.id
        }
      ]
    }
    dataFlows: [
      {
        streams: [ streamName ]
        destinations: [ 'auditWorkspace' ]
        transformKql: 'source'
        outputStream: 'Custom-${tableName}'
      }
    ]
  }
}

output dataCollectionRuleName string = auditDcr.name
output dataCollectionRuleImmutableId string = auditDcr.properties.immutableId
output logsIngestionEndpoint string = auditDcr.properties.endpoints.logsIngestion
output streamName string = streamName
output tableName string = tableName
