// main.bicep — Full-stack deployment for M365 MCP server.
// Creates: Storage Account + Key Vault + Container Registry + Log Analytics
// (with the audit table and its data collection rule) + Application Insights +
// Container App + Custom Domain binding.
//
// No storage key exists anywhere in the deployment: the account has shared-key
// access disabled and the app reaches it with its runtime identity. The client
// secret and both credential-at-rest keys live in Key Vault (key-vault.bicep);
// the Container App holds references to them, not values.
//
// Usage:
//   az deployment group create \
//     --resource-group rg-m365-mcp \
//     --template-file infra/main.bicep \
//     --parameters infra/parameters.dev.json

@description('Azure region for all resources.')
param location string = resourceGroup().location

@description('Short name prefix used for all resource names.')
param appName string = 'm365-mcp'

@description('Container image tag to deploy.')
param imageTag string = 'latest'

@description('Custom domain for the Container App (e.g. your-mcp-host.example.com). Leave empty to skip.')
param customDomain string = ''

@description('Key Vault name, globally unique. The default is derived from the resource group.')
param keyVaultName string = 'kv-${uniqueString(resourceGroup().id)}'

// ── Secrets (passed in at deploy time — never hard-code) ──────────────────────
@secure()
param azureClientId string
@secure()
param azureClientSecret string
@secure()
param azureTenantId string

// — credential-at-rest hardening keys.
// Generate fresh values per environment with `openssl rand -hex 32` and pass
// them in at deploy time. NEVER reuse across tenants or environments. They,
// and azureClientSecret, are written to Key Vault, not to the Container App.
@secure()
@description('64-hex-char HMAC-SHA256 key for hashing session tokens at rest.')
param mcpSessionHmacKey string
@secure()
@description('64-hex-char AES-256-GCM data encryption key for access tokens and MSAL cache at rest.')
param mcpDataEncryptionKey string

// ── Derived names ──────────────────────────────────────────────────────────────
var acrName = replace('${appName}acr', '-', '')
var environmentName = '${appName}-env'
var containerAppName = appName
var storageAccountName = replace('st${appName}', '-', '')    // e.g. stm365mcp

var auditDcrName = '${appName}-audit-dcr'

// Startup, readiness and liveness probes against GET /health on the container
// port. Kept in probes.json rather than inline because the deploy workflow
// converges the same definition onto apps that were never deployed from Bicep;
// src/__tests__/containerPortInvariant.test.ts holds its port to the image's.
var probes = loadJsonContent('probes.json')

// Storage Table Data Contributor: read and write entities and create tables.
// The runtime identity's only role on the storage account.
var storageTableDataContributorRoleId = '0a9a7e1f-b9d0-4cc4-a60d-0319b160aaa3'

// AcrPull built-in role
var acrPullRoleId = '7f951dda-4ed3-4680-a7ca-43fe172d538d'

// Monitoring Metrics Publisher: the only role the Logs Ingestion API needs to
// accept an upload, and it is granted on the audit DCR alone.
var monitoringMetricsPublisherRoleId = '3913510d-42f4-4e42-8a64-420c390055eb'

// ── Storage Account (sessions, MSAL cache, deny lists) ────────────────────────
resource storageAccount 'Microsoft.Storage/storageAccounts@2023-01-01' = {
  // checkov:skip=CKV_AZURE_43: storage account name is derived from appName param; naming convention is compliant (lowercase alphanumeric, 3-24 chars) but Checkov cannot statically evaluate the var expression
  // checkov:skip=CKV_AZURE_206: ZRS provides zone redundancy within the region; cross-region (GRS) replication is unnecessary for ephemeral MCP session and MSAL cache data that can be repopulated on re-auth
  // checkov:skip=CKV_AZURE_35: Container App on Consumption workload profile reaches storage over its public endpoint; Azure Container Apps is not on the Storage trusted-services bypass list, so defaultAction:Deny would break runtime table storage access. Every request needs an Entra token with a data role (shared-key access is off). Network restriction requires VNet injection with private endpoint — separate architectural enhancement.
  name: storageAccountName
  location: location
  sku: {
    name: 'Standard_ZRS'  // zone-redundant within region; upgradeable to GRS if data durability SLA requires it
  }
  kind: 'StorageV2'
  properties: {
    supportsHttpsTrafficOnly: true
    minimumTlsVersion: 'TLS1_2'
    allowBlobPublicAccess: false
    // Entra authorization only: no account key, connection string or SAS
    // signed with one is accepted. The app uses runtimeIdentity below.
    allowSharedKeyAccess: false
    defaultToOAuthAuthentication: true
  }
}

// Pre-create the required tables via Table Services
resource tableService 'Microsoft.Storage/storageAccounts/tableServices@2023-01-01' = {
  parent: storageAccount
  name: 'default'
}

resource mcpSessionsTable 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-01-01' = {
  parent: tableService
  name: 'mcpSessions'
}

resource mcpMsalCacheTable 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-01-01' = {
  parent: tableService
  name: 'mcpMsalCache'
}

resource globalDenyListTable 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-01-01' = {
  parent: tableService
  name: 'GlobalDenyList'
}

resource userDenyListTable 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-01-01' = {
  parent: tableService
  name: 'UserDenyList'
}

resource serviceSettingsTable 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-01-01' = {
  parent: tableService
  name: 'serviceSettings'
}

resource allowedSitesTable 'Microsoft.Storage/storageAccounts/tableServices/tables@2023-01-01' = {
  parent: tableService
  name: 'allowedSites'
}

// ── Runtime identity: reads Key Vault and Table Storage ───────────────────────
// User-assigned for the same ordering reason as acrPullIdentity: the Container
// App resolves its Key Vault references while it is being created, so the
// identity and its vault role have to exist first.
resource runtimeIdentity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: '${appName}-runtime'
  location: location
}

resource storageDataRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: storageAccount
  name: guid(storageAccount.id, runtimeIdentity.id, storageTableDataContributorRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', storageTableDataContributorRoleId)
    principalId: runtimeIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Key Vault: client secret and credential-at-rest keys ──────────────────────
module keyVault 'key-vault.bicep' = {
  name: '${appName}-key-vault'
  params: {
    location: location
    vaultName: keyVaultName
    readerPrincipalId: runtimeIdentity.properties.principalId
    logAnalyticsWorkspaceId: logAnalyticsWorkspace.id
    azureClientSecret: azureClientSecret
    mcpSessionHmacKey: mcpSessionHmacKey
    mcpDataEncryptionKey: mcpDataEncryptionKey
  }
}

// ── Azure Container Registry ──────────────────────────────────────────────────
resource acr 'Microsoft.ContainerRegistry/registries@2023-07-01' = {
  // checkov:skip=CKV_AZURE_139: ACR public networking disabled after private endpoint is provisioned; Container Apps Consumption requires public routing until custom VNet integration is configured
  // checkov:skip=CKV_AZURE_163: Vulnerability scanning is enabled via Microsoft Defender for Containers at subscription scope, not at the registry resource level
  // checkov:skip=CKV_AZURE_166: ACR quarantine requires a dedicated quarantine processor pipeline to release images before Container Apps can pull them; enabling without this infrastructure blocks all deployments. Image security scanning is handled by Microsoft Defender for Containers at subscription scope.
  name: acrName
  location: location
  sku: {
    name: 'Standard'  // Standard SKU for better retention, content trust, and geo-replication capabilities
  }
  properties: {
    adminUserEnabled: false  // CKV_AZURE_137: use managed identity (AcrPull role) instead of admin credentials
  }
}

// ── User-assigned identity for ACR pull (must exist before Container App is created) ──
// Using user-assigned identity avoids the first-deploy ordering problem: system-assigned
// identity does not exist until after the Container App is provisioned, but image pull
// needs AcrPull permission during provisioning. User-assigned identity is created and
// granted AcrPull first, so image pull succeeds on first deploy.
resource acrPullIdentity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  name: '${appName}-acr-pull'
  location: location
}

// ── Grant user-assigned identity AcrPull on the registry (before Container App) ─
resource acrPullRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: acr
  name: guid(acr.id, acrPullIdentity.id, acrPullRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', acrPullRoleId)
    principalId: acrPullIdentity.properties.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Log Analytics Workspace (90-day retention for MCP access logs) ────────────
//: retentionInDays=90 satisfies the 90-day MCP access log target
// in the AI Tool Permissioning framework.
resource logAnalyticsWorkspace 'Microsoft.OperationalInsights/workspaces@2022-10-01' = {
  name: '${appName}-logs'
  location: location
  properties: {
    retentionInDays: 90
    sku: {
      name: 'PerGB2018'
    }
  }
}

// ── Audit trail: custom table + data collection rule (Logs Ingestion API) ─────
// The authoritative audit record. Every logAccess event lands in
// M365McpAudit_CL in this workspace; see infra/audit-ingestion.bicep and
// docs/operations-runbook.md (Audit trail in Log Analytics).
module auditIngestion 'audit-ingestion.bicep' = {
  name: '${appName}-audit-ingestion'
  params: {
    location: location
    workspaceName: logAnalyticsWorkspace.name
    dataCollectionRuleName: auditDcrName
  }
}

// ── Application Insights (workspace-based, linked to Log Analytics) ────────────
resource applicationInsights 'Microsoft.Insights/components@2020-02-02' = {
  name: '${appName}-insights'
  location: location
  kind: 'web'
  properties: {
    Application_Type: 'web'
    WorkspaceResourceId: logAnalyticsWorkspace.id
  }
}

// ── Container Apps Environment (Consumption) ───────────────────────────────────
resource cae 'Microsoft.App/managedEnvironments@2023-05-01' = {
  name: environmentName
  location: location
  properties: {
    zoneRedundant: false
    appLogsConfiguration: {
      destination: 'log-analytics'
      logAnalyticsConfiguration: {
        customerId: logAnalyticsWorkspace.properties.customerId
        sharedKey: logAnalyticsWorkspace.listKeys().primarySharedKey
      }
    }
  }
}

// ── Derived URLs (depend on custom domain) ─────────────────────────────────────
var effectiveDomain = !empty(customDomain) ? customDomain : '${containerAppName}.${cae.properties.defaultDomain}'
var oauthRedirectUri = 'https://${effectiveDomain}/api/auth/callback'
var frontendUrl = 'https://${effectiveDomain}/admin'

// ── Container App ──────────────────────────────────────────────────────────────
resource containerApp 'Microsoft.App/containerApps@2023-05-01' = {
  name: containerAppName
  location: location
  dependsOn: [
    acrPullRoleAssignment      // ensure AcrPull RBAC is assigned before Container App pulls from ACR
    storageDataRoleAssignment  // and the runtime identity can reach storage before the first request
  ]
  identity: {
    type: 'SystemAssigned, UserAssigned'
    userAssignedIdentities: {
      '${acrPullIdentity.id}': {}
      '${runtimeIdentity.id}': {}
    }
  }
  properties: {
    managedEnvironmentId: cae.id
    configuration: {
      ingress: {
        external: true
        // Container listens on 8080 so the Functions host can run as a non-root
        // user (privileged port 80 is unbindable non-root — see Dockerfile,
        // F8 /). Public 80→443 redirect is unaffected (ACA edge).
        targetPort: 8080
        transport: 'auto'
        allowInsecure: false
      }
      registries: [
        {
          server: acr.properties.loginServer
          identity: acrPullIdentity.id  // user-assigned identity; AcrPull granted before Container App is created
        }
      ]
      secrets: [
        {
          name: 'azure-client-id'
          value: azureClientId
        }
        // Key Vault references: the Container App stores the URI, the runtime
        // identity resolves the value. Versioned URIs, so a new version in the
        // vault does not rotate a key until a deployment points here.
        {
          name: 'azure-client-secret'
          keyVaultUrl: keyVault.outputs.clientSecretUri
          identity: runtimeIdentity.id
        }
        {
          name: 'azure-tenant-id'
          value: azureTenantId
        }
        {
          name: 'oauth-redirect-uri'
          value: oauthRedirectUri
        }
        {
          name: 'frontend-url'
          value: frontendUrl
        }
        {
          name: 'mcp-session-hmac-key'
          keyVaultUrl: keyVault.outputs.sessionHmacKeyUri
          identity: runtimeIdentity.id
        }
        {
          name: 'mcp-data-encryption-key'
          keyVaultUrl: keyVault.outputs.dataEncryptionKeyUri
          identity: runtimeIdentity.id
        }
        {
          name: 'appinsights-connection-string'
          value: applicationInsights.properties.ConnectionString
        }
      ]
    }
    template: {
      containers: [
        {
          name: containerAppName
          image: '${acr.properties.loginServer}/${containerAppName}:${imageTag}'
          resources: {
            cpu: json('0.5')
            memory: '1Gi'
          }
          probes: probes
          env: [
            { name: 'AZURE_CLIENT_ID',                 secretRef: 'azure-client-id' }
            { name: 'AZURE_CLIENT_SECRET',              secretRef: 'azure-client-secret' }
            { name: 'AZURE_TENANT_ID',                  secretRef: 'azure-tenant-id' }
            { name: 'OAUTH_REDIRECT_URI',               secretRef: 'oauth-redirect-uri' }
            { name: 'FRONTEND_URL',                     secretRef: 'frontend-url' }
            { name: 'MCP_SESSION_HMAC_KEY',             secretRef: 'mcp-session-hmac-key' }
            { name: 'MCP_DATA_ENCRYPTION_KEY',          secretRef: 'mcp-data-encryption-key' }
            { name: 'APPLICATIONINSIGHTS_CONNECTION_STRING', secretRef: 'appinsights-connection-string' }
            // Table Storage over Entra auth (src/services/storageClient.ts). Not
            // secrets: the endpoint is public and the client id names an identity
            // only this app can obtain tokens for.
            { name: 'AZURE_STORAGE_TABLE_ENDPOINT',     value: storageAccount.properties.primaryEndpoints.table }
            { name: 'AZURE_STORAGE_IDENTITY_CLIENT_ID', value: runtimeIdentity.properties.clientId }
            // Audit trail → Log Analytics. Not secrets: the endpoint and rule id
            // only work for a caller holding the publisher role on the rule.
            { name: 'AUDIT_LOGS_INGESTION_ENDPOINT',    value: auditIngestion.outputs.logsIngestionEndpoint }
            { name: 'AUDIT_DCR_IMMUTABLE_ID',           value: auditIngestion.outputs.dataCollectionRuleImmutableId }
            { name: 'AUDIT_DCR_STREAM_NAME',            value: auditIngestion.outputs.streamName }
            { name: 'FUNCTIONS_EXTENSION_VERSION',      value: '~4' }
            { name: 'WEBSITE_NODE_DEFAULT_VERSION',     value: '~20' }
            //: server-side default deny list for sensitive Outlook
            // mail folders. Enforced on top of the admin-managed table so these
            // folders are unreachable even before an admin seeds it and can't be
            // un-denied by clearing the table. Matched case-insensitively by
            // folder display name (see src/services/denyList.ts).
            { name: 'DEFAULT_MAIL_DENY_FOLDERS',        value: 'Finance,HR,Legal,IR,Management' }
          ]
        }
      ]
      scale: {
        minReplicas: 1
        maxReplicas: 3
        rules: [
          {
            name: 'http-scaling'
            http: {
              metadata: {
                concurrentRequests: '20'
              }
            }
          }
        ]
      }
    }
  }
}

// ── Let the Container App's system-assigned identity send audit rows ──────────
// Scoped to the audit DCR only, not the resource group or the workspace: the
// identity can upload to this one stream and read nothing.
resource auditDcrRef 'Microsoft.Insights/dataCollectionRules@2023-03-11' existing = {
  name: auditDcrName
}

resource auditPublisherRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: auditDcrRef
  name: guid(resourceGroup().id, auditDcrName, containerApp.id, monitoringMetricsPublisherRoleId)
  // Ordered after the DCR through containerApp, whose env reads the module's outputs.
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', monitoringMetricsPublisherRoleId)
    principalId: containerApp.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

// ── Outputs ────────────────────────────────────────────────────────────────────
output storageAccountName string = storageAccount.name
output acrLoginServer string = acr.properties.loginServer
output containerAppFqdn string = containerApp.properties.configuration.ingress.fqdn
output containerAppPrincipalId string = containerApp.identity.principalId
output acrPullIdentityId string = acrPullIdentity.id
output runtimeIdentityId string = runtimeIdentity.id
output keyVaultName string = keyVault.outputs.vaultName
output oauthRedirectUri string = oauthRedirectUri
output frontendUrl string = frontendUrl
output logAnalyticsWorkspaceName string = logAnalyticsWorkspace.name
output applicationInsightsName string = applicationInsights.name
output auditTableName string = auditIngestion.outputs.tableName
output auditDataCollectionRuleName string = auditIngestion.outputs.dataCollectionRuleName
