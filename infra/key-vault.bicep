// key-vault.bicep — where the server's secrets live.
//
// Creates an RBAC-mode Key Vault holding the Entra client secret and the two
// credential-at-rest keys (session HMAC key, data encryption key), and lets the
// app's runtime identity read them. The Container App binds its secrets to
// these as Key Vault references, so the values are not Container App secrets:
// `az containerapp secret show` returns the reference, not the key.
//
// Reading a value takes a data-plane role on the vault (Key Vault Secrets User
// or Officer). Contributor on the resource group does not grant one, and
// granting it takes Owner or User Access Administrator. Every read is logged to
// the instance's Log Analytics workspace (AzureDiagnostics, category AuditEvent).
//
// Each deployment writes the values from its parameters as a new secret
// version, and the outputs are versioned URIs: the key in use changes only
// when the template is deployed with a different value. See "Application keys"
// in docs/operations-runbook.md before re-running with values that differ from
// the vault.

@description('Azure region.')
param location string

@description('Key Vault name: 3-24 characters, letters, digits and hyphens, globally unique.')
param vaultName string

@description('Principal id of the identity that reads the secrets at runtime.')
param readerPrincipalId string

@description('Resource id of the Log Analytics workspace that receives the vault audit log.')
param logAnalyticsWorkspaceId string

@secure()
param azureClientSecret string
@secure()
param mcpSessionHmacKey string
@secure()
param mcpDataEncryptionKey string

// Key Vault Secrets User: read secret values, nothing else.
var keyVaultSecretsUserRoleId = '4633458b-17de-408a-b874-0445c86b69e6'

resource vault 'Microsoft.KeyVault/vaults@2023-07-01' = {
  // checkov:skip=CKV_AZURE_109: the Container App runs on the Consumption profile without VNet integration, so it reaches the vault over public endpoints; Azure Container Apps is not a Key Vault trusted service. Reads need an Entra token with a data role on this vault. Network restriction arrives with VNet integration and private endpoints, tracked separately.
  // checkov:skip=CKV_AZURE_189: same reason as CKV_AZURE_109.
  name: vaultName
  location: location
  properties: {
    tenantId: subscription().tenantId
    sku: {
      family: 'A'
      name: 'standard'
    }
    enableRbacAuthorization: true
    enableSoftDelete: true
    softDeleteRetentionInDays: 90
    // Losing the data key loses every stored refresh token; nobody, including
    // a subscription owner, can purge a deleted secret before retention ends.
    enablePurgeProtection: true
    publicNetworkAccess: 'Enabled'
    networkAcls: {
      bypass: 'AzureServices'
      defaultAction: 'Allow'
    }
  }
}

resource clientSecret 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  // checkov:skip=CKV_AZURE_41: these secrets have no natural expiry; an expired key stops the container from starting. Rotation is a deliberate act (docs/operations-runbook.md, Rotation).
  parent: vault
  name: 'azure-client-secret'
  properties: {
    value: azureClientSecret
    contentType: 'text/plain'
  }
}

resource sessionHmacKey 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  // checkov:skip=CKV_AZURE_41: see clientSecret.
  parent: vault
  name: 'mcp-session-hmac-key'
  properties: {
    value: mcpSessionHmacKey
    contentType: 'text/plain'
  }
}

resource dataEncryptionKey 'Microsoft.KeyVault/vaults/secrets@2023-07-01' = {
  // checkov:skip=CKV_AZURE_41: see clientSecret.
  parent: vault
  name: 'mcp-data-encryption-key'
  properties: {
    value: mcpDataEncryptionKey
    contentType: 'text/plain'
  }
}

resource vaultAuditLog 'Microsoft.Insights/diagnosticSettings@2021-05-01-preview' = {
  scope: vault
  name: 'audit-to-log-analytics'
  properties: {
    workspaceId: logAnalyticsWorkspaceId
    logs: [
      {
        category: 'AuditEvent'
        enabled: true
      }
    ]
  }
}

resource readerRoleAssignment 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  scope: vault
  name: guid(vault.id, readerPrincipalId, keyVaultSecretsUserRoleId)
  properties: {
    roleDefinitionId: subscriptionResourceId('Microsoft.Authorization/roleDefinitions', keyVaultSecretsUserRoleId)
    principalId: readerPrincipalId
    principalType: 'ServicePrincipal'
  }
}

output vaultName string = vault.name
output vaultId string = vault.id
// A caller that reads these outputs waits for the whole module, so the reader
// role exists before anything resolves a reference to them.
output clientSecretUri string = clientSecret.properties.secretUriWithVersion
output sessionHmacKeyUri string = sessionHmacKey.properties.secretUriWithVersion
output dataEncryptionKeyUri string = dataEncryptionKey.properties.secretUriWithVersion
