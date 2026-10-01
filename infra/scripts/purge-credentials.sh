#!/usr/bin/env bash
# purge-credentials.sh — Wipe mcpSessions and mcpMsalCache for a single tenant.
#
# This is a one-time migration step run AFTER deploying the hardened
# code. The pre- storage rows hold plaintext session tokens, plaintext
# Microsoft Graph access tokens, and plaintext MSAL refresh tokens — none of
# them are usable by the new code (which expects HMAC hashes and AES-GCM
# envelopes). The simplest migration is to delete every row and force users
# to re-OAuth via the install script.
#
# Usage:
#   ./infra/scripts/purge-credentials.sh \
#     --resource-group example-mcp-rg \
#     --app-name example-m365-mcp
#
# Optional flags:
#   --dry-run    List rows that would be deleted without actually deleting
#   --skip-msal  Only purge mcpSessions, leave mcpMsalCache alone (rare)
#
# Prerequisites:
#   - az CLI logged in with read/delete on the storage account behind the
#     Container App's AZURE_STORAGE_CONNECTION_STRING secret
#   - jq installed (for JSON parsing)

set -euo pipefail

RESOURCE_GROUP=""
APP_NAME=""
DRY_RUN=false
SKIP_MSAL=false

while [[ $# -gt 0 ]]; do
  case "$1" in
    --resource-group) RESOURCE_GROUP="$2"; shift 2 ;;
    --app-name)       APP_NAME="$2";       shift 2 ;;
    --dry-run)        DRY_RUN=true;        shift ;;
    --skip-msal)      SKIP_MSAL=true;      shift ;;
    *) echo "Unknown argument: $1" >&2; exit 1 ;;
  esac
done

if [[ -z "$RESOURCE_GROUP" || -z "$APP_NAME" ]]; then
  echo "Usage: $0 --resource-group <rg> --app-name <containerapp>" >&2
  exit 1
fi

CURRENT_USER=$(az account show --query user.name -o tsv 2>/dev/null || echo "?")
echo "==> Logged in as: $CURRENT_USER"
echo "==> Resource group: $RESOURCE_GROUP"
echo "==> App name: $APP_NAME"
if $DRY_RUN; then
  echo "==> DRY RUN — no rows will be deleted"
fi
echo ""

# ── Resolve the storage account connection string from the Container App ──
echo "==> Reading storage connection string from Container App secret..."
CONN=$(az containerapp secret show \
  --name "$APP_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --secret-name azure-storage-connection-string \
  --query value -o tsv 2>/dev/null)

if [[ -z "$CONN" ]]; then
  echo "ERROR: Could not read 'azure-storage-connection-string' secret from $APP_NAME" >&2
  exit 1
fi
echo "    OK"

# ── Helper: list-and-delete every row in a table ──
purge_table() {
  local table="$1"
  echo ""
  echo "==> Scanning table: $table"

  # Use az storage entity query to enumerate. Returns JSON array of entities.
  local entities
  entities=$(az storage entity query \
    --table-name "$table" \
    --connection-string "$CONN" \
    --query "items[].{pk:PartitionKey, rk:RowKey}" \
    -o json 2>/dev/null || echo "[]")

  local count
  count=$(echo "$entities" | jq 'length')
  echo "    Found $count rows"

  if [[ "$count" == "0" ]]; then
    return 0
  fi

  if $DRY_RUN; then
    echo "$entities" | jq -r '.[] | "    would delete: " + .pk + "/" + .rk'
    return 0
  fi

  # Delete each row.
  echo "$entities" | jq -c '.[]' | while read -r row; do
    local pk rk
    pk=$(echo "$row" | jq -r '.pk')
    rk=$(echo "$row" | jq -r '.rk')
    az storage entity delete \
      --table-name "$table" \
      --partition-key "$pk" \
      --row-key "$rk" \
      --connection-string "$CONN" >/dev/null 2>&1 || {
        echo "    WARN: failed to delete $pk/$rk"
      }
    echo "    deleted: $pk/$rk"
  done
}

purge_table mcpSessions
if ! $SKIP_MSAL; then
  purge_table mcpMsalCache
else
  echo ""
  echo "==> Skipping mcpMsalCache (--skip-msal)"
fi

echo ""
if $DRY_RUN; then
  echo "==> Dry run complete. Re-run without --dry-run to actually delete."
else
  echo "==> Purge complete."
  echo ""
  echo "    Next steps:"
  echo "      1. Verify the new container revision is healthy"
  echo "      2. Re-run install-mcp.sh against this tenant to re-issue your session"
  echo "         curl -fsSL https://<MCP_HOST>/install | bash"
  echo "      3. Tell any other users on this tenant to do the same"
fi
