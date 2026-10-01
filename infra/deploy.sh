#!/usr/bin/env bash
# deploy.sh — Build, push, and deploy a new Container App revision.
#
# Usage:
#   ./infra/deploy.sh --tag v1.2.3 [--resource-group rg-m365-mcp] [--app-name m365-mcp]
#
# Prerequisites:
#   - az CLI logged in with Owner/Contributor on the target resource group
#   - Docker daemon running
#   - jq installed (for parsing az output)

set -euo pipefail

# ── Defaults ──────────────────────────────────────────────────────────────────
TAG=""
RESOURCE_GROUP="rg-m365-mcp"
APP_NAME="m365-mcp"
ACR_NAME="m365mcpacr"   # must match acrName in bicep (no hyphens)

# ── Argument parsing ──────────────────────────────────────────────────────────
while [[ $# -gt 0 ]]; do
  case "$1" in
    --tag)           TAG="$2";            shift 2 ;;
    --resource-group) RESOURCE_GROUP="$2"; shift 2 ;;
    --app-name)      APP_NAME="$2";       shift 2 ;;
    --acr-name)      ACR_NAME="$2";       shift 2 ;;
    *) echo "Unknown argument: $1" >&2; exit 1 ;;
  esac
done

if [[ -z "$TAG" ]]; then
  echo "Error: --tag is required (e.g. --tag v1.0.0)" >&2
  exit 1
fi

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

echo "==> Tag:            $TAG"
echo "==> Resource group: $RESOURCE_GROUP"
echo "==> App name:       $APP_NAME"
echo "==> ACR:            $ACR_NAME"

# ── 1. Resolve ACR login server ───────────────────────────────────────────────
ACR_SERVER=$(az acr show --name "$ACR_NAME" --resource-group "$RESOURCE_GROUP" \
  --query loginServer -o tsv)
echo "==> ACR login server: $ACR_SERVER"

# ── 2. ACR login ──────────────────────────────────────────────────────────────
az acr login --name "$ACR_NAME"

# ── 3. Docker build ───────────────────────────────────────────────────────────
# Pass AZURE_CLIENT_ID and AZURE_TENANT_ID so the Vite build bakes them in.
# Read from environment — never hard-code secrets in this script.
: "${AZURE_CLIENT_ID:?AZURE_CLIENT_ID env var must be set}"
: "${AZURE_TENANT_ID:?AZURE_TENANT_ID env var must be set}"

IMAGE="${ACR_SERVER}/${APP_NAME}:${TAG}"
echo "==> Building image: $IMAGE"

docker build \
  --build-arg AZURE_CLIENT_ID="$AZURE_CLIENT_ID" \
  --build-arg AZURE_TENANT_ID="$AZURE_TENANT_ID" \
  -t "$IMAGE" \
  "$REPO_ROOT"

# ── 4. Push to ACR ────────────────────────────────────────────────────────────
echo "==> Pushing $IMAGE"
docker push "$IMAGE"

# ── 5. Update Container App to new revision ───────────────────────────────────
echo "==> Updating Container App '$APP_NAME' to image tag '$TAG'"
az containerapp update \
  --name "$APP_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --image "$IMAGE"

echo ""
echo "==> Done. Active FQDN:"
az containerapp show \
  --name "$APP_NAME" \
  --resource-group "$RESOURCE_GROUP" \
  --query "properties.configuration.ingress.fqdn" -o tsv
