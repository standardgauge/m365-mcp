# Documentation

- [Entra setup](entra-setup.md) — the app registration, redirect URIs, client
  secret, and the delegated permissions to consent, in order
- [Operations runbook](operations-runbook.md) — day-two operations: rotating
  secrets and keys, reading logs, the reserved-path trap, containment
- [Calendar access controls](calendar-access-controls.md) — how calendar reads
  and writes are gated
- [SharePoint search access controls](sharepoint-search-access-controls.md) —
  how deny lists and the site allow-list apply to both search surfaces
- [SharePoint MCP server spec](sharepoint-mcp-server-spec.md) — the SharePoint
  tool surface
- [Trivy suppressions](trivy-suppressions.md) — why each image-scan finding is
  suppressed, and what would un-suppress it

Deployment is [`.github/workflows/deploy.yml`](../.github/workflows/deploy.yml),
which is inert until you set `DEPLOY_ENABLED`. The Bicep templates are under
[`infra/`](../infra). Start from the README's
[Deploying your own instance](../README.md#deploying-your-own-instance).
