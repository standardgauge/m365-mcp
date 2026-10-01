# ── Stage 1: build ────────────────────────────────────────────────────────────
# Keep this Node major aligned with the azure-functions/node:4-node20 runtime in
# stage 2. Dependabot major bumps of the node image are ignored (.github/dependabot.yml)
# so the builder never drifts ahead of the runtime; move both together on an LTS bump.
FROM node:20-slim AS builder
WORKDIR /build

COPY package*.json ./
RUN npm ci

COPY tsconfig.json vite.config.ts ./
COPY src/ ./src/

# TypeScript compile (outputs to dist/functions/, dist/services/, dist/mcp/)
RUN npm run build

# Vite admin UI compile (outputs to dist/admin/)
# AZURE_CLIENT_ID / AZURE_TENANT_ID are baked in at build time via vite define;
# pass them as build args so the SPA gets the correct values.
ARG AZURE_CLIENT_ID
ARG AZURE_TENANT_ID
ENV AZURE_CLIENT_ID=${AZURE_CLIENT_ID} \
    AZURE_TENANT_ID=${AZURE_TENANT_ID}
RUN npm run build:admin

# ── Stage 2: runtime ──────────────────────────────────────────────────────────
FROM mcr.microsoft.com/azure-functions/node:4-node20

# The Azure Functions host defaults to binding privileged port 80. A non-root
# user (see USER below, F8 /) cannot bind ports < 1024, so point
# ASPNETCORE_URLS at a non-privileged port. The Container App ingress
# targetPort is set to 8080 to match (infra/container-app.bicep,
# infra/main.bicep). The public 80→443 redirect is handled at the ACA
# ingress edge and is unaffected by this internal container port.
ENV AzureWebJobsScriptRoot=/home/site/wwwroot \
    AzureFunctionsJobHost__Logging__Console__IsEnabled=true \
    AzureFunctionsJobHost__Logging__LogLevel__Default=Information \
    FUNCTIONS_EXTENSION_VERSION=~4 \
    WEBSITE_NODE_DEFAULT_VERSION=~20 \
    ASPNETCORE_URLS=http://+:8080

WORKDIR /home/site/wwwroot

# Deployed commit SHA, surfaced at runtime by the GET /health endpoint. Set by
# the deploy workflow from ${GITHUB_SHA}; empty for a plain local `docker build`.
ARG GIT_SHA
ENV GIT_SHA=${GIT_SHA}

# Production dependencies only
COPY package*.json ./
RUN npm ci --omit=dev

# Compiled output: functions import from dist/services/ so copy the whole dist/
COPY --from=builder /build/dist/ ./dist/
COPY host.json ./

# Run the Functions host as an unprivileged user (F8). Any
# code-exec in a handler then runs as uid 10001 rather than root. Hand the
# user the app root it reads and the /home tree the host may create log/data
# dirs under during bootstrap; /tmp (extension-bundle unpack target) is already
# world-writable. The host binaries under /azure-functions-host are
# world-readable in the base image, so no chown there.
RUN groupadd --system --gid 10001 app \
    && useradd --system --uid 10001 --gid app --home-dir /home/site --shell /usr/sbin/nologin app \
    && chown -R app:app /home/site

USER app

EXPOSE 8080
