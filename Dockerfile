# ── Stage 1: build ────────────────────────────────────────────────────────────
# Built on the same image as the runtime in stage 2. It installs Node 22 with
# npm, so the builder's Node major cannot drift from the runtime's, and every
# build pulls from one registry instead of two. The builder used to come from
# the ECR Public mirror of the Docker Official node image (and before that from
# Docker Hub). Both rate-limit anonymous pulls from shared CI runner addresses:
# on 2026-10-22 ECR Public answered every pull with 429 "Data limit exceeded",
# which failed CI and the deploy on a tracking fork and held the release.
# The runtime image is pulled from MCR on every build already. Keep the tag and
# digest identical to stage 2 and bump both together.
FROM mcr.microsoft.com/azure-functions/node:4.1055.250-3-node22@sha256:10cdbf38aef3a7083c67d8b51ed887750047734db635cdd3e363a88d56865682 AS builder
WORKDIR /build

COPY package*.json ./
RUN npm ci

COPY tsconfig.json vite.config.ts ./
COPY src/ ./src/

# TypeScript compile (outputs to dist/functions/, dist/services/, dist/mcp/)
RUN npm run build

# Vite admin UI compile (outputs to dist/admin/). The SPA carries no Entra
# configuration: sign-in runs on the server, so the image is the same for
# every tenant.
RUN npm run build:admin

# ── Stage 2: runtime ──────────────────────────────────────────────────────────
# Pinned to a versioned host build rather than the rolling `4-node22` tag. The
# Functions host is the ASP.NET Core server in front of every request, so a .NET
# runtime CVE in this image is reachable, not vendor noise. Microsoft publishes
# the versioned tag first and moves `4-node22` later: in October 2026
# 4.1055.250-1 shipped .NET 8.0.32 a day before the rolling tag did, while the
# rolling tag still served 8.0.31 with a CRITICAL ASP.NET Core advisory. Pinned,
# the patched host is one Dependabot or hand bump away instead of a wait on
# Microsoft's retag. Dependabot's docker watch proposes newer versioned tags;
# security updates skip its cooldown. See docs/trivy-suppressions.md.
FROM mcr.microsoft.com/azure-functions/node:4.1055.250-3-node22@sha256:10cdbf38aef3a7083c67d8b51ed887750047734db635cdd3e363a88d56865682

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
    WEBSITE_NODE_DEFAULT_VERSION=~22 \
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
#
# The one exception is /azure-functions-host/Secrets. The host opens its key
# store there on any request that carries a function key, which includes any
# request with a `code` query parameter, before it routes the request and
# whatever the function's authLevel. The OAuth redirect to /api/auth/callback
# always carries `?code=`, so when the host cannot create that directory every
# sign-in gets an empty 500 and the callback function never runs. Create it
# here and hand it to the app user.
#
# The user is named mcp, not app: the node22 base image already carries its own
# `app` user and group (uid/gid 1654), and a second groupadd of that name fails.
RUN groupadd --system --gid 10001 mcp \
    && useradd --system --uid 10001 --gid mcp --home-dir /home/site --shell /usr/sbin/nologin mcp \
    && chown -R mcp:mcp /home/site \
    && mkdir -p /azure-functions-host/Secrets \
    && chown mcp:mcp /azure-functions-host/Secrets

USER mcp

EXPOSE 8080

# Container health for Docker and anything else that reads the image's own
# check (Trivy DS-0026). Azure Container Apps ignores HEALTHCHECK and runs the
# probes in infra/probes.json instead; both hit the same route on the same port.
# Asserts the JSON body rather than the status alone, because an unmatched path
# can still answer 200 from the admin SPA (docs/operations-runbook.md, Health
# probe). Node rather than curl: the runtime image is not guaranteed to ship
# curl, and node is what it runs.
HEALTHCHECK --interval=30s --timeout=5s --start-period=120s --retries=3 \
    CMD ["node", "-e", "fetch('http://127.0.0.1:8080/health',{signal:AbortSignal.timeout(4000)}).then(r=>r.json()).then(b=>process.exit(b.status==='ok'?0:1),()=>process.exit(1))"]

# Validate MCP_SESSION_HMAC_KEY and MCP_DATA_ENCRYPTION_KEY before the Functions
# host starts. The keys are otherwise read lazily on first use, so a container
# missing one booted, passed /health, and failed on the first sign-in. With the
# check in front, a bad key exits the container: the new revision never becomes
# ready and the previous one keeps serving. The exec keeps the host as the
# container's main process. /opt/startup/start_nonappservice.sh is the base
# image's own CMD; src/__tests__/startupKeyCheck.test.ts pins this line.
CMD ["/bin/sh", "-c", "node dist/startup/checkKeys.js && exec /opt/startup/start_nonappservice.sh"]
