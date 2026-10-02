#!/usr/bin/env bash
# identifier-scan.sh: refuse a tree that names a deployment, a tenant, a client
# or an internal ticket.
#
# This repository is the development repository for a server that other
# organisations deploy into their own tenants. Operational detail about any
# particular deployment (hostnames, resource names, the people and domains in a
# test fixture, the ticket a change was tracked under) belongs with that
# deployment, not here. The scan ran once by hand at the first public release;
# CI runs it on every pull request so nothing drifts back in.
#
# Usage: scripts/identifier-scan.sh [dir]   (default: the repository root)
# Exit 0 clean, 1 when a match remains. Matches are printed with file:line.
#
# Extend the pattern when a new deployment-specific word turns up in review.
# Fixtures use example.com, fabrikam.com, northwindtraders.com and the like;
# pick one of those rather than adding an exception here.
set -euo pipefail

ROOT="${1:-$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)}"
cd "$ROOT"

TICKET_PREFIXES='AC|AP|NEN|FCTO|LIFE|DATAFLOW|SKATTER|WOVED|GOODBAY|OPS'
PATTERN='prodromou|goodbay|somaequity|amitycoast|\bdataflow\b|\bsoma\b|skatter|woved|apk8s|confluence|stankovski|stsoma|optionsgroup|vendex|xterra|mcclure|clements|aburris|nimble|guardsman|crosstown|medley|('"${TICKET_PREFIXES}"')-[0-9]+'

if grep -rinE "${PATTERN}" . \
     --exclude-dir=node_modules --exclude-dir=.git --exclude-dir=dist \
     --exclude=package-lock.json --exclude=identifier-scan.sh; then
  echo "identifier-scan: deployment-specific identifiers found (see above)" >&2
  exit 1
fi
echo "identifier-scan: clean"
