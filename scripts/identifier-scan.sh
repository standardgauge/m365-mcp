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
# Exit 0 clean, 1 when a match remains, 2 when the scan itself could not run.
# Matches are printed with file:line.
#
# The scan goes through `git grep` rather than a recursive grep of the working
# directory, for two reasons. It only looks at files git tracks or would track
# (`--untracked`), so a build output or dependency tree that `.gitignore`
# already excludes is never read and needs no exclusion list here. And it never
# reads git's own metadata: in a worktree `.git` is a file holding the path of
# the primary clone, and a recursive grep matched the owner's name in that path
# and failed a clean tree. Pre-flight runs from a worktree per branch, so that
# layout has to pass.
#
# Extend the pattern when a new deployment-specific word turns up in review.
# Fixtures use example.com, fabrikam.com, northwindtraders.com and the like;
# pick one of those rather than adding an exception here.
set -euo pipefail

ROOT="${1:-$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)}"
cd "$ROOT"

TICKET_PREFIXES='AC|AP|NEN|FCTO|LIFE|DATAFLOW|SKATTER|WOVED|GOODBAY|OPS'
PATTERN='prodromou|goodbay|somaequity|amitycoast|\bdataflow\b|\bsoma\b|skatter|woved|apk8s|confluence|stankovski|stsoma|optionsgroup|vendex|xterra|mcclure|clements|aburris|nimble|guardsman|crosstown|medley|('"${TICKET_PREFIXES}"')-[0-9]+'

# git grep exits 0 on a match, 1 on none, anything else on an error (not a
# repository, bad pattern). `set -e` would hide the difference, so take it by
# hand: a clean tree is the only path to exit 0.
status=0
git grep -inE --untracked "${PATTERN}" -- . \
  ':(exclude)package-lock.json' ':(exclude)scripts/identifier-scan.sh' || status=$?

case "$status" in
  0)
    echo "identifier-scan: deployment-specific identifiers found (see above)" >&2
    exit 1
    ;;
  1)
    echo "identifier-scan: clean"
    ;;
  *)
    echo "identifier-scan: could not scan ${ROOT} (git grep exited ${status})" >&2
    exit 2
    ;;
esac
