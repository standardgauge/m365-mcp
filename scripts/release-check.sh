#!/usr/bin/env bash
# release-check.sh: hold CHANGELOG.md, package.json and a release tag together.
#
# A deployment that pins a release reads two things before it moves: the tag,
# and the changelog section that says what changed and what the deployer has to
# do about it. Both are only worth pinning to if they agree with the code at the
# tag. This script is the check, so a release never ships with a version the
# tree does not carry or notes that were never written. docs/releasing.md is the
# process it enforces.
#
# Usage:
#   scripts/release-check.sh                 check CHANGELOG.md's structure and
#                                            that package.json matches its
#                                            newest release (CI, every PR)
#   scripts/release-check.sh --tag vX.Y.Z    the above, and that the tag names
#                                            package.json's version and has a
#                                            non-empty section (release.yml)
#   scripts/release-check.sh --notes X.Y.Z   print that section's body, for
#                                            `gh release create --notes-file`
#
# Runs from the repository root, or from the directory in RELEASE_CHECK_ROOT.
# Exit 0 clean, 1 when a check fails, 2 on a usage error or an unreadable file.
set -euo pipefail

ROOT="${RELEASE_CHECK_ROOT:-$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)}"
CHANGELOG="${ROOT}/CHANGELOG.md"
PACKAGE="${ROOT}/package.json"

MODE=check
ARG=
case "${1:-}" in
  '') ;;
  --tag|--notes)
    [ $# -eq 2 ] || { echo "release-check: $1 takes one argument" >&2; exit 2; }
    MODE="${1#--}"; ARG="$2" ;;
  *) echo "release-check: unknown argument '$1'" >&2; exit 2 ;;
esac

[ -r "$CHANGELOG" ] || { echo "release-check: cannot read ${CHANGELOG}" >&2; exit 2; }
[ -r "$PACKAGE" ] || { echo "release-check: cannot read ${PACKAGE}" >&2; exit 2; }

SEMVER='[0-9]+\.[0-9]+\.[0-9]+'

# The body of one version's section: everything after its heading up to the
# next level-two heading or the link references at the foot, blank lines at
# either end trimmed.
section_body() {
  awk -v want="## [$1]" '
    index($0, "## [") == 1 { on = (substr($0, 1, length(want)) == want); next }
    /^\[[^]]+\]: / { on = 0 }
    on { lines[++n] = $0 }
    END {
      s = 1; while (s <= n && lines[s] ~ /^[[:space:]]*$/) s++
      e = n; while (e >= s && lines[e] ~ /^[[:space:]]*$/) e--
      for (i = s; i <= e; i++) print lines[i]
    }' "$CHANGELOG"
}

if [ "$MODE" = notes ]; then
  echo "$ARG" | grep -qxE "$SEMVER" || { echo "release-check: '$ARG' is not X.Y.Z" >&2; exit 2; }
  body=$(section_body "$ARG")
  [ -n "$body" ] || { echo "release-check: CHANGELOG.md has no section for ${ARG}" >&2; exit 1; }
  printf '%s\n' "$body"
  exit 0
fi

fail=0
err() { echo "release-check: $*" >&2; fail=1; }

# package.json is plain JSON, so a line match is enough and keeps this free of
# a jq or node dependency: the first "version" key at the top level.
pkg_version=$(sed -nE 's/^  "version": "([^"]*)",?$/\1/p' "$PACKAGE" | head -1)
[ -n "$pkg_version" ] || err "package.json has no top-level version"

# Every level-two heading is either Unreleased, which comes first, or a release
# `## [X.Y.Z] - YYYY-MM-DD`, newest first and each version once.
headings=$(grep -nE '^## ' "$CHANGELOG" || true)
first=$(echo "$headings" | head -1 | cut -d: -f2-)
[ "$first" = "## [Unreleased]" ] || err "the first section must be '## [Unreleased]', found '${first:-nothing}'"

[ "$(grep -cxF '## [Unreleased]' "$CHANGELOG")" -le 1 ] || err "more than one Unreleased section"

versions=()
while IFS= read -r line; do
  [ -n "$line" ] || continue
  n=${line%%:*}; h=${line#*:}
  [ "$h" = "## [Unreleased]" ] && continue
  if echo "$h" | grep -qxE "## \[${SEMVER}\] - [0-9]{4}-[0-9]{2}-[0-9]{2}"; then
    v=$(echo "$h" | sed -E 's/^## \[([^]]+)\].*/\1/')
    for seen in ${versions[@]+"${versions[@]}"}; do
      [ "$seen" != "$v" ] || err "line ${n}: ${v} appears twice"
    done
    versions+=("$v")
  else
    err "line ${n}: '${h}' is not '## [X.Y.Z] - YYYY-MM-DD'"
  fi
done <<<"$headings"

if [ ${#versions[@]} -gt 1 ]; then
  sorted=$(printf '%s\n' "${versions[@]}" | sort -t. -k1,1nr -k2,2nr -k3,3nr)
  [ "$sorted" = "$(printf '%s\n' "${versions[@]}")" ] || err "releases are not newest first"
fi

# Within a section, only the Keep a Changelog groups plus the one this project
# adds for deployers.
while IFS= read -r line; do
  [ -n "$line" ] || continue
  h=${line#*:}
  case "$h" in
    '### Upgrade notes'|'### Added'|'### Changed'|'### Deprecated'|'### Removed'|'### Fixed'|'### Security') ;;
    *) err "line ${line%%:*}: '${h}' is not one of Upgrade notes, Added, Changed, Deprecated, Removed, Fixed, Security" ;;
  esac
done <<<"$(grep -nE '^### ' "$CHANGELOG" || true)"

# The tree says which release it is. Before the first release there is nothing
# to match; after it, package.json carries the newest released version until
# the next release bumps both together.
if [ ${#versions[@]} -gt 0 ] && [ -n "$pkg_version" ] && [ "$pkg_version" != "${versions[0]}" ]; then
  err "package.json is ${pkg_version} but the newest release in CHANGELOG.md is ${versions[0]}"
fi

if [ "$MODE" = tag ]; then
  if ! echo "$ARG" | grep -qxE "v${SEMVER}"; then
    err "tag '${ARG}' is not vX.Y.Z"
  else
    tv=${ARG#v}
    [ "$tv" = "$pkg_version" ] || err "tag ${ARG} but package.json is ${pkg_version}"
    [ -n "$(section_body "$tv")" ] || err "CHANGELOG.md has no section, or an empty one, for ${tv}"
  fi
fi

[ "$fail" -eq 0 ] || exit 1
echo "release-check: clean${ARG:+ for ${ARG}}"
