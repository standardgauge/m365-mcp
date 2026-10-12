#!/usr/bin/env bash
# dco-check.sh: refuse a range whose commits are not signed off.
#
# Contributions to this repository arrive under the licence it already uses, so
# there is no contributor licence agreement and nothing is assigned. What stands
# in its place is a Developer Certificate of Origin sign-off on every commit,
# which puts the origin of each line on the record. CONTRIBUTING.md has what a
# contributor is certifying; https://developercertificate.org/ has the DCO text.
#
# The sign-off address must match the commit author: a trailer naming someone
# else certifies nothing about the person who wrote the commit.
#
# Usage: scripts/dco-check.sh <base> [head]   (head defaults to HEAD)
# Exit 0 clean, 1 when a commit is unsigned or its sign-off does not match the
# author, 2 when the range cannot be walked in full: a ref that does not
# resolve, no merge-base between the two, or a shallow clone whose cut falls
# inside the range. It never fetches; a shallow clone runs `git fetch
# --unshallow` (in CI, `fetch-depth: 0`) first.
# Merge commits are skipped: they carry no authorship of their own.
set -euo pipefail

BASE="${1:?usage: dco-check.sh <base> [head]}"
HEAD_REF="${2:-HEAD}"

# Resolve the range before walking it. A loop fed by process substitution never
# sees rev-list's exit status, so an unresolvable ref (an unfetched origin/main,
# a typo, a shallow clone missing the base) would walk zero commits and report
# clean. Fail closed instead.
for ref in "${BASE}" "${HEAD_REF}"; do
  if ! git rev-parse --verify --quiet "${ref}^{commit}" >/dev/null; then
    echo "dco-check: cannot resolve '${ref}' to a commit; fetch it first." >&2
    exit 2
  fi
done

# Both refs resolving is not enough. In a shallow clone the tip of each branch
# can be present while the commits between them are not: rev-list then stops at
# the shallow cut and reports whatever it managed to walk as the whole range.
# Two checks close that. The refs must share a merge-base, which a cut above the
# fork point removes, and no commit inside the range may be a shallow graft,
# which a cut below the fork point (reachable through another parent of a merge)
# would leave behind.
if ! git merge-base "${BASE}" "${HEAD_REF}" >/dev/null 2>&1; then
  echo "dco-check: no merge-base between '${BASE}' and '${HEAD_REF}'; the" \
       "history between them is missing (shallow clone?). Fetch it first." >&2
  exit 2
fi

if ! commits="$(git rev-list --no-merges "${BASE}..${HEAD_REF}")"; then
  echo "dco-check: git rev-list failed for ${BASE}..${HEAD_REF}." >&2
  exit 2
fi

if [ "$(git rev-parse --is-shallow-repository)" = "true" ]; then
  shallow_file="$(git rev-parse --git-path shallow)"
  # grep reads the whole list rather than stopping at the first match, so
  # rev-list is never cut off by SIGPIPE, which pipefail would report as a
  # failed pipeline and turn a hit into a miss.
  if [ -s "${shallow_file}" ] \
     && git rev-list "${BASE}..${HEAD_REF}" | grep -xF -f "${shallow_file}" >/dev/null; then
    echo "dco-check: ${BASE}..${HEAD_REF} crosses a shallow-clone boundary, so" \
         "part of the range is not available to check. Fetch it first." >&2
    exit 2
  fi
fi

fail=0
checked=0

while read -r sha; do
  [ -n "${sha}" ] || continue
  checked=$((checked + 1))
  author="$(git show -s --format='%ae' "${sha}" | tr '[:upper:]' '[:lower:]')"
  subject="$(git show -s --format='%s' "${sha}")"
  signoffs="$(
    git show -s --format='%B' "${sha}" \
      | sed -n 's/^[Ss]igned-off-by:[[:space:]]*.*<\([^>]*\)>.*/\1/p' \
      | tr '[:upper:]' '[:lower:]'
  )"

  if [ -z "${signoffs}" ]; then
    echo "${sha:0:12} ${subject}: no Signed-off-by trailer" >&2
    fail=1
    continue
  fi

  if ! printf '%s\n' "${signoffs}" | grep -qxF "${author}"; then
    echo "${sha:0:12} ${subject}: signed off by $(printf '%s' "${signoffs}" | tr '\n' ' ')" \
         "but authored by ${author}" >&2
    fail=1
  fi
done <<< "${commits}"

if [ "${fail}" -ne 0 ]; then
  echo "dco-check: sign-off missing or mismatched in ${BASE}..${HEAD_REF} (see above)." >&2
  echo "dco-check: 'git commit --amend -s' fixes the last commit; rebase to sign a range." >&2
  exit 1
fi

echo "dco-check: clean (${checked} commit(s) in ${BASE}..${HEAD_REF})"
