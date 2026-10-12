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
# author. Needs real history for the range, so a shallow clone fetches first.
# Merge commits are skipped: they carry no authorship of their own.
set -euo pipefail

BASE="${1:?usage: dco-check.sh <base> [head]}"
HEAD_REF="${2:-HEAD}"

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
done < <(git rev-list --no-merges "${BASE}..${HEAD_REF}")

if [ "${fail}" -ne 0 ]; then
  echo "dco-check: sign-off missing or mismatched in ${BASE}..${HEAD_REF} (see above)." >&2
  echo "dco-check: 'git commit --amend -s' fixes the last commit; rebase to sign a range." >&2
  exit 1
fi

echo "dco-check: clean (${checked} commit(s) in ${BASE}..${HEAD_REF})"
