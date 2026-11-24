#!/usr/bin/env bash
# release-source-check.sh: refuse to release a commit that is not tested main.
#
# A release is a tag on a commit of `main` whose CI passed. The SBOM and the
# provenance attestation prove what a release contains and which workflow built
# it from which commit; they say nothing about whether that commit was ever on
# `main` or passed its gates. release.yml runs this before it builds or signs
# anything, so an off-main or untested tag gets no official artifacts, and a
# deployer can run it against a published tag to check the same thing.
#
# It holds the rule in docs/supply-chain.md ("Release evidence") as code:
#
#   1. `main` contains the commit.
#   2. The newest `ci.yml` push run on `main` for exactly that commit completed,
#      and every job in it succeeded, except DCO Sign-off and Container Boot
#      (PR), which skip on a push by design.
#   3. The pull request that merged the commit into `main` has a `ci.yml` run on
#      its head commit, the newest of which completed with every job succeeding
#      except Container Scan (Trivy), which skips on a pull request by design.
#      Container Boot (PR) and the installer release gates run only there.
#
# Usage: scripts/release-source-check.sh --repo <owner/name> --sha <commit>
#          [--wait <seconds>] [--interval <seconds>]
#
# --wait keeps polling while a run it needs is missing or still in progress, for
# a release published moments after its merge. Default 0: decide now.
#
# Reads through `gh api`, so it needs `gh` signed in, or GH_TOKEN with read
# access to contents, actions and pull requests. Exit 0 clean, 1 when the commit
# breaks a rule above, 2 when the answer cannot be read (bad arguments, an API
# error) or a run was still unfinished when --wait ran out.
set -euo pipefail

REPO=""
SHA=""
WAIT=0
INTERVAL=30
BRANCH=main

usage() {
  echo "usage: release-source-check.sh --repo <owner/name> --sha <commit> [--wait <s>] [--interval <s>]" >&2
  exit 2
}

while [ $# -gt 0 ]; do
  case "$1" in
    --repo) REPO="${2:-}"; shift 2 || usage ;;
    --sha) SHA="${2:-}"; shift 2 || usage ;;
    --wait) WAIT="${2:-}"; shift 2 || usage ;;
    --interval) INTERVAL="${2:-}"; shift 2 || usage ;;
    *) usage ;;
  esac
done

[[ "$REPO" =~ ^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$ ]] || usage
# A full SHA, not a ref: a branch or tag name could move between this check and
# the build.
[[ "$SHA" =~ ^[0-9a-f]{40}$ ]] || { echo "release-source-check: --sha must be a full 40-character commit SHA" >&2; exit 2; }
[[ "$WAIT" =~ ^[0-9]+$ && "$INTERVAL" =~ ^[1-9][0-9]*$ ]] || usage

fail() { echo "release-source-check: $*" >&2; exit 1; }
unreadable() { echo "release-source-check: $*" >&2; exit 2; }

# gh api with the failure made visible. A command substitution would otherwise
# hand an empty string to jq and read as "nothing found".
api() {
  local out
  if ! out=$(gh api "$1" 2>&1); then
    unreadable "cannot read $1: ${out}"
  fi
  printf '%s' "$out"
}

# ── 1. main contains the commit ──────────────────────────────────────────────
# compare/<sha>...main is "ahead" when main is the commit plus later work, and
# "identical" when the commit is main's head. "behind" or "diverged" means the
# commit is not in main's history.
status=$(api "repos/${REPO}/compare/${SHA}...${BRANCH}" | jq -r '.status // empty')
case "$status" in
  ahead|identical) echo "on ${BRANCH}: ${SHA} (${status})" ;;
  "") unreadable "compare ${SHA}...${BRANCH} returned no status" ;;
  *) fail "${SHA} is not on ${BRANCH} (compare status: ${status}). A release is a tag on a commit of ${BRANCH}." ;;
esac

# ── CI runs ──────────────────────────────────────────────────────────────────
# Prints the id of the newest ci.yml run for a commit and event once it has
# completed. Polls until --wait runs out while that run is missing or unfinished.
newest_completed_run() {
  local head="$1" event="$2" filter="$3" deadline runs run state
  deadline=$(( $(date +%s) + WAIT ))
  while :; do
    # This runs inside a command substitution, where errexit is off, so a
    # failure exits by hand.
    runs=$(api "repos/${REPO}/actions/workflows/ci.yml/runs?head_sha=${head}&event=${event}&per_page=100") || exit $?
    # Newest by creation. A re-run is a new attempt of the same run, so this is
    # the run as it stands after any re-run.
    run=$(jq -c --arg head "$head" "[.workflow_runs[] | select(.head_sha == \$head) ${filter}]
      | sort_by(.created_at) | last // empty" <<<"$runs") || exit 2
    if [ -n "$run" ]; then
      state=$(jq -r '.status' <<<"$run")
      if [ "$state" = "completed" ]; then
        jq -r '.id' <<<"$run"
        return 0
      fi
    else
      state="not found"
    fi
    if [ "$(date +%s)" -ge "$deadline" ]; then
      unreadable "no completed ci.yml ${event} run for ${head} (newest: ${state}). Wait for CI, then re-run this workflow."
    fi
    echo "waiting for the ci.yml ${event} run for ${head} (${state})" >&2
    sleep "$INTERVAL"
  done
}

# Fails unless the run concluded success and every job did, other than the ones
# named in $2 (a JSON array of job names), which may also be skipped.
require_green() {
  local id="$1" may_skip="$2" label="$3" run conclusion bad
  run=$(api "repos/${REPO}/actions/runs/${id}")
  conclusion=$(jq -r '.conclusion // empty' <<<"$run")
  [ "$conclusion" = "success" ] \
    || fail "${label} run ${id} concluded '${conclusion:-none}', not success: $(jq -r '.html_url' <<<"$run")"
  bad=$(api "repos/${REPO}/actions/runs/${id}/jobs?filter=latest&per_page=100" \
    | jq -r --argjson skip "$may_skip" '
        if (.jobs | length) == 0 then "(no jobs)"
        else .jobs[]
          | select(.conclusion != "success"
                   and (.conclusion != "skipped" or (.name | IN($skip[]) | not)))
          | "\(.name): \(.conclusion // .status)"
        end')
  [ -z "$bad" ] || fail "${label} run ${id} has jobs that did not pass: ${bad//$'\n'/; }"
  echo "${label} run ${id}: green ($(jq -r '.html_url' <<<"$run"))"
}

# ── 2. the push run on main for the commit itself ────────────────────────────
push_run=$(newest_completed_run "$SHA" push "| select(.head_branch == \"${BRANCH}\")")
require_green "$push_run" '["DCO Sign-off", "Container Boot (PR)"]' "push"

# ── 3. the pull request that merged it ───────────────────────────────────────
# AGENTS.md: nothing lands on main without a pull request, squash merges only.
# So the commit is the merge commit of exactly one merged pull request into main.
pr=$(api "repos/${REPO}/commits/${SHA}/pulls" | jq -c --arg sha "$SHA" --arg base "$BRANCH" '
  [.[] | select(.merged_at != null and .merge_commit_sha == $sha and .base.ref == $base)] | first // empty')
[ -n "$pr" ] || fail "${SHA} is not the merge commit of a pull request into ${BRANCH}, so its pull-request CI (container boot, installer gates) cannot be found."
pr_number=$(jq -r '.number' <<<"$pr")
pr_head=$(jq -r '.head.sha' <<<"$pr")
echo "merged by #${pr_number} (head ${pr_head})"
pr_run=$(newest_completed_run "$pr_head" pull_request "")
require_green "$pr_run" '["Container Scan (Trivy)"]' "pull request #${pr_number}"

echo "release-source-check: ${SHA} is tested ${BRANCH}"
