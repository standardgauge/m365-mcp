/**
 * scripts/release-source-check.sh is the gate release.yml runs before it
 * builds or signs a release: the tagged commit has to be on main, and its push
 * run and its pull request's run in ci.yml have to be green. Without it, a tag
 * on an off-main or untested commit got a signed SBOM and provenance like any
 * other release.
 *
 * These cases run the real script against a fake `gh` on PATH that answers
 * each `gh api <path>` from a fixture, so every way the gate can refuse is
 * exercised without a network.
 */
import { spawnSync } from 'child_process';
import { chmodSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

const root = join(__dirname, '..', '..');
const script = join(root, 'scripts', 'release-source-check.sh');

const REPO = 'example/m365-mcp';
const SHA = 'a'.repeat(40);
const PR_HEAD = 'b'.repeat(40);

// Answers `gh api <path>` with $FAKE_GH/<path, non-alphanumerics as _>.json.
// When <key>.<n>.json exists it answers the n-th call with it instead, so a
// test can script a run that is still in progress and then completes.
// A missing fixture is an API error.
const FAKE_GH = `#!/usr/bin/env bash
set -euo pipefail
[ "$1" = api ] || { echo "fake gh: only api is faked" >&2; exit 1; }
key=$(printf '%s' "$2" | tr -c 'A-Za-z0-9' '_')
n=$(( $(cat "$FAKE_GH/$key.count" 2>/dev/null || echo 0) + 1 ))
echo "$n" > "$FAKE_GH/$key.count"
for f in "$FAKE_GH/$key.$n.json" "$FAKE_GH/$key.json"; do
  [ -f "$f" ] && { cat "$f"; exit 0; }
done
echo "gh: Not Found (HTTP 404)" >&2
exit 1
`;

const key = (path: string): string => path.replace(/[^A-Za-z0-9]/g, '_');
const runsPath = (head: string, event: string): string =>
  `repos/${REPO}/actions/workflows/ci.yml/runs?head_sha=${head}&event=${event}&per_page=100`;

type Job = { name: string; conclusion: string | null; status?: string };

const PUSH_JOBS: Job[] = [
  { name: 'Lint & Type Check', conclusion: 'success' },
  { name: 'Test', conclusion: 'success' },
  { name: 'DCO Sign-off', conclusion: 'skipped' },
  { name: 'Container Boot (PR)', conclusion: 'skipped' },
  { name: 'Container Scan (Trivy)', conclusion: 'success' },
];
const PR_JOBS: Job[] = [
  { name: 'Lint & Type Check', conclusion: 'success' },
  { name: 'Test', conclusion: 'success' },
  { name: 'DCO Sign-off', conclusion: 'success' },
  { name: 'Container Boot (PR)', conclusion: 'success' },
  { name: 'Container Scan (Trivy)', conclusion: 'skipped' },
];

describe('release-source-check.sh', () => {
  let dir: string;
  let bin: string;

  const fixture = (path: string, body: unknown, call?: number): void =>
    writeFileSync(join(dir, `${key(path)}${call ? `.${call}` : ''}.json`), JSON.stringify(body));

  const run = (id: number, head: string, extra: Record<string, unknown> = {}) => ({
    id,
    head_sha: head,
    head_branch: 'main',
    head_repository: { full_name: REPO },
    status: 'completed',
    conclusion: 'success',
    created_at: '2026-11-24T20:00:00Z',
    html_url: `https://github.com/${REPO}/actions/runs/${id}`,
    ...extra,
  });

  const runDetail = (r: ReturnType<typeof run>): void => fixture(`repos/${REPO}/actions/runs/${r.id}`, r);
  const jobs = (id: number, list: Job[]): void =>
    fixture(`repos/${REPO}/actions/runs/${id}/jobs?filter=latest&per_page=100`, { jobs: list });

  const merged = (head: Record<string, unknown> = {}) => ({
    number: 42,
    merged_at: '2026-11-24T19:59:00Z',
    merge_commit_sha: SHA,
    base: { ref: 'main' },
    head: { sha: PR_HEAD, ref: 'feature', repo: { full_name: REPO }, ...head },
  });

  // Tested main: on main, both runs green, merged by #42 from `feature`.
  function green(): void {
    fixture(`repos/${REPO}/compare/${SHA}...main`, { status: 'ahead' });
    const push = run(1, SHA);
    fixture(runsPath(SHA, 'push'), { workflow_runs: [push] });
    runDetail(push);
    jobs(1, PUSH_JOBS);
    fixture(`repos/${REPO}/commits/${SHA}/pulls`, [merged()]);
    const pr = run(2, PR_HEAD, { head_branch: 'feature' });
    fixture(runsPath(PR_HEAD, 'pull_request'), { workflow_runs: [pr] });
    runDetail(pr);
    jobs(2, PR_JOBS);
  }

  function check(...extra: string[]): { status: number | null; out: string } {
    const r = spawnSync('bash', [script, '--repo', REPO, '--sha', SHA, ...extra], {
      encoding: 'utf8',
      env: { ...process.env, PATH: `${bin}:${process.env.PATH}`, FAKE_GH: dir },
    });
    return { status: r.status, out: `${r.stdout}${r.stderr}` };
  }

  beforeEach(() => {
    dir = mkdtempSync(join(tmpdir(), 'release-source-check-'));
    bin = join(dir, 'bin');
    mkdirSync(bin);
    writeFileSync(join(bin, 'gh'), FAKE_GH);
    chmodSync(join(bin, 'gh'), 0o755);
    green();
  });

  afterEach(() => {
    rmSync(dir, { recursive: true, force: true });
  });

  it('passes a commit of main whose push and pull-request runs are green', () => {
    const r = check();
    expect(r.out).toContain(`release-source-check: ${SHA} is tested main`);
    expect(r.status).toBe(0);
  });

  it("passes main's head (compare status identical)", () => {
    fixture(`repos/${REPO}/compare/${SHA}...main`, { status: 'identical' });
    expect(check().status).toBe(0);
  });

  it.each(['diverged', 'behind'])('refuses a commit that is not on main (%s)', (status) => {
    fixture(`repos/${REPO}/compare/${SHA}...main`, { status });
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('is not on main');
  });

  it('refuses a commit with no push run on main', () => {
    fixture(runsPath(SHA, 'push'), { workflow_runs: [] });
    const r = check();
    expect(r.status).toBe(2);
    expect(r.out).toContain('no completed ci.yml push run');
  });

  it('ignores a push run for the commit on another branch', () => {
    fixture(runsPath(SHA, 'push'), { workflow_runs: [run(1, SHA, { head_branch: 'release-candidate' })] });
    const r = check();
    expect(r.status).toBe(2);
    expect(r.out).toContain('no completed ci.yml push run');
  });

  it('refuses a push run that failed', () => {
    const failed = run(1, SHA, { conclusion: 'failure' });
    fixture(runsPath(SHA, 'push'), { workflow_runs: [failed] });
    runDetail(failed);
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain("concluded 'failure', not success");
  });

  it('judges the newest push run, not an older green one', () => {
    const older = run(1, SHA, { created_at: '2026-11-24T19:00:00Z' });
    const newer = run(3, SHA, { conclusion: 'failure', created_at: '2026-11-24T21:00:00Z' });
    fixture(runsPath(SHA, 'push'), { workflow_runs: [newer, older] });
    runDetail(newer);
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('push run 3');
  });

  it('refuses a push run that skipped a job that must run', () => {
    jobs(1, PUSH_JOBS.map((j) => (j.name === 'Test' ? { ...j, conclusion: 'skipped' } : j)));
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('Test: skipped');
  });

  it('refuses a run that reports no jobs', () => {
    jobs(1, []);
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('(no jobs)');
  });

  it('refuses a commit that no pull request merged into main', () => {
    fixture(`repos/${REPO}/commits/${SHA}/pulls`, [
      { number: 7, merged_at: null, merge_commit_sha: SHA, base: { ref: 'main' }, head: { sha: PR_HEAD } },
      { number: 8, merged_at: '2026-11-24T19:59:00Z', merge_commit_sha: SHA, base: { ref: 'other' }, head: { sha: PR_HEAD } },
    ]);
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('is not the merge commit of a pull request into main');
  });

  it('refuses a pull-request run where Container Boot (PR) was skipped', () => {
    jobs(2, PR_JOBS.map((j) => (j.name === 'Container Boot (PR)' ? { ...j, conclusion: 'skipped' } : j)));
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('Container Boot (PR): skipped');
  });

  it('refuses a pull-request run that failed a job', () => {
    jobs(2, PR_JOBS.map((j) => (j.name === 'DCO Sign-off' ? { ...j, conclusion: 'failure' } : j)));
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('DCO Sign-off: failure');
  });

  // The same commit can be the head of two pull requests. A newer green run on
  // the other branch must not stand in for the merging PR's failed run.
  it("judges the merging PR's own branch, not a newer green run on another branch with the same SHA", () => {
    const ours = run(2, PR_HEAD, { head_branch: 'feature', conclusion: 'failure', created_at: '2026-11-24T19:00:00Z' });
    const other = run(4, PR_HEAD, { head_branch: 'feature-copy', created_at: '2026-11-24T21:00:00Z' });
    fixture(runsPath(PR_HEAD, 'pull_request'), { workflow_runs: [other, ours] });
    runDetail(ours);
    runDetail(other);
    jobs(4, PR_JOBS);
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('pull request #42 run 2');
    expect(r.out).toContain("concluded 'failure', not success");
  });

  it('ignores a green run on a fork branch of the same name and SHA', () => {
    const ours = run(2, PR_HEAD, { head_branch: 'feature', conclusion: 'failure', created_at: '2026-11-24T19:00:00Z' });
    const fork = run(4, PR_HEAD, {
      head_branch: 'feature',
      head_repository: { full_name: 'someone/m365-mcp' },
      created_at: '2026-11-24T21:00:00Z',
    });
    fixture(runsPath(PR_HEAD, 'pull_request'), { workflow_runs: [fork, ours] });
    runDetail(ours);
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('pull request #42 run 2');
  });

  it("refuses when the only pull-request run for the SHA is on another branch", () => {
    fixture(runsPath(PR_HEAD, 'pull_request'), { workflow_runs: [run(4, PR_HEAD, { head_branch: 'feature-copy' })] });
    const r = check();
    expect(r.status).toBe(2);
    expect(r.out).toContain(`no completed ci.yml pull_request run for ${PR_HEAD} on ${REPO}:feature`);
  });

  it('refuses a merging PR whose head repository is gone, rather than matching on SHA alone', () => {
    fixture(`repos/${REPO}/commits/${SHA}/pulls`, [merged({ repo: null })]);
    const r = check();
    expect(r.status).toBe(1);
    expect(r.out).toContain('cannot be bound to it');
  });

  it('waits for a push run that is still in progress', () => {
    fixture(runsPath(SHA, 'push'), { workflow_runs: [run(1, SHA, { status: 'in_progress', conclusion: null })] }, 1);
    const r = check('--wait', '10', '--interval', '1');
    expect(r.out).toContain('waiting for the ci.yml push run');
    expect(r.status).toBe(0);
  });

  it('gives up on a run still in progress when the wait runs out', () => {
    fixture(runsPath(SHA, 'push'), { workflow_runs: [run(1, SHA, { status: 'in_progress', conclusion: null })] });
    const r = check();
    expect(r.status).toBe(2);
    expect(r.out).toContain('newest: in_progress');
  });

  it('fails closed on an API error rather than reading it as nothing found', () => {
    rmSync(join(dir, `${key(runsPath(PR_HEAD, 'pull_request'))}.json`));
    const r = check();
    expect(r.status).toBe(2);
    expect(r.out).toContain('HTTP 404');
  });

  it('refuses a ref where a full commit SHA is required', () => {
    const r = spawnSync('bash', [script, '--repo', REPO, '--sha', 'v1.0.0'], { encoding: 'utf8' });
    expect(r.status).toBe(2);
    expect(r.stderr).toContain('full 40-character commit SHA');
  });
});

describe('release.yml runs the source check', () => {
  const wf = readFileSync(join(root, '.github', 'workflows', 'release.yml'), 'utf8');
  const step = wf.indexOf('- name: Tagged commit is on main and passed CI');

  it('on a release, from main, against the tagged commit', () => {
    expect(step).toBeGreaterThan(-1);
    const body = wf.slice(step, wf.indexOf('- name:', step + 1));
    expect(body).toContain("if: github.event_name == 'release'");
    expect(body).toContain('git show FETCH_HEAD:scripts/release-source-check.sh');
    expect(body).toMatch(/--repo "\$GITHUB_REPOSITORY" --sha "\$GITHUB_SHA"/);
  });

  it.each(['- name: Pack the source', '- name: Build image', '- name: Attest build provenance'])(
    'before %s',
    (later) => {
      expect(wf.indexOf(later)).toBeGreaterThan(step);
    },
  );

  it('with read access to the runs and pull requests it checks', () => {
    expect(wf).toMatch(/^\s+actions: read\b/m);
    expect(wf).toMatch(/^\s+pull-requests: read\b/m);
  });
});
