/**
 * scripts/identifier-scan.sh is the pre-flight gate against deployment-specific
 * names in the tree. It has to pass on a clean checkout whatever the git
 * layout: a worktree keeps a `.git` file that holds the primary clone's path,
 * and a scan that read git's own metadata failed a clean worktree whenever
 * that path contained a banned word. These cases run the real script against
 * throwaway repositories so that cannot come back silently.
 *
 * The banned word is assembled at runtime; written out, it would fail the
 * scan of this very file.
 */
import { spawnSync } from 'child_process';
import { copyFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

const script = join(__dirname, '..', '..', 'scripts', 'identifier-scan.sh');
const banned = ['prod', 'romou'].join('');

function git(cwd: string, ...args: string[]): void {
  const r = spawnSync('git', ['-c', 'user.name=t', '-c', 'user.email=t@example.com', ...args], {
    cwd,
    encoding: 'utf8',
  });
  if (r.status !== 0) throw new Error(`git ${args.join(' ')} failed: ${r.stderr}`);
}

function scan(cwd: string, env: NodeJS.ProcessEnv = {}): { status: number | null; out: string } {
  const r = spawnSync('bash', ['scripts/identifier-scan.sh'], {
    cwd,
    encoding: 'utf8',
    env: { ...process.env, ...env },
  });
  return { status: r.status, out: `${r.stdout}${r.stderr}` };
}

describe('identifier-scan.sh', () => {
  let tmp: string;
  let primary: string;
  let worktree: string;

  beforeEach(() => {
    tmp = mkdtempSync(join(tmpdir(), 'identifier-scan-'));
    // The primary clone sits under a path that contains a banned word, as it
    // does when the clone lives in a directory named after its owner.
    primary = join(tmp, `${banned}-primary`);
    mkdirSync(join(primary, 'scripts'), { recursive: true });
    copyFileSync(script, join(primary, 'scripts', 'identifier-scan.sh'));
    writeFileSync(join(primary, '.gitignore'), 'node_modules/\ndist/\n');
    writeFileSync(join(primary, 'README.md'), 'A fixture at example.com\n');
    git(primary, 'init', '-q');
    git(primary, 'add', '-A');
    git(primary, 'commit', '-q', '-m', 'init');
    worktree = join(tmp, 'wt');
    git(primary, 'worktree', 'add', '-q', worktree);
  });

  afterEach(() => {
    rmSync(tmp, { recursive: true, force: true });
  });

  it('exits 0 on a clean primary clone', () => {
    const r = scan(primary);
    expect(r.status).toBe(0);
    expect(r.out).toContain('identifier-scan: clean');
  });

  it('exits 0 on a clean worktree whose .git file names a banned path', () => {
    const r = scan(worktree);
    expect(r.status).toBe(0);
    expect(r.out).toContain('identifier-scan: clean');
    expect(r.out).not.toContain('.git');
  });

  it('exits 1 when a tracked file carries a banned identifier', () => {
    writeFileSync(join(worktree, 'README.md'), `deployed for ${banned}\n`);
    git(worktree, 'add', 'README.md');
    const r = scan(worktree);
    expect(r.status).toBe(1);
    expect(r.out).toContain('README.md:1:');
  });

  it('exits 1 when an untracked file carries a banned identifier', () => {
    // Pre-flight runs before anything is staged; a new file still counts.
    writeFileSync(join(worktree, 'notes.md'), `${banned} wrote this\n`);
    const r = scan(worktree);
    expect(r.status).toBe(1);
    expect(r.out).toContain('notes.md:1:');
  });

  it('ignores files .gitignore already excludes', () => {
    for (const dir of ['node_modules', 'dist']) {
      mkdirSync(join(worktree, dir));
      writeFileSync(join(worktree, dir, 'x.js'), `// ${banned}\n`);
    }
    expect(scan(worktree).status).toBe(0);
  });

  it('exits 2 rather than reporting clean when there is no repository to scan', () => {
    const bare = join(tmp, 'not-a-repo', 'scripts');
    mkdirSync(bare, { recursive: true });
    copyFileSync(script, join(bare, 'identifier-scan.sh'));
    // Stop git's discovery at the temp dir so a repository above it cannot
    // make this pass by accident.
    const r = scan(join(tmp, 'not-a-repo'), { GIT_CEILING_DIRECTORIES: tmp });
    expect(r.status).toBe(2);
    expect(r.out).not.toContain('identifier-scan: clean');
  });
});
