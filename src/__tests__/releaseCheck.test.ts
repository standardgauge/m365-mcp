/**
 * scripts/release-check.sh holds CHANGELOG.md, package.json and a release tag
 * together. CI runs it on every pull request and release.yml runs it with
 * --tag before building a release, so a deployment that pins a tag gets the
 * code, the version and the notes it was promised. These cases run the real
 * script against throwaway trees.
 */
import { spawnSync } from 'child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'fs';
import { tmpdir } from 'os';
import { join } from 'path';

const script = join(__dirname, '..', '..', 'scripts', 'release-check.sh');

const HEAD = '# Changelog\n\nIntro.\n\n';

function run(
  root: string,
  ...args: string[]
): { status: number | null; out: string; stdout: string } {
  const r = spawnSync('bash', [script, ...args], {
    encoding: 'utf8',
    env: { ...process.env, RELEASE_CHECK_ROOT: root },
  });
  return { status: r.status, out: `${r.stdout}${r.stderr}`, stdout: r.stdout };
}

describe('release-check.sh', () => {
  let root: string;

  beforeEach(() => {
    root = mkdtempSync(join(tmpdir(), 'release-check-'));
  });

  afterEach(() => {
    rmSync(root, { recursive: true, force: true });
  });

  function tree(version: string, changelog: string): void {
    writeFileSync(
      join(root, 'package.json'),
      `{\n  "name": "m365-mcp",\n  "version": "${version}",\n  "license": "AGPL-3.0-only"\n}\n`,
    );
    writeFileSync(join(root, 'CHANGELOG.md'), HEAD + changelog);
  }

  const released =
    '## [Unreleased]\n\n### Fixed\n\n- pending\n\n' +
    '## [1.1.0] - 2026-12-01\n\n### Upgrade notes\n\n- Consent X.\n\n### Added\n\n- A tool.\n\n' +
    '## [1.0.0] - 2026-11-30\n\n### Added\n\n- First release.\n\n' +
    '[Unreleased]: https://example.com/compare/v1.1.0...HEAD\n[1.1.0]: https://example.com/v1.1.0\n';

  it("passes this repository's own changelog", () => {
    const r = spawnSync('bash', [script], { encoding: 'utf8' });
    expect(`${r.stdout}${r.stderr}`).toContain('clean');
    expect(r.status).toBe(0);
  });

  it('passes a changelog with only Unreleased, whatever package.json says', () => {
    tree('0.1.0', '## [Unreleased]\n\n### Added\n\n- Something.\n');
    expect(run(root).status).toBe(0);
  });

  it('passes released sections when package.json carries the newest', () => {
    tree('1.1.0', released);
    expect(run(root).status).toBe(0);
  });

  it('fails when package.json does not match the newest release', () => {
    tree('1.0.0', released);
    const r = run(root);
    expect(r.status).toBe(1);
    expect(r.out).toContain('package.json is 1.0.0 but the newest release in CHANGELOG.md is 1.1.0');
  });

  it('requires Unreleased first, and only once', () => {
    tree('1.0.0', '## [1.0.0] - 2026-11-30\n\n### Added\n\n- x\n\n## [Unreleased]\n');
    expect(run(root).out).toContain("the first section must be '## [Unreleased]'");

    tree('0.1.0', '## [Unreleased]\n\n## [Unreleased]\n');
    expect(run(root).out).toContain('more than one Unreleased section');
  });

  it('refuses a malformed release heading', () => {
    tree('1.0.0', '## [Unreleased]\n\n## [1.0.0]\n\n### Added\n\n- x\n');
    const r = run(root);
    expect(r.status).toBe(1);
    expect(r.out).toContain("is not '## [X.Y.Z] - YYYY-MM-DD'");

    tree('1.0', '## [Unreleased]\n\n## [1.0] - 2026-11-30\n');
    expect(run(root).status).toBe(1);
  });

  it('refuses releases out of order or repeated', () => {
    tree(
      '1.0.0',
      '## [Unreleased]\n\n## [1.0.0] - 2026-11-30\n\n- x\n\n## [1.10.0] - 2026-12-02\n\n- y\n',
    );
    expect(run(root).out).toContain('releases are not newest first');

    tree('1.0.0', '## [Unreleased]\n\n## [1.0.0] - 2026-11-30\n\n- x\n\n## [1.0.0] - 2026-11-29\n');
    expect(run(root).out).toContain('1.0.0 appears twice');
  });

  it('orders versions numerically, not as text', () => {
    tree(
      '1.10.0',
      '## [Unreleased]\n\n## [1.10.0] - 2026-12-02\n\n- y\n\n## [1.9.0] - 2026-11-30\n\n- x\n',
    );
    expect(run(root).status).toBe(0);
  });

  it('refuses a group outside the allowed set', () => {
    tree('0.1.0', '## [Unreleased]\n\n### Breaking\n\n- x\n');
    const r = run(root);
    expect(r.status).toBe(1);
    expect(r.out).toContain("'### Breaking' is not one of");
  });

  describe('--tag', () => {
    it('passes a tag that names package.json and has a section', () => {
      tree('1.1.0', released);
      const r = run(root, '--tag', 'v1.1.0');
      expect(r.status).toBe(0);
      expect(r.out).toContain('clean for v1.1.0');
    });

    it('fails a tag that is not package.json', () => {
      tree('1.1.0', released);
      expect(run(root, '--tag', 'v1.0.0').out).toContain('tag v1.0.0 but package.json is 1.1.0');
    });

    it('fails a tag with no section', () => {
      tree('1.2.0', '## [Unreleased]\n\n- x\n');
      expect(run(root, '--tag', 'v1.2.0').out).toContain('no section, or an empty one, for 1.2.0');
    });

    it('fails a tag with an empty section', () => {
      tree('1.2.0', '## [Unreleased]\n\n## [1.2.0] - 2026-12-03\n\n## [1.1.0] - 2026-12-01\n\n- x\n');
      expect(run(root, '--tag', 'v1.2.0').out).toContain('no section, or an empty one, for 1.2.0');
    });

    it('fails a tag whose section has group headings but no entries', () => {
      tree('1.2.0', '## [Unreleased]\n\n## [1.2.0] - 2026-12-03\n\n### Added\n\n');
      const r = run(root, '--tag', 'v1.2.0');
      expect(r.status).toBe(1);
      expect(r.out).toContain('the section for 1.2.0 has headings but no entries');

      tree(
        '1.2.0',
        '## [Unreleased]\n\n## [1.2.0] - 2026-12-03\n\n### Upgrade notes\n\n### Fixed\n\n' +
          '## [1.1.0] - 2026-12-01\n\n- x\n',
      );
      expect(run(root, '--tag', 'v1.2.0').out).toContain('has headings but no entries');
    });

    it('fails a tag that is not vX.Y.Z', () => {
      tree('1.1.0', released);
      for (const tag of ['1.1.0', 'v1.1', 'v1.1.0-rc.1', 'release-1']) {
        const r = run(root, '--tag', tag);
        expect(r.status).toBe(1);
        expect(r.out).toContain('is not vX.Y.Z');
      }
    });
  });

  describe('--notes', () => {
    it("prints one section's body, trimmed, without the next section or the link references", () => {
      tree('1.1.0', released);
      const r = run(root, '--notes', '1.1.0');
      expect(r.status).toBe(0);
      expect(r.stdout).toBe('### Upgrade notes\n\n- Consent X.\n\n### Added\n\n- A tool.\n');

      const last = run(root, '--notes', '1.0.0');
      expect(last.stdout).toBe('### Added\n\n- First release.\n');
    });

    it('refuses to print a section with headings but no entries', () => {
      tree('1.2.0', '## [Unreleased]\n\n## [1.2.0] - 2026-12-03\n\n### Added\n');
      const r = run(root, '--notes', '1.2.0');
      expect(r.status).toBe(1);
      expect(r.stdout).toBe('');
      expect(r.out).toContain('has headings but no entries');
    });

    it('does not match a version that only shares a prefix', () => {
      tree('1.0.10', '## [Unreleased]\n\n## [1.0.10] - 2026-12-01\n\n- ten\n');
      expect(run(root, '--notes', '1.0.1').status).toBe(1);
    });
  });

  it('exits 2 on a usage error or a missing file', () => {
    tree('1.0.0', '## [Unreleased]\n');
    expect(run(root, '--bogus').status).toBe(2);
    expect(run(root, '--tag').status).toBe(2);
    rmSync(join(root, 'CHANGELOG.md'));
    expect(run(root).status).toBe(2);
  });
});
