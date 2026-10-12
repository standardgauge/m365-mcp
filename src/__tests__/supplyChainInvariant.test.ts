/**
 * Supply-chain invariant.
 *
 * A release has to carry an SBOM and a signed provenance attestation, and a
 * deployed image has to carry its provenance in the registry, or the claims in
 * docs/supply-chain.md stop being true without anything failing. Nothing else
 * reads these workflows between releases, so this holds the parts the doc
 * relies on in place.
 */
import { readFileSync } from 'fs';
import { join } from 'path';

const root = join(__dirname, '..', '..');
const read = (p: string): string => readFileSync(join(root, p), 'utf8');

// The text of one job: from its key to the next job key at the same indent.
function job(wf: string, name: string): string {
  const m = wf.match(new RegExp(`^  ${name}:\\n([\\s\\S]*?)(?=^  [\\w-]+:\\n|(?![\\s\\S]))`, 'm'));
  expect(m).not.toBeNull();
  return m![1];
}

describe('release workflow', () => {
  const wf = read(join('.github', 'workflows', 'release.yml'));
  const build = job(wf, 'build');
  const publish = job(wf, 'publish');

  it('runs on a published release and on pull requests that change the build', () => {
    expect(wf).toMatch(/release:\s*\n\s*types: \[published\]/);
    for (const path of ['.github/workflows/release.yml', 'Dockerfile', 'package.json', 'package-lock.json']) {
      expect(wf).toContain(`- ${path}`);
    }
  });

  it('is read-only by default and only the release-gated job can write', () => {
    expect(wf).toMatch(/^permissions:\n {2}contents: read\n/m);
    expect(build).toMatch(/contents: read/);
    expect(build).not.toMatch(/contents: write/);
    expect(publish).toMatch(/contents: write/);
    expect(publish).toContain("if: github.event_name == 'release'");
  });

  it('builds the SBOM from an image stamped with the commit, in SPDX and CycloneDX', () => {
    expect(build).toContain('--build-arg GIT_SHA="$GITHUB_SHA"');
    expect(build).toContain('spdx-json=release-assets/');
    expect(build).toContain('cyclonedx-json=release-assets/');
    // The completeness check runs before anything is signed or uploaded.
    expect(build.indexOf('production dependency')).toBeLessThan(build.indexOf('attest-build-provenance'));
  });

  it('attests every release asset and verifies it the way the doc tells a customer to', () => {
    expect(build).toMatch(/id-token: write/);
    expect(build).toMatch(/attestations: write/);
    expect(build).toMatch(/attest-build-provenance@[0-9a-f]{40} /);
    expect(build).toContain('subject-path: release-assets/*');
    expect(build).toContain('gh attestation verify');
    expect(build).toContain('--signer-workflow "${GITHUB_REPOSITORY}/.github/workflows/release.yml"');
    // Checksums are written before signing, so SHA256SUMS is a signed subject too.
    expect(build.indexOf('sha256sum')).toBeLessThan(build.indexOf('attest-build-provenance'));
  });

  it('never signs where it cannot: fork pull requests and private repositories', () => {
    expect(build).toContain('github.event.repository.private == false');
    expect(build).toContain('github.event.pull_request.head.repo.full_name == github.repository');
    expect(build).toMatch(/name: Attest build provenance\n\s*if: env\.ATTEST == 'true'/);
    expect(build).toMatch(/name: Verify the attestations\n\s*if: env\.ATTEST == 'true'/);
  });
});

describe('deploy workflow', () => {
  it('pushes the image with full build provenance', () => {
    const wf = read(join('.github', 'workflows', 'deploy.yml'));
    expect(wf).toMatch(/docker\/build-push-action@[0-9a-f]{40}[\s\S]*?provenance: mode=max/);
  });
});
