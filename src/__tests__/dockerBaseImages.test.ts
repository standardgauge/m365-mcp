/**
 * Base-image registry invariant.
 *
 * CI builds the Dockerfile on every pull request and every push to main, with
 * no registry login. A bare `FROM node:20-slim` resolves to Docker Hub, which
 * rate-limits anonymous pulls: on 2026-10-09 four consecutive main runs failed
 * at the build with 429 Too Many Requests and main stayed red for hours with no
 * code defect. Every base image therefore names its registry explicitly, never
 * Docker Hub, and carries a digest so the mirror cannot serve a different image
 * under the same tag.
 *
 * The builder's Node major also has to match the azure-functions runtime's, so
 * native modules compiled in stage 1 load in stage 2 (.github/dependabot.yml).
 * The runtime is read from either the rolling `4-node22` form or a versioned
 * host build such as `4.1055.250-1-node22`.
 *
 * The production Dockerfile goes further: its builder stage runs on the
 * runtime image itself, so stage 1 pulls from MCR rather than ECR Public, which
 * rate-limited CI builds. The two `FROM` lines are separate literals,
 * so the test pins them equal, tag and digest, against a one-sided bump.
 */
import { readFileSync, readdirSync } from 'fs';
import { join } from 'path';

const root = join(__dirname, '..', '..');
const read = (p: string): string => readFileSync(join(root, p), 'utf8');

const dockerfiles = readdirSync(root).filter((f) => /^Dockerfile(\..+)?$/.test(f));

function baseImages(file: string): string[] {
  return [...read(file).matchAll(/^FROM\s+(?:--platform=\S+\s+)?(\S+)/gm)].map((m) => m[1]);
}

// A reference names a registry when its first path component looks like a
// host (contains a dot or a port, or is localhost), per the docker reference
// grammar. Anything else is shorthand for Docker Hub.
function registryOf(image: string): string {
  const first = image.split('/')[0];
  const isHost = image.includes('/') && (/[.:]/.test(first) || first === 'localhost');
  return isHost ? first : 'docker.io';
}

describe('Docker base images', () => {
  it('there is at least one Dockerfile', () => {
    expect(dockerfiles).toContain('Dockerfile');
  });

  const images = dockerfiles.flatMap((file) => baseImages(file).map((image) => [file, image]));

  it.each(images)('%s: %s is not pulled from Docker Hub', (_file, image) => {
    expect(registryOf(image)).not.toMatch(/^(docker\.io|index\.docker\.io|registry-1\.docker\.io)$/);
  });

  const nodeImages = images.filter(([, image]) => /\/library\/node[:@]/.test(image));

  it.each(nodeImages)('%s: %s is pinned by digest', (_file, image) => {
    expect(image).toMatch(/@sha256:[0-9a-f]{64}$/);
  });

  it.each(dockerfiles.filter((f) => /\/library\/node[:@]/.test(read(f)) && /azure-functions\/node/.test(read(f))))(
    '%s: the builder Node major matches the Functions runtime',
    (file) => {
      const bases = baseImages(file);
      const builder = bases.map((b) => b.match(/\/library\/node:(\d+)/)).find(Boolean);
      const runtime = bases.map((b) => b.match(/azure-functions\/node:\d[\d.-]*-node(\d+)/)).find(Boolean);
      expect(builder).toBeTruthy();
      expect(runtime).toBeTruthy();
      expect(builder![1]).toBe(runtime![1]);
    },
  );

  it('Dockerfile: the builder stage uses the runtime image, tag and digest', () => {
    const bases = baseImages('Dockerfile');
    expect(bases.length).toBeGreaterThanOrEqual(2);
    const builder = bases[0];
    const runtime = bases[bases.length - 1];
    expect(runtime).toMatch(/^mcr\.microsoft\.com\/azure-functions\/node:[^@]+@sha256:[0-9a-f]{64}$/);
    expect(builder).toBe(runtime);
  });
});
