/**
 * Container-port invariant.
 *
 * The port the image binds is declared in three places that ship separately:
 * the Dockerfile (ASPNETCORE_URLS / EXPOSE), the Bicep ingress targetPort, and
 * the deploy workflows that converge a live tenant's ingress. When #110 moved
 * the image from 80 to 8080 the Bicep moved with it, but nothing applies Bicep
 * on deploy, so every tenant kept targeting 80 and every deploy from 2026-09-26
 * failed its smoke while the old image kept serving. This pins all three to one
 * value so they cannot drift apart again without a red test.
 */
import { readFileSync, readdirSync } from 'fs';
import { join } from 'path';

const root = join(__dirname, '..', '..');
const read = (p: string): string => readFileSync(join(root, p), 'utf8');

function imagePort(): string {
  const dockerfile = read('Dockerfile');
  const urls = dockerfile.match(/ASPNETCORE_URLS=http:\/\/\+:(\d+)/);
  const expose = dockerfile.match(/^EXPOSE (\d+)$/m);
  expect(urls).not.toBeNull();
  expect(expose).not.toBeNull();
  expect(expose![1]).toBe(urls![1]);
  return urls![1];
}

describe('container port invariant', () => {
  const port = imagePort();

  it('the image binds a non-privileged port, as its non-root user requires', () => {
    expect(Number(port)).toBeGreaterThanOrEqual(1024);
  });

  // Discovered by content, not by name: the public seed strips the tenant Bicep
  // and tenant workflows and overlays a generic deploy.yml, and this test has to
  // hold in both trees.
  const bicep = readdirSync(join(root, 'infra'))
    .filter((f) => f.endsWith('.bicep'))
    .map((f) => join('infra', f))
    .filter((f) => /targetPort:/.test(read(f)));

  it('at least one Bicep file declares an ingress targetPort', () => {
    expect(bicep.length).toBeGreaterThan(0);
  });

  it.each(bicep)('%s targets the port the image binds', (file) => {
    const ports = [...read(file).matchAll(/targetPort:\s*(\d+)/g)].map((m) => m[1]);
    for (const p of ports) expect(p).toBe(port);
  });

  const workflows = readdirSync(join(root, '.github', 'workflows')).filter(
    (f) => /\.ya?ml$/.test(f) && /az containerapp update/.test(read(join('.github', 'workflows', f))),
  );

  it('there is at least one workflow that deploys the Container App', () => {
    expect(workflows.length).toBeGreaterThan(0);
  });

  it.each(workflows)('%s converges ingress to the image port before the image update', (file) => {
    const wf = read(join('.github', 'workflows', file));
    const ensure = wf.indexOf('- name: Ensure ingress targets the container port');
    const update = wf.indexOf('- name: Update Container App');
    expect(ensure).toBeGreaterThan(-1);
    expect(update).toBeGreaterThan(ensure);
    const declared = wf.slice(ensure, update).match(/CONTAINER_PORT:\s*"(\d+)"/);
    expect(declared).not.toBeNull();
    expect(declared![1]).toBe(port);
  });
});
