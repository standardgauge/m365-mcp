/**
 * Deploy CI gate.
 *
 * Every instance tracking this repository deploys automatically, so a deploy
 * that does not wait for CI ships a red main to all of them. The deploy
 * workflow used to trigger on push and ran alongside CI on the same commit,
 * never after it. This pins the trigger to a successful CI run on main and the
 * build to the commit CI actually passed on.
 */
import { readFileSync, readdirSync } from 'fs';
import { join } from 'path';

const root = join(__dirname, '..', '..');
const read = (p: string): string => readFileSync(join(root, p), 'utf8');

// The text between `on:` and the next top-level key.
function triggers(wf: string): string {
  const m = wf.match(/^on:\n([\s\S]*?)^\S/m);
  expect(m).not.toBeNull();
  return m![1];
}

describe('deploy CI gate', () => {
  const ci = read(join('.github', 'workflows', 'ci.yml'));

  it('CI is named the way the deploy trigger refers to it, and runs on pushes to main', () => {
    expect(ci).toMatch(/^name: CI$/m);
    expect(triggers(ci)).toMatch(/push:\s*\n\s*branches: \[main\]/);
  });

  // Discovered by content, as in containerPortInvariant.test.ts.
  const workflows = readdirSync(join(root, '.github', 'workflows')).filter(
    (f) => /\.ya?ml$/.test(f) && /az containerapp update/.test(read(join('.github', 'workflows', f))),
  );

  it('there is at least one workflow that deploys the Container App', () => {
    expect(workflows.length).toBeGreaterThan(0);
  });

  it.each(workflows)('%s deploys only after CI succeeds on main', (file) => {
    const wf = read(join('.github', 'workflows', file));
    const on = triggers(wf);
    expect(on).not.toMatch(/^\s*push:/m);
    expect(on).toMatch(/workflow_run:\s*\n\s*workflows: \[CI\]\s*\n\s*types: \[completed\]\s*\n\s*branches: \[main\]/);
    expect(wf).toContain("github.event.workflow_run.conclusion == 'success'");
    expect(wf).toContain("github.event.workflow_run.event == 'push'");
    expect(wf).toContain('github.event.workflow_run.head_repository.full_name == github.repository');
  });

  it.each(workflows)('%s builds and proves the commit CI passed on', (file) => {
    const wf = read(join('.github', 'workflows', file));
    expect(wf).toContain('DEPLOY_SHA: ${{ github.event.workflow_run.head_sha || github.sha }}');
    expect(wf).toContain('ref: ${{ env.DEPLOY_SHA }}');
    expect(wf).toContain('GIT_SHA=${{ env.DEPLOY_SHA }}');
    expect(wf).toContain('EXPECTED_SHA: ${{ env.DEPLOY_SHA }}');
    expect(wf).not.toMatch(/GITHUB_SHA/);
  });
});
