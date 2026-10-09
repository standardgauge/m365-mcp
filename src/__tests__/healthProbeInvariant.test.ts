/**
 * Health-probe invariant.
 *
 * Three things check the container's health and ship separately: the
 * Dockerfile HEALTHCHECK, the Container Apps probes in infra/probes.json (read
 * by the Bicep), and the deploy workflow that converges those probes onto a
 * live app, since nothing applies Bicep on deploy. All of them must hit the
 * route src/functions/health.ts registers, on the port the image binds. A probe
 * on the wrong path is worse than none: it restarts a healthy container.
 */
import { readFileSync, readdirSync } from 'fs';
import { join } from 'path';

const root = join(__dirname, '..', '..');
const read = (p: string): string => readFileSync(join(root, p), 'utf8');

interface Probe {
  type: string;
  httpGet?: { path: string; port: number; scheme?: string };
  initialDelaySeconds?: number;
  periodSeconds?: number;
  timeoutSeconds?: number;
  failureThreshold?: number;
  successThreshold?: number;
}

const imagePort = read('Dockerfile').match(/ASPNETCORE_URLS=http:\/\/\+:(\d+)/)![1];

// host.json sets routePrefix to "", so the registered route is the URL path.
const healthPath = (): string => {
  const prefix = JSON.parse(read('host.json')).extensions?.http?.routePrefix ?? 'api';
  const route = read('src/functions/health.ts').match(/route:\s*'([^']+)'/)![1];
  return `${prefix ? `/${prefix}` : ''}/${route}`;
};

describe('health probe invariant', () => {
  const path = healthPath();
  const probes: Probe[] = JSON.parse(read('infra/probes.json'));

  it('the health route is served at /health', () => {
    expect(path).toBe('/health');
  });

  it('infra/probes.json defines one startup, readiness and liveness probe', () => {
    expect(probes.map((p) => p.type).sort()).toEqual(['Liveness', 'Readiness', 'Startup']);
  });

  it.each(['Startup', 'Readiness', 'Liveness'])('the %s probe GETs the health route on the image port', (type) => {
    const probe = probes.find((p) => p.type === type)!;
    expect(probe.httpGet).toEqual({ path, port: Number(imagePort), scheme: 'HTTP' });
  });

  // Container Apps rejects a template outside these bounds, which fails the
  // deploy rather than the probe.
  it.each(probes.map((p) => [p.type, p] as const))('the %s probe is within Container Apps limits', (type, p) => {
    if (p.initialDelaySeconds !== undefined) {
      expect(p.initialDelaySeconds).toBeGreaterThanOrEqual(1);
      expect(p.initialDelaySeconds).toBeLessThanOrEqual(60);
    }
    for (const v of [p.periodSeconds, p.timeoutSeconds]) {
      if (v !== undefined) {
        expect(v).toBeGreaterThanOrEqual(1);
        expect(v).toBeLessThanOrEqual(240);
      }
    }
    for (const v of [p.failureThreshold, p.successThreshold]) {
      if (v !== undefined) {
        expect(v).toBeGreaterThanOrEqual(1);
        expect(v).toBeLessThanOrEqual(10);
      }
    }
    if (type !== 'Readiness' && p.successThreshold !== undefined) expect(p.successThreshold).toBe(1);
    // A probe that can time out after the next one fires overlaps itself.
    if (p.periodSeconds !== undefined && p.timeoutSeconds !== undefined) {
      expect(p.timeoutSeconds).toBeLessThan(p.periodSeconds);
    }
  });

  it('the Dockerfile HEALTHCHECK hits the health route on the image port', () => {
    const check = read('Dockerfile').match(/^HEALTHCHECK (?:[^\n]*\\\n)*[^\n]*/m);
    expect(check).not.toBeNull();
    expect(check![0]).toContain(`http://127.0.0.1:${imagePort}${path}`);
    // The SPA catch-all can answer 200 for an unmatched path; check the body.
    expect(check![0]).toMatch(/status\s*===\s*'ok'/);
  });

  // Discovered by content, as in containerPortInvariant.test.ts.
  const bicep = readdirSync(join(root, 'infra'))
    .filter((f) => f.endsWith('.bicep'))
    .map((f) => join('infra', f))
    .filter((f) => /Microsoft\.App\/containerApps@/.test(read(f)));

  it.each(bicep)('%s gives its container the probes from infra/probes.json', (file) => {
    const src = read(file);
    expect(src).toMatch(/var probes = loadJsonContent\('probes\.json'\)/);
    expect(src).toMatch(/^\s*probes: probes$/m);
  });

  const workflows = readdirSync(join(root, '.github', 'workflows')).filter(
    (f) => /\.ya?ml$/.test(f) && /az containerapp update/.test(read(join('.github', 'workflows', f))),
  );

  it.each(workflows)('%s converges probes from infra/probes.json before the image update', (file) => {
    const wf = read(join('.github', 'workflows', file));
    const ensure = wf.indexOf('- name: Ensure container probes target /health');
    const update = wf.indexOf('- name: Update Container App');
    expect(ensure).toBeGreaterThan(-1);
    expect(update).toBeGreaterThan(ensure);
    expect(wf.slice(ensure, update)).toContain('infra/probes.json');
  });
});
