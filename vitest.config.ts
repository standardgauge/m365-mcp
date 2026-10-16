import { defineConfig, mergeConfig } from 'vitest/config';
import viteConfig from './vite.config';

// The admin SPA's component tests. Vite's root is src/admin, so the paths
// below are relative to it. The server's Jest suite (src/__tests__) is
// separate and does not match these files.
export default mergeConfig(
  viteConfig,
  defineConfig({
    test: {
      environment: 'jsdom',
      include: ['__tests__/**/*.test.tsx'],
      setupFiles: ['__tests__/setup.ts'],
      restoreMocks: true,
      unstubGlobals: true,
    },
  }),
);
