import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { resolve } from 'path';

export default defineConfig({
  plugins: [react()],
  root: 'src/admin',
  build: {
    outDir: resolve(__dirname, 'dist/admin'),
    emptyOutDir: true,
  },
  define: {
    __MSAL_CLIENT_ID__: JSON.stringify(process.env.AZURE_CLIENT_ID ?? ''),
    __MSAL_TENANT_ID__: JSON.stringify(process.env.AZURE_TENANT_ID ?? ''),
  },
  server: {
    port: 5173,
    proxy: {
      '/api': {
        target: process.env.VITE_API_TARGET ?? 'http://127.0.0.1:7071',
        changeOrigin: true,
      },
    },
  },
});
