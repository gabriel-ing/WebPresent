import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      '@webpresent/pptx-engine': path.resolve(
        __dirname,
        'packages/pptx-engine/src/browser.ts',
      ),
    },
  },
  optimizeDeps: {
    exclude: ['adm-zip', 'fast-xml-parser'],
  },
  build: {
    rollupOptions: {
      external: [
        'fs',
        'fs/promises',
        'path',
        'node:fs',
        'node:fs/promises',
        'node:path',
        'adm-zip',
        'fast-xml-parser',
      ],
    },
  },
});
