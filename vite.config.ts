import { defineConfig } from 'vite';
import { resolve } from 'path';

export default defineConfig({
  build: {
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      formats: ['es'],
      fileName: 'index',
    },
    rollupOptions: {
      external: /^lit/,
    },
  },
  test: {
    exclude: ['e2e/**', 'node_modules/**'],
  },
});
