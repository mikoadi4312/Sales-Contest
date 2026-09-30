import { defineConfig } from 'vite';
export default defineConfig({
  root: 'web', publicDir: false, worker: { format: 'es' },
  build: { outDir: '../dist', emptyOutDir: true, target: 'es2022' },
});
