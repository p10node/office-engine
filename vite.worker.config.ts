import { defineConfig } from 'vite'

// Standalone module worker for strict-CSP hosts: <DocViewer worker={() => new Worker(url, { type: 'module' })} />
export default defineConfig({
  build: {
    outDir: 'dist',
    emptyOutDir: false,
    target: 'es2022',
    lib: { entry: 'src/core/worker/worker.ts', formats: ['es'], fileName: () => 'docs-engine.worker.js' },
  },
})
