import { defineConfig } from 'vite'
import { thirdPartyLicenses } from './scripts/third-party-licenses.ts'

// Standalone module worker for strict-CSP hosts: <DocViewer worker={() => new Worker(url, { type: 'module' })} />
export default defineConfig({
  plugins: [thirdPartyLicenses()],
  build: {
    outDir: 'dist',
    emptyOutDir: false,
    target: 'es2022',
    lib: { entry: 'src/core/worker/worker.ts', formats: ['es'], fileName: () => 'office-engine.worker.js' },
  },
})
