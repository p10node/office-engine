import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { thirdPartyLicenses } from './scripts/third-party-licenses.ts'

// Library build: ESM, React external, worker inlined (no extra asset to host).
export default defineConfig({
  plugins: [react(), thirdPartyLicenses()],
  build: {
    outDir: 'dist',
    emptyOutDir: false,
    sourcemap: true,
    target: 'es2022',
    lib: {
      entry: { index: 'src/index.ts', core: 'src/core/index.ts', embed: 'src/embed.ts', collab: 'src/collab/index.ts' },
      formats: ['es'],
    },
    rollupOptions: {
      external: ['react', 'react-dom', 'react-dom/client', 'react/jsx-runtime'],
    },
  },
  worker: { format: 'es', plugins: () => [thirdPartyLicenses()] },
})
