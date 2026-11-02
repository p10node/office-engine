import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// Library build: ESM, React external, worker inlined (no extra asset to host).
export default defineConfig({
  plugins: [react()],
  build: {
    outDir: 'dist',
    emptyOutDir: false,
    sourcemap: true,
    target: 'es2022',
    lib: {
      entry: { index: 'src/index.ts', core: 'src/core/index.ts', embed: 'src/embed.ts' },
      formats: ['es'],
    },
    rollupOptions: {
      external: ['react', 'react-dom', 'react-dom/client', 'react/jsx-runtime'],
    },
  },
  worker: { format: 'es' },
})
