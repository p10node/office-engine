import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

// Standalone <script> build (React bundled): window.DocsEngine + <docs-viewer> element.
export default defineConfig({
  plugins: [react()],
  define: { 'process.env.NODE_ENV': JSON.stringify('production') },
  build: {
    outDir: 'dist',
    emptyOutDir: false,
    target: 'es2020',
    lib: {
      entry: 'src/embed.ts',
      name: 'DocsEngine',
      formats: ['iife'],
      fileName: () => 'docs-engine.embed.js',
    },
  },
  worker: { format: 'es' },
})
