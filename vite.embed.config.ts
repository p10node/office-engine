import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'
import { thirdPartyLicenses } from './scripts/third-party-licenses.ts'

// Standalone <script> build (React bundled): window.OfficeEngine + <office-viewer> element.
export default defineConfig({
  plugins: [react(), thirdPartyLicenses()],
  define: { 'process.env.NODE_ENV': JSON.stringify('production') },
  build: {
    outDir: 'dist',
    emptyOutDir: false,
    target: 'es2020',
    lib: {
      entry: 'src/embed.ts',
      name: 'OfficeEngine',
      formats: ['iife'],
      fileName: () => 'office-engine.embed.js',
    },
  },
  worker: { format: 'es', plugins: () => [thirdPartyLicenses()] },
})
