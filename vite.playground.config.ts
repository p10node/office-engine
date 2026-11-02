import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  root: 'playground',
  base: './', // relative assets: works at / (dev) and /office-engine/ (GitHub Pages)
  plugins: [react()],
  server: { port: 5173 },
  build: { outDir: '../dist-playground', emptyOutDir: true },
  worker: { format: 'es' },
})
