import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

export default defineConfig({
  root: 'playground',
  plugins: [react()],
  server: { port: 5173 },
  build: { outDir: '../dist-playground', emptyOutDir: true },
  worker: { format: 'es' },
})
