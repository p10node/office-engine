import react from '@vitejs/plugin-react'
import { defineConfig } from 'vite'

/** fixture app for the browser tests (e2e/): samples come from the playground */
export default defineConfig({
  root: 'e2e/app',
  publicDir: '../../playground/public',
  plugins: [react()],
  logLevel: 'error',
  worker: { format: 'es' },
})
