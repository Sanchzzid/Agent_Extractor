import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

// In dev, /api is proxied to the backend so no CORS setup is needed.
// Set the backend address with VITE_DEV_API in frontend/.env (e.g. http://localhost:8034).
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  return {
    plugins: [react()],
    server: {
      port: 5173,
      proxy: { '/api': env.VITE_DEV_API || 'http://localhost:8031' },
    },
  }
})
