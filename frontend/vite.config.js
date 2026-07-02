import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig(({ command }) => ({
  base: command === 'build' ? '/oe-decipher/' : '/',
  plugins: [react()],
  server: {
    port: 5300,
    proxy: {
      '/api': 'http://127.0.0.1:8001',
    },
  },
}))
