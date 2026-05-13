import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],
  server: {
    host: true,
    port: 1500,
    proxy: {
      '/api': {
        target: 'http://127.0.0.1:1501',
        changeOrigin: true
      },
      '/process-excel': {
        target: 'http://127.0.0.1:1501',
        changeOrigin: true
      }
    }
  }
})
