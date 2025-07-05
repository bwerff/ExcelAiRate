import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import basicSsl from '@vitejs/plugin-basic-ssl'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    react(),
    // Enable HTTPS for Office Add-in development
    basicSsl({
      name: 'excelairate',
      domains: ['localhost'],
      certDir: './.cert'
    })
  ],
  server: {
    https: true,
    port: 5173,
    headers: {
      "Access-Control-Allow-Origin": "*",
    }
  },
  build: {
    rollupOptions: {
      input: {
        main: 'index.html',
        runtime: 'runtime.html'
      }
    }
  }
})