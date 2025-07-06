import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import basicSsl from '@vitejs/plugin-basic-ssl'
import path from 'path'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    react(),
    // Enable HTTPS for Office Add-in development (temporarily disabled)
    // basicSsl({
    //   name: 'excelairate',
    //   domains: ['localhost'],
    //   certDir: './.cert'
    // })
  ],
  server: {
    https: false,
    port: 5173,
    headers: {
      "Access-Control-Allow-Origin": "https://localhost:3000",
    },
    cors: {
      origin: ['https://localhost:3000', 'https://localhost:5173', 'https://excel.office.com'],
      credentials: true
    }
  },
  build: {
    rollupOptions: {
      input: 'index.html'
    }
  }
})