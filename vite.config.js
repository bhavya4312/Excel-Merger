import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  base: '/Excel-Merger/',
  build: {
    target: 'esnext',
    cssMinify: true,
    minify: 'esbuild',
    chunkSizeWarningLimit: 600,
    rollupOptions: {
      output: {
        manualChunks: {
          'vendor-react': ['react', 'react-dom'],
          'vendor-jspdf': ['jspdf', 'jspdf-autotable'],
          'vendor-jszip': ['jszip'],
          'vendor-icons': ['lucide-react'],
        }
      }
    }
  }
})