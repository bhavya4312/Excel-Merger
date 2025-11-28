import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// [https://vitejs.dev/config/](https://vitejs.dev/config/)
export default defineConfig({
  plugins: [react()],
  // IMPORTANT: Replace 'excel-merger' with your actual GitHub repository name
  base: '/Excel-Merger/', 
})