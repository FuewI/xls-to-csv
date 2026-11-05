import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// GitHub Pages serves this site from https://fuewi.github.io/xls-to-csv/
// (dev server: http://localhost:5173/xls-to-csv/)
export default defineConfig({
  base: '/xls-to-csv/',
  plugins: [react(), tailwindcss()],
  // SheetJS + code pages for legacy .xls are ~1 MB; acceptable for a local tool.
  build: { chunkSizeWarningLimit: 1500 },
})
