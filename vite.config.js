import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// GitHub Pages serves this site from https://fuewi.github.io/excel-to-csv/
// Production builds (and `npm run preview`) use that sub-path; `npm run dev` keeps using "/".
export default defineConfig(({ command, isPreview }) => ({
  base: command === 'build' || isPreview ? '/excel-to-csv/' : '/',
  plugins: [react(), tailwindcss()],
  // SheetJS + code pages for legacy .xls are ~1 MB; acceptable for a local tool.
  build: { chunkSizeWarningLimit: 1500 },
}))
