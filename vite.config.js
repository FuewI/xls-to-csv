import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// GitHub Pages serves the site from /<repository-name>/.
// The deploy workflow sets BASE_PATH automatically (e.g. "/xls-to-csv/");
// locally it falls back to "/" so `npm run dev` keeps working unchanged.
const base = process.env.BASE_PATH || '/'

export default defineConfig({
  base: base.endsWith('/') ? base : `${base}/`,
  plugins: [react(), tailwindcss()],
  // SheetJS + code pages for legacy .xls are ~1 MB; acceptable for a local tool.
  build: { chunkSizeWarningLimit: 1500 },
})
