// vite.config.js
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// IMPORTANTE: cambia 'taller-doe' por el nombre EXACTO de tu repo de GitHub
export default defineConfig({
  plugins: [react()],
  base: '/taller-doe/',
})
