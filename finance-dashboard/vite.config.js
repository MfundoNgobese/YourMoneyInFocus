import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Relative base so the same build runs on Vercel, Netlify, and GitHub Pages.
export default defineConfig({
  base: './',
  plugins: [react()],
})
