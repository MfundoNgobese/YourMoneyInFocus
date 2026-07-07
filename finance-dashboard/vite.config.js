import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// base './' keeps asset paths relative, so the same build works on
// Vercel, Netlify, and GitHub Pages (even under a sub-path) without changes.
export default defineConfig({
  base: './',
  plugins: [react()],
})
