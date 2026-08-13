// Timestamp: 2025-11-22 | Version: 2.0.2
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';

export default defineConfig({
  // FIX: Set the base path to correctly resolve assets when deployed to the 
  // GitHub Pages subfolder (https://Northbridge-University.github.io/Student-Retention-Add-in/react/dist/).
  // Updated to include 'dist/' because the production files are served from that subdirectory.
  base: '/Student-Retention-Add-in/react/dist/', 
  plugins: [react()],
  css: {
    // Note: path.resolve(__dirname, 'postcss.config.cjs') might not work 
    // depending on the execution environment; './postcss.config.cjs' might be safer.
    postcss: path.resolve(__dirname, 'postcss.config.cjs'),
  },
});