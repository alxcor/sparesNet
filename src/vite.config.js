import { defineConfig } from 'vite';

export default defineConfig({
  base: './', // Important: makes paths relative so Electron can find them
  build: {
    outDir: 'dist',
    emptyOutDir: true,
  }
});
