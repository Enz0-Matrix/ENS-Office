import { defineConfig } from 'vite';

// Общие модули (лента, команды, шрифты, значки) лежат в ../table/src.
export default defineConfig({
  base: '/',
  build: { target: 'es2022' },
  server: { fs: { allow: ['..'] } },
  clearScreen: false,
});
