import { defineConfig } from 'vite';

// WebView2 (Windows) и WebKitGTK (Linux) поддерживают ES2022, в том числе await
// на верхнем уровне модуля.
export default defineConfig({
  build: { target: 'es2022' },
  clearScreen: false,
});
