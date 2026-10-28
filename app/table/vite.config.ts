import { defineConfig } from 'vite';

// WebView2 (Windows) и WebKitGTK (Linux) поддерживают ES2022, в том числе await
// на верхнем уровне модуля.
// ENS_SINGLE=1 — версия одним HTML-файлом (scripts/build-single.mjs):
// весь код одним модулем, файлы складываются в dist-single для вкладывания.
const single = process.env.ENS_SINGLE === '1';

export default defineConfig({
  base: single ? './' : '/',
  build: single
    ? {
      target: 'es2022', outDir: 'dist-single', emptyOutDir: true, cssCodeSplit: false, assetsInlineLimit: 0,
      rollupOptions: { output: { inlineDynamicImports: true } },
    }
    : { target: 'es2022' },
  clearScreen: false,
});
