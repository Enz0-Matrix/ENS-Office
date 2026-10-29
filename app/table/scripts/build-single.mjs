// ENS-Table одним HTML-файлом: для компьютеров, где настольная версия не
// запускается (старые системы, например РЕД ОС 7 с glibc 2.28). Файл
// открывается двойным щелчком в Firefox/Chromium/Яндекс.Браузере, без
// интернета и без установки. Внутри: код, стили, шрифты и движок (base64).
//
//   ENS_SINGLE=1 vite build && node scripts/build-single.mjs [версия]

import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { execFileSync } from 'node:child_process';

const dir = 'dist-single';
const version = process.argv[2] ?? JSON.parse(readFileSync('package.json', 'utf8')).version;
let html = readFileSync(join(dir, 'index.html'), 'utf8');

const assetPath = (url) => join(dir, url.replace(/^\.?\//, ''));

// Стили: шрифты вкладываем только woff2 с кириллицей и латиницей —
// остальные наборы (греческий, вьетнамский…) и запасные woff не нужны.
const cssMatch = html.match(/<link rel="stylesheet"[^>]*href="([^"]+)"[^>]*>/);
let css = cssMatch ? readFileSync(assetPath(cssMatch[1]), 'utf8') : '';
const KEEP = /(cyrillic|latin)(-ext)?-\d{3}-(normal|italic)/;
css = css.replace(/@font-face\s*{[^}]*}/g, (block) => {
  const woff2 = block.match(/url\(([^)]+\.woff2)\)/);
  if (!woff2 || !KEEP.test(woff2[1])) return '';
  const file = woff2[1].replace(/^\.\//, '').replace(/^\/?assets\//, '');
  const path = join(dir, 'assets', file);
  if (!existsSync(path)) return '';
  const data = readFileSync(path).toString('base64');
  return block.replace(/src:[^;}]*/, `src:url(data:font/woff2;base64,${data}) format("woff2")`);
});

// Код: один модуль. Движок — base64 в отдельном скрипте до модуля.
const jsMatch = html.match(/<script type="module"[^>]*src="([^"]+)"[^>]*><\/script>/);
if (!jsMatch) throw new Error('не найден модуль в index.html');
const js = readFileSync(assetPath(jsMatch[1]), 'utf8').replace(/<\/script/gi, '<\\/script');
const wasm = readFileSync('src/engine-pkg/engine_bg.wasm').toString('base64');

html = html
  // Замена функцией, а не строкой: в коде бывают «$'», «$&» — строка их подставляла бы особо.
  .replace(cssMatch ? cssMatch[0] : '', () => `<style>${css}</style>`)
  .replace(jsMatch[0], () => `<script>window.__ENS_WASM_B64__="${wasm}";</script>\n<script type="module">${js}</script>`)
  .replace('<title>', () => `<meta name="generator" content="ENS-Table ${version} (один файл)">\n  <title>`);

// Страховка: код в файле не должен быть искажён при сборке (раньше «$'» в коде ломал страницу).
const embedded = html.slice(html.indexOf('<script type="module">') + 22, html.lastIndexOf('</script>'));
writeFileSync(join(dir, 'check.mjs'), embedded);
try { execFileSync(process.execPath, ['--check', join(dir, 'check.mjs')], { stdio: 'pipe' }); }
catch (e) { throw new Error(`код в одном файле искажён: ${e.stderr?.toString().split('\n')[0] ?? e}`); }

const out = join(dir, `ENS-Table_${version}_browser.html`);
writeFileSync(out, html);
console.log(`Готово: ${out}, ${(html.length / 1048576).toFixed(1)} МБ`);
