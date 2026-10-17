// Запуск обоих прототипов в headless-Chrome (без фонового троттлинга) и сводка результатов.
// Использование: node run-bench.mjs 1000 10000 50000
import puppeteer from './univer-build/node_modules/puppeteer-core/lib/esm/puppeteer/puppeteer-core.js';
const BROWSER = process.env.BENCH_BROWSER || 'C:/Program Files/Google/Chrome/Application/chrome.exe';
const LOOKUP = process.env.BENCH_LOOKUP ?? '1';
const sizes = process.argv.slice(2).map(Number); if (!sizes.length) sizes.push(1000);
const browser = await puppeteer.launch({ executablePath: BROWSER, headless: true, args: ['--enable-precise-memory-info'], protocolTimeout: 1800000 });
const results = [];
for (const rows of sizes) for (const page of ['ironcalc', 'univer']) {
  const p = await browser.newPage(); await p.setViewport({ width: 1280, height: 800 });
  p.on('pageerror', e => console.error(page, 'pageerror', e.message));
  const t = Date.now();
  await p.goto(`http://localhost:8731/${page}.html?rows=${rows}&lookup=${LOOKUP}`);
  try { await p.waitForFunction('window.__bench', { timeout: 1500000, polling: 500 }); results.push({ page, ...(await p.evaluate('window.__bench')) }); }
  catch (e) { results.push({ page, rows, error: e.message, report: await p.evaluate("document.getElementById('report').textContent") }); }
  console.error(`${page} ${rows}: ${(Date.now() - t) / 1000}s`);
  await p.close();
}
await browser.close();
console.log(JSON.stringify(results, null, 1));
