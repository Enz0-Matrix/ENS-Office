// SPDX-License-Identifier: AGPL-3.0-or-later
// Проверка Б3.1, часть 1 (ADR-013): структурированные ссылки в интерфейсе.
// Запуск: в app/table — npm run engine, npx vite --port 5173; затем node tests/structured-refs.mjs
// (нужен playwright: npm i --no-save playwright).
import { chromium } from 'playwright';
import { fileURLToPath } from 'node:url';

const URL = process.env.ENS_URL ?? 'http://localhost:5173/';
const FILE = fileURLToPath(new globalThis.URL('../../../core/ironcalc/xlsx/tests/calc_test_no_export/tables.xlsx', import.meta.url));
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1300, height: 800 } });
const errors = [];
page.on('pageerror', (e) => errors.push(e.message));
await page.goto(URL);
await page.waitForFunction(() => window.__ens?.model);

let fails = 0;
const ok = (cond, what) => { console.log(`${cond ? 'OK  ' : 'FAIL'} ${what}`); if (!cond) fails++; };

await page.setInputFiles('#file-input', FILE);
await page.waitForFunction(() => window.__ens.model().getWorksheetsProperties()[0].name === 'WithHeaders');
await page.evaluate(() => { window.__ens.grid.select(2, 6); });
await page.waitForTimeout(200);
const bar = await page.inputValue('#formula');
ok(bar === '=СУММ(Table1[@[Jan]:[Dec]])', `строка формул F2: ${bar}`);

// ввод вне таблицы — с именем таблицы
await page.evaluate(() => { window.__ens.grid.select(2, 8); document.getElementById('cell-editor').focus(); });
await page.keyboard.type('=Table1[@Jan]*10');
await page.keyboard.press('Enter');
await page.waitForTimeout(200);
const r = await page.evaluate(() => { const m = window.__ens.model(); return [m.getCellContent(0, 2, 8), m.getFormattedCellValue(0, 2, 8), m.getFormattedCellValue(0, 2, 1)]; });
ok(r[0] === '=Table1[@Jan]*10' && Number(r[1]) === Number(r[2]) * 10, `Table1[@Jan]*10 вне таблицы: ${r}`);

// переименование заголовка переписывает формулу
await page.evaluate(() => { window.__ens.grid.select(1, 1); document.getElementById('cell-editor').focus(); });
await page.keyboard.type('Январь');
await page.keyboard.press('Enter');
await page.waitForTimeout(200);
const f = await page.evaluate(() => window.__ens.model().getCellContent(0, 2, 8));
ok(f === '=Table1[@Январь]*10', `после переименования столбца: ${f}`);

// внутри таблицы — без имени: [@Январь] — своя таблица, эта строка
await page.evaluate(() => { window.__ens.grid.select(3, 6); document.getElementById('cell-editor').focus(); });
await page.keyboard.type('=[@Январь]+[@Feb]');
await page.keyboard.press('Enter');
await page.waitForTimeout(200);
const g = await page.evaluate(() => { const m = window.__ens.model(); return [m.getCellContent(0, 3, 6), Number(m.getFormattedCellValue(0, 3, 6)), Number(m.getFormattedCellValue(0, 3, 1)) + Number(m.getFormattedCellValue(0, 3, 2))]; });
ok(g[0] === '=Table1[@Январь]+Table1[@Feb]' && g[1] === g[2], `[@…] без имени таблицы: ${g}`);

console.log(errors.length ? `Ошибки страницы: ${errors.join('\n')}` : 'Ошибок страницы нет');
console.log(fails ? `ПРОВАЛОВ: ${fails}` : 'ВСЁ ПРОШЛО');
await browser.close();
process.exit(fails ? 1 : 0);
