// SPDX-License-Identifier: AGPL-3.0-or-later
// Проверка блока 2 (Б2.1–Б2.3) в двух окнах книг одного браузера.
// Запуск: в app/table — npm run engine, npx vite --port 5173; затем node tests/sheets-between-books.mjs
// (нужен playwright: npm i --no-save playwright).
import { chromium } from 'playwright';

const URL = process.env.ENS_URL ?? 'http://localhost:5173/';
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1300, height: 800 } });
const A = await ctx.newPage();
const B = await ctx.newPage();
const errors = [];
for (const [n, p] of [['A', A], ['B', B]]) p.on('pageerror', (e) => errors.push(`${n}: ${e.message}`));
await A.goto(URL); await B.goto(URL);
await A.waitForFunction(() => window.__ens?.model);
await B.waitForFunction(() => window.__ens?.model);

let fails = 0;
const ok = (cond, what) => { console.log(`${cond ? 'OK  ' : 'FAIL'} ${what}`); if (!cond) fails++; };
const m = (p, f, ...a) => p.evaluate(([f, a]) => new Function('m', 'a', `return (${f})(m, ...a)`)(window.__ens.model(), a), [f.toString(), a]);
const tabs = (p) => p.$$eval('.sheet-tab', (els) => els.map((e) => ({ t: e.textContent, g: e.classList.contains('grouped'), a: e.classList.contains('active') })));
const menuClick = async (p, text) => { await p.locator('#context-menu button, #context-submenu button', { hasText: text }).first().click(); };

// --- Б2.1: группа ярлыков ---
await A.click('.sheet-add'); await A.click('.sheet-add');
await A.locator('.sheet-tab', { hasText: 'Лист1' }).click();
await A.locator('.sheet-tab', { hasText: 'Лист2' }).click({ modifiers: ['Control'] });
let t = await tabs(A);
ok(t.filter((x) => x.g).length === 2, `Ctrl+щелчок: в группе 2 листа ${JSON.stringify(t)}`);
await A.locator('.sheet-tab', { hasText: 'Лист3' }).click({ modifiers: ['Shift'] });
t = await tabs(A);
ok(t.filter((x) => x.g).length === 3, 'Shift+щелчок: в группе 3 листа');
await A.locator('.sheet-tab', { hasText: 'Лист3' }).click({ modifiers: ['Control'] });
await A.locator('.sheet-tab', { hasText: 'Лист2' }).click({ button: 'right' });
const labels = await A.$$eval('#context-menu button', (b) => b.map((x) => x.textContent));
ok(labels.some((l) => l.includes('Скрыть листы (2)')) && labels.some((l) => l.includes('Удалить листы (2)')), `меню группы: ${labels.join(' | ')}`);
await A.locator('#context-menu .swatch').nth(1).click();
let colors = await m(A, (m) => m.getWorksheetsProperties().map((p) => p.color ?? ''));
ok(colors[0] && colors[1] && !colors[2], `цвет на двух ярлыках: ${colors}`);
await A.keyboard.press('Control+z'); await A.waitForTimeout(150);
colors = await m(A, (m) => m.getWorksheetsProperties().map((p) => p.color ?? ''));
ok(!colors[0] && !colors[1], `одна отмена снимает цвет с обоих: ${colors}`);
// скрыть группу и вернуть одной отменой
await A.locator('.sheet-tab', { hasText: 'Лист1' }).click();
await A.locator('.sheet-tab', { hasText: 'Лист2' }).click({ modifiers: ['Control'] });
await A.locator('.sheet-tab', { hasText: 'Лист2' }).click({ button: 'right' });
await menuClick(A, 'Скрыть листы (2)');
let st = await m(A, (m) => m.getWorksheetsProperties().map((p) => p.state));
ok(st.join() === 'hidden,hidden,visible', `скрыты два листа: ${st}`);
await A.keyboard.press('Control+z'); await A.waitForTimeout(150);
st = await m(A, (m) => m.getWorksheetsProperties().map((p) => p.state));
ok(st.join() === 'visible,visible,visible', `одна отмена показывает оба: ${st}`);
// удалить группу и вернуть
await A.evaluate(() => { const { grid } = window.__ens; });
await A.locator('.sheet-tab', { hasText: 'Лист3' }).click();
await A.locator('.sheet-tab', { hasText: 'Лист3' }).click();
await A.locator('.sheet-tab', { hasText: 'Лист2' }).click({ modifiers: ['Control'] });
A.once('dialog', (d) => d.accept());
await A.locator('.sheet-tab', { hasText: 'Лист2' }).click({ button: 'right' });
await menuClick(A, 'Удалить листы (2)');
let names = await m(A, (m) => m.getWorksheetsProperties().map((p) => p.name));
ok(names.join() === 'Лист1', `удалены два листа: ${names}`);
await A.keyboard.press('Control+z'); await A.waitForTimeout(150);
names = await m(A, (m) => m.getWorksheetsProperties().map((p) => p.name));
ok(names.join() === 'Лист1,Лист2,Лист3', `одна отмена возвращает оба: ${names}`);
await A.evaluate(() => window.__ens.grid.render());

// --- Б2.2: перенос листа в другую открытую книгу ---
await m(A, (m) => { m.setUserInput(0, 1, 1, '5'); m.setUserInput(1, 1, 1, '=Лист1!A1*2'); m.setUserInput(1, 2, 1, 'текст'); m.updateRangeStyle({ sheet: 1, row: 1, column: 1, width: 1, height: 1 }, 'font.b', 'true'); });
await A.evaluate(() => window.__ens.grid.render());
await A.locator('.sheet-tab', { hasText: 'Лист2' }).click({ button: 'right' });
await menuClick(A, 'Переместить или скопировать');
await A.waitForSelector('#recovery [data-f="book"]');
const opts = await A.$$eval('#recovery [data-f="book"] option', (o) => o.map((x) => x.textContent));
ok(opts.length === 3, `в списке книг: ${opts.join(' | ')}`);
await A.selectOption('#recovery [data-f="book"]', { index: 1 });
await A.selectOption('#recovery [data-f="before"]', '-1');
await A.click('#recovery [data-act="ok"]');
await A.waitForTimeout(800);
names = await m(A, (m) => m.getWorksheetsProperties().map((p) => p.name));
ok(names.join() === 'Лист1,Лист3', `в книге A лист ушёл: ${names}`);
let bInfo = await m(B, (m) => ({ names: m.getWorksheetsProperties().filter((p) => p.state === 'visible').map((p) => p.name), v: m.getFormattedCellValue(1, 1, 1), f: m.getCellContent(1, 1, 1), t: m.getFormattedCellValue(1, 2, 1), b: m.getCellStyle(1, 1, 1).style.font.b }));
ok(bInfo.names.join() === 'Лист1,Лист2' && bInfo.v === '10' && bInfo.t === 'текст' && bInfo.b, `в книге B: ${JSON.stringify(bInfo)}`);
ok(bInfo.f === "='[Книга1]Лист1'!A1*2", `ссылка стала внешней: ${bInfo.f}`);
await A.keyboard.press('Control+z'); await A.waitForTimeout(150);
names = await m(A, (m) => m.getWorksheetsProperties().map((p) => p.name));
ok(names.join() === 'Лист1,Лист2,Лист3', `Ctrl+Z в A возвращает лист: ${names}`);

// копия в новую книгу
await A.evaluate(() => window.__ens.grid.render());
await A.locator('.sheet-tab', { hasText: 'Лист2' }).click({ button: 'right' });
await menuClick(A, 'Переместить или скопировать');
await A.waitForSelector('#recovery [data-f="book"]');
await A.selectOption('#recovery [data-f="book"]', 'new');
await A.check('#recovery [data-f="copy"]');
const [C] = await Promise.all([ctx.waitForEvent('page'), A.click('#recovery [data-act="ok"]')]);
C.on('pageerror', (e) => errors.push(`C: ${e.message}`));
await C.waitForFunction(() => window.__ens?.model);
await C.waitForTimeout(1500);
const cInfo = await m(C, (m) => ({ names: m.getWorksheetsProperties().filter((p) => p.state === 'visible').map((p) => p.name), v: m.getFormattedCellValue(0, 1, 1) }));
ok(cInfo.names.join() === 'Лист2' && cInfo.v === '10', `новая книга: только перенесённый лист ${JSON.stringify(cInfo)}`);
names = await m(A, (m) => m.getWorksheetsProperties().map((p) => p.name));
ok(names.join() === 'Лист1,Лист2,Лист3', `копия: в A лист остался: ${names}`);
await C.close();

// --- Б2.3: Вставить связь из другой книги ---
await m(A, (m) => { m.setUserInput(0, 2, 2, '42'); m.setUserInput(0, 3, 2, '7'); });
await A.evaluate(() => { window.__ens.grid.render(); });
await A.locator('.sheet-tab', { hasText: 'Лист1' }).click();
await A.evaluate(() => window.__ens.grid.selectRange(2, 2, 3, 2));
await A.keyboard.press('Control+c');
await B.bringToFront();
await B.locator('.sheet-tab', { hasText: 'Лист1' }).click();
await B.evaluate(() => window.__ens.grid.select(5, 3));
await B.keyboard.press('Control+Alt+v');
const psLabels = await B.$$eval('#context-menu button', (b) => b.map((x) => x.textContent));
ok(psLabels.some((l) => l.includes('Вставить связь')), `спецвставка в B: ${psLabels.join(' | ')}`);
await menuClick(B, 'Вставить связь');
await B.waitForTimeout(800);
const link = await m(B, (m) => [m.getCellContent(0, 5, 3), m.getFormattedCellValue(0, 5, 3), m.getCellContent(0, 6, 3), m.getFormattedCellValue(0, 6, 3)]);
ok(link[0] === "='[Книга1]Лист1'!B2" && link[1] === '42' && link[3] === '7', `связь из A: ${link}`);
await B.keyboard.press('Control+z'); await B.waitForTimeout(150);
const undone = await m(B, (m) => [m.getCellContent(0, 5, 3), m.getCellContent(0, 6, 3)]);
ok(!undone[0] && !undone[1], `одна отмена убирает обе связи: ${undone}`);

// своя книга: связь на другой лист
await A.bringToFront();
await A.evaluate(() => window.__ens.grid.select(2, 2));
await A.keyboard.press('Control+c');
await A.locator('.sheet-tab', { hasText: 'Лист3' }).click();
await A.evaluate(() => window.__ens.grid.select(1, 1));
await A.keyboard.press('Control+Alt+v');
await menuClick(A, 'Вставить связь');
await A.waitForTimeout(300);
const own = await m(A, (m) => [m.getCellContent(2, 1, 1), m.getFormattedCellValue(2, 1, 1)]);
ok(own[0] === '=Лист1!$B$2' && own[1] === '42', `связь в своей книге: ${own}`);

console.log(errors.length ? `Ошибки страниц: ${errors.join('\n')}` : 'Ошибок страниц нет');
console.log(fails ? `ПРОВАЛОВ: ${fails}` : 'ВСЁ ПРОШЛО');
await browser.close();
process.exit(fails ? 1 : 0);
