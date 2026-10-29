// Проверка ENS-Tekst в Chromium (Playwright): ввод, форматирование, списки, таблицы, изображения,
// docx туда и обратно, восстановление, поиск. Запуск: npm test (Playwright должен быть доступен).
import { spawn, execSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { readFileSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');
let pw;
try {
  pw = createRequire(import.meta.url)('playwright');
} catch {
  pw = createRequire(join(execSync('npm root -g').toString().trim(), 'x.js'))('playwright');
}
const PORT = 5175;
const server = spawn(process.execPath, [join(root, 'node_modules/vite/bin/vite.js'), '--port', String(PORT), '--strictPort'], { cwd: root, stdio: 'pipe' });
await new Promise((res, rej) => {
  server.stdout.on('data', (d) => { if (String(d).includes('Local')) res(); });
  server.on('exit', () => rej(new Error('vite не запустился')));
  setTimeout(() => rej(new Error('vite: таймаут запуска')), 30000);
});

const browser = await pw.chromium.launch();
let failed = 0;
const errors = [];
async function page(url = `http://localhost:${PORT}/`) {
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
  const p = await ctx.newPage();
  p.on('pageerror', (e) => errors.push(e.message));
  await p.goto(url);
  await p.waitForSelector('.ProseMirror');
  await p.waitForFunction(() => window.ens);
  return p;
}
const json = (p) => p.evaluate(() => window.ens.view.state.doc.toJSON());
const text = (p) => p.evaluate(() => window.ens.view.state.doc.textBetween(0, window.ens.view.state.doc.content.size, '|'));
function eq(a, b, what) {
  const [x, y] = [JSON.stringify(a), JSON.stringify(b)];
  if (x === y) return;
  const i = [...x].findIndex((c, k) => c !== y[k]);
  const cut = (s) => s.slice(Math.max(0, i - 80), i + 120);
  throw new Error(`${what}: расхождение в позиции ${i}: ожидалось …${cut(y)}…, получено …${cut(x)}…`);
}
function ok(c, what) {
  if (!c) throw new Error(what);
}
async function test(name, f) {
  try {
    await f();
    console.log(`  ✓ ${name}`);
  } catch (e) {
    failed++;
    console.log(`  ✗ ${name}\n      ${e.message}`);
  }
}
const fixture = (n) => [...readFileSync(join(here, 'fixtures', n))];

console.log('ENS-Tekst');

await test('ввод по-русски, Enter, отмена и повтор', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Привет, мир!');
  await p.keyboard.press('Enter');
  await p.keyboard.type('Второй');
  eq(await text(p), 'Привет, мир!|Второй', 'текст');
  await p.keyboard.press('Control+KeyZ');
  await p.keyboard.press('Control+KeyZ');
  ok((await text(p)).length < 12, 'отмена');
  await p.keyboard.press('Control+KeyY');
  ok((await text(p)).includes('Привет'), 'повтор');
  await p.context().close();
});

await test('полужирный, курсив, шрифт и размер, выравнивание', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Текст');
  await p.keyboard.press('Control+KeyA');
  await p.keyboard.press('Control+KeyB');
  await p.keyboard.press('Control+KeyI');
  await p.selectOption('select[data-command="format.fontName"]', 'Arial');
  await p.selectOption('select[data-command="format.fontSize"]', '20');
  await p.keyboard.press('Control+KeyE');
  const d = await json(p);
  const para = d.content[0];
  eq(para.attrs.align, 'center', 'выравнивание');
  eq(para.content[0].marks.map((m) => m.type).sort(), ['em', 'strong', 'textStyle'], 'знаки');
  const ts = para.content[0].marks.find((m) => m.type === 'textStyle').attrs;
  eq([ts.font, ts.size], ['Arial', 20], 'шрифт');
  await p.context().close();
});

await test('заголовок: Enter в конце даёт обычный абзац', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.selectOption('select[data-command="para.style"]', '1');
  await p.keyboard.type('Глава');
  await p.keyboard.press('Enter');
  await p.keyboard.type('Текст');
  const d = await json(p);
  eq([d.content[0].type, d.content[0].attrs.level, d.content[1].type], ['heading', 1, 'paragraph'], 'типы');
  await p.context().close();
});

await test('списки: создать, Enter, Tab вкладывает, пустой пункт выходит', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.click('.btn[data-command="list.bullet"]');
  await p.keyboard.type('Один');
  await p.keyboard.press('Enter');
  await p.keyboard.type('Два');
  await p.keyboard.press('Tab');
  await p.keyboard.press('Enter');
  await p.keyboard.press('Enter');
  await p.keyboard.press('Enter');
  await p.keyboard.press('Enter');
  await p.keyboard.type('После');
  const d = await json(p);
  eq(d.content[0].type, 'bullet_list', 'список');
  ok(d.content[0].content.some((li) => li.content.some((c) => c.type === 'bullet_list')), 'вложенный список');
  eq(d.content[d.content.length - 1].type, 'paragraph', 'выход из списка');
  ok((await text(p)).endsWith('После'), 'текст после списка');
  await p.context().close();
});

await test('таблица: вставка, Tab между ячейками', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.click('.tab:has-text("Вставка")');
  await p.click('.btn[data-command="insert.table"]');
  await p.fill('#dialog input[type=number] >> nth=0', '2');
  await p.fill('#dialog input[type=number] >> nth=1', '2');
  await p.click('#dialog button.primary');
  await p.keyboard.type('А');
  await p.keyboard.press('Tab');
  await p.keyboard.type('Б');
  await p.keyboard.press('Tab');
  await p.keyboard.type('В');
  const d = await json(p);
  const table = d.content.find((n) => n.type === 'table');
  ok(table, 'таблица есть');
  eq(table.content.length, 2, 'строк');
  eq(table.content[0].content.length, 2, 'столбцов');
  ok(JSON.stringify(table.content[1]).includes('"В"'), 'текст в третьей ячейке (вторая строка)');
  await p.context().close();
});

await test('изображение: вставка из файла', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');
  await p.setInputFiles('#image-input', { name: 'точка.png', mimeType: 'image/png', buffer: png });
  await p.waitForFunction(() => JSON.stringify(window.ens.view.state.doc.toJSON()).includes('"image"'));
  ok(JSON.stringify(await json(p)).includes('data:image/png;base64'), 'данные картинки');
  await p.context().close();
});

await test('docx: образец открывается; сохранение и повторное открытие дают тот же документ', async () => {
  const p = await page();
  const res = await p.evaluate((arr) => {
    const E = window.ens;
    const a = E.importDocx(new Uint8Array(arr));
    E.loadDoc(a.doc, a.meta, 'sample');
    const bytes = E.exportDocx(E.view.state.doc, E.getMeta(), 'sample');
    const b = E.importDocx(bytes);
    return { a: a.doc.toJSON(), b: b.doc.toJSON(), pa: a.meta.page, pb: b.meta.page, lossy: a.meta.lossy };
  }, fixture('sample.docx'));
  eq(res.b, res.a, 'документ после круга');
  eq(res.pb, res.pa, 'страница после круга');
  eq(res.lossy, [], 'потерь нет');
  const s = JSON.stringify(res.a);
  for (const need of ['"heading"', '"strong"', '"em"', '"underline"', '"textStyle"', '"bullet_list"', '"ordered_list"', '"table"', '"colspan":2', '"rowspan":2', '"image"', '"link"', '"pageBreakBefore":true']) {
    ok(s.includes(need), `в документе нет ${need}`);
  }
  await p.context().close();
});

await test('docx: колонтитулы не переносятся — программа предупреждает', async () => {
  const p = await page();
  const lossy = await p.evaluate((arr) => window.ens.importDocx(new Uint8Array(arr)).meta.lossy, fixture('header.docx'));
  ok(lossy.some((x) => x.includes('колонтитулы')), `lossy: ${JSON.stringify(lossy)}`);
  await p.context().close();
});

await test('открытие через окно «Открыть»; не-docx — понятная ошибка, документ цел', async () => {
  const p = await page();
  await p.setInputFiles('#file-input', join(here, 'fixtures', 'header.docx'));
  await p.waitForSelector('#notice:not([hidden])');
  ok((await p.textContent('#notice')).includes('колонтитулы'), 'уведомление о потерях');
  ok((await text(p)).includes('Текст с колонтитулом'), 'текст открыт');
  await p.setInputFiles('#file-input', { name: 'письмо.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('это не zip') });
  await p.waitForSelector('#dialog:not([hidden])');
  ok((await p.textContent('#dialog')).includes('Не удалось открыть файл'), 'сообщение об ошибке');
  ok((await text(p)).includes('Текст с колонтитулом'), 'прежний документ не тронут');
  await p.context().close();
});

await test('сохранение: скачивается docx, который читается обратно', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Сохраняемый текст');
  await p.click('#qa-save');
  await p.waitForSelector('#dialog:not([hidden])');
  await p.fill('#dialog input[type=text]', 'report');
  const [dl] = await Promise.all([p.waitForEvent('download'), p.click('#dialog button.primary')]);
  eq(dl.suggestedFilename(), 'report.docx', 'имя файла');
  const path = join(mkdtempSync(join(tmpdir(), 'ens-')), 'out.docx');
  await dl.saveAs(path);
  const back = await p.evaluate((arr) => window.ens.importDocx(new Uint8Array(arr)).doc.textContent, [...readFileSync(path)]);
  eq(back, 'Сохраняемый текст', 'текст в файле');
  await p.waitForFunction(() => !window.ens.isDirty());
  await p.context().close();
});

await test('восстановление после закрытия без сохранения', async () => {
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  const p = await ctx.newPage();
  await p.goto(`http://localhost:${PORT}/`);
  await p.waitForSelector('.ProseMirror');
  await p.click('.ProseMirror');
  await p.keyboard.type('Важный текст, который нельзя терять');
  await p.waitForTimeout(2200);
  const p2 = await ctx.newPage();
  await p2.goto(`http://localhost:${PORT}/`);
  await p2.waitForSelector('#dialog:not([hidden])');
  ok((await p2.textContent('#dialog')).includes('Восстановление документов'), 'окно восстановления');
  await p2.click('#dialog button.primary');
  ok((await text(p2)).includes('Важный текст, который нельзя терять'), 'текст восстановлен');
  await ctx.close();
});

await test('поиск и замена всех', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('кот и кот, и ещё Кот');
  await p.keyboard.press('Control+KeyH');
  await p.fill('#findbar input[type=text] >> nth=0', 'кот');
  await p.fill('#findbar input[type=text] >> nth=1', 'пёс');
  await p.click('#findbar button:has-text("Заменить все")');
  eq(await text(p), 'пёс и пёс, и ещё пёс', 'замена без учёта регистра');
  await p.context().close();
});

await test('вставка HTML сохраняет форматирование и не пускает скрипты', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.evaluate(() => {
    const dt = new DataTransfer();
    dt.setData('text/html', '<p style="text-align:center"><b>Жирный</b> и <i>курсив</i><script>window.__x=1</script></p><ul><li>раз</li><li>два</li></ul>');
    dt.setData('text/plain', 'Жирный и курсив');
    document.querySelector('.ProseMirror').dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  });
  const d = JSON.stringify(await json(p));
  ok(d.includes('"strong"') && d.includes('"em"') && d.includes('bullet_list'), 'форматирование вставлено');
  ok(!d.includes('script') && !(await p.evaluate(() => window.__x)), 'скрипт не попал в документ');
  await p.context().close();
});

await test('страница: поля и ориентация меняют вид и сохраняются в файле', async () => {
  const p = await page();
  await p.click('.tab:has-text("Макет")');
  await p.selectOption('select[data-command="page.margins"]', 'narrow');
  await p.selectOption('select[data-command="page.orientation"]', 'landscape');
  const m = await p.evaluate(() => window.ens.getMeta().page);
  ok(m.w > m.h, 'альбомная');
  ok(Math.abs(m.left - 36) < 1, `поле 1,27 см, а не ${m.left}`);
  const back = await p.evaluate(() => {
    const E = window.ens;
    return E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta())).meta.page;
  });
  eq(back.w > back.h, true, 'ориентация в файле');
  await p.context().close();
});

await test('разрыв страницы: Ctrl+Enter', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('До');
  await p.keyboard.press('Control+Enter');
  await p.keyboard.type('После');
  const d = await json(p);
  eq(d.content.map((n) => n.attrs.pageBreakBefore), [false, true], 'разрыв перед вторым абзацем');
  await p.context().close();
});

ok(errors.length === 0, `ошибки страницы: ${errors.join('; ')}`);
await browser.close();
server.kill();
console.log(failed ? `\nПровалено: ${failed}` : '\nВсе проверки пройдены');
process.exit(failed ? 1 : 0);
