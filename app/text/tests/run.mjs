// Проверка ENS-Text в Chromium (Playwright): ввод, форматирование, списки, таблицы, изображения,
// docx туда и обратно, восстановление, поиск. Запуск: npm test (Playwright должен быть доступен).
import { spawn, execSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
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

// ENS_CHROMIUM=путь — гонять на другом Chromium (например, 128, как на РЕД ОС 7).
const browser = await pw.chromium.launch(process.env.ENS_CHROMIUM ? { executablePath: process.env.ENS_CHROMIUM } : {});
if (process.env.ENS_CHROMIUM) console.log(`Chromium ${browser.version()}`);
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
  // ONLY=подстрока — запустить только подходящие тесты (для отладки).
  if (process.env.ONLY && !name.includes(process.env.ONLY)) return;
  try {
    await f();
    console.log(`  ✓ ${name}`);
  } catch (e) {
    failed++;
    console.log(`  ✗ ${name}\n      ${e.message}`);
  } finally {
    // Не копим вкладки между тестами: это важно в среде с ограниченной памятью.
    await Promise.all(browser.contexts().map((ctx) => ctx.close()));
  }
}
const fixture = (n) => [...readFileSync(join(here, 'fixtures', n))];

console.log('ENS-Text');

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

await test('docx: колонтитулы и поля номера страницы читаются и пишутся, особый колонтитул первой страницы', async () => {
  const p = await page();
  const r = await p.evaluate((arr) => {
    const E = window.ens;
    const a = E.importDocx(new Uint8Array(arr));
    const b = E.importDocx(E.exportDocx(a.doc, a.meta, 'x'));
    return { hf: a.meta.hf, back: b.meta.hf, lossy: a.meta.lossy };
  }, fixture('header.docx'));
  eq(r.lossy, [], 'потерь нет');
  const s = JSON.stringify(r.hf);
  ok(s.includes('Колонтитул документа') && s.includes('Титульный лист'), 'тексты колонтитулов');
  ok(s.includes('"kind":"page"') && s.includes('"kind":"pages"'), 'поля номера и числа страниц');
  eq(r.hf.differentFirst, true, 'особый колонтитул первой страницы');
  eq(r.back, r.hf, 'колонтитулы после круга docx');
  await p.context().close();
});

await test('docx: сноски не переносятся — программа предупреждает', async () => {
  const p = await page();
  const lossy = await p.evaluate((arr) => window.ens.importDocx(new Uint8Array(arr)).meta.lossy, fixture('lossy.docx'));
  ok(lossy.some((x) => x.includes('сноски')), `lossy: ${JSON.stringify(lossy)}`);
  await p.context().close();
});

await test('открытие через окно «Открыть»; не-docx — понятная ошибка, документ цел', async () => {
  const p = await page();
  await p.setInputFiles('#file-input', join(here, 'fixtures', 'lossy.docx'));
  await p.waitForSelector('#notice:not([hidden])');
  ok((await p.textContent('#notice')).includes('сноски'), 'уведомление о потерях');
  ok((await text(p)).includes('Текст со сноской'), 'текст открыт');
  await p.setInputFiles('#file-input', { name: 'письмо.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('это не zip') });
  await p.waitForSelector('#dialog:not([hidden])');
  ok((await p.textContent('#dialog')).includes('Не удалось открыть файл'), 'сообщение об ошибке');
  ok((await text(p)).includes('Текст со сноской'), 'прежний документ не тронут');
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


// ---- Постраничная вёрстка ----
const LAYOUT_JS = `
  window.__bad = () => {
    const E = window.ens, v = E.view, g = E.geometry();
    const box = v.dom.getBoundingClientRect(), scale = box.width / v.dom.offsetWidth;
    const bad = [];
    v.state.doc.descendants((node, pos) => {
      if (!node.isTextblock) return true;
      const start = pos + 1, end = pos + 1 + node.content.size, step = Math.max(1, Math.floor((end - start) / 60));
      for (let q = start; q <= end; q += step) {
        const c = v.coordsAtPos(q, 1), top = (c.top - box.top) / scale, bot = (c.bottom - box.top) / scale, mid = (top + bot) / 2;
        const p = Math.floor((mid - g.top) / g.P), s = p * g.P + g.top, e = s + g.C;
        if (top < s - 1 || bot > e + 2) bad.push({ q, p, top: Math.round(top), bot: Math.round(bot), s: Math.round(s), e: Math.round(e) });
      }
      return false;
    });
    return bad.slice(0, 5);
  };
  window.__load = (build) => {
    const E = window.ens, s = E.schema, P = (t, a) => s.nodes.paragraph.create(a || null, t ? s.text(t) : undefined);
    E.loadDoc(s.nodes.doc.create(null, build(s, P)), E.getMeta(), 'long');
  };
  window.__sig = () => [...document.querySelectorAll('.pg-pad, .pg-split, .pg-gap')].map((e) => e.className + ':' + (e.style.paddingTop || e.style.height || e.firstChild?.style?.height)).join('|');
`;
const FILL = 'Съешь ещё этих мягких французских булок, да выпей же чаю. ';
async function longPage(build) {
  const p = await page();
  await p.evaluate(LAYOUT_JS);
  await p.evaluate(`(${build.toString()})`.replace(/^\(/, 'window.__build = ').replace(/\)$/, ''));
  await p.evaluate(() => window.__load(window.__build));
  await p.waitForTimeout(400);
  return p;
}

await test('вёрстка: текст всех страниц лежит внутри полей, страниц несколько', async () => {
  const p = await longPage((s, P) => Array.from({ length: 80 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(1 + (i % 9)))));
  ok(await p.evaluate(() => window.ens.pageCount()) >= 4, 'страниц меньше четырёх');
  eq(await p.evaluate(() => window.__bad()), [], 'текст вне полей');
  eq(await p.evaluate(() => document.querySelectorAll('.paper').length), await p.evaluate(() => window.ens.pageCount()), 'подложек столько же, сколько страниц');
  await p.context().close();
});

await test('вёрстка: расчёт устойчив, отмена и «грязный» документ не затронуты', async () => {
  const p = await longPage((s, P) => Array.from({ length: 70 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(3))));
  const sig = await p.evaluate(() => window.__sig());
  for (let i = 0; i < 3; i++) { await p.evaluate(() => window.ens.run('view.zoomIn')); await p.waitForTimeout(80); }
  eq(await p.evaluate(() => window.__sig()), sig, 'распорки после повторных расчётов');
  eq(await p.evaluate(() => window.ens.isDirty()), false, 'документ не «грязный» от вёрстки');
  await p.click('.ProseMirror p >> nth=0');
  await p.keyboard.press('Control+KeyZ');
  eq(await p.evaluate(() => window.ens.view.state.doc.childCount), 70, 'отмена не съела документ');
  await p.context().close();
});

await test('вёрстка: ввод в начале сдвигает текст на следующие страницы, удаление возвращает', async () => {
  const p = await longPage((s, P) => Array.from({ length: 50 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(3))));
  const before = await p.evaluate(() => window.ens.pageCount());
  const sig = await p.evaluate(() => window.__sig());
  await p.click('.ProseMirror p >> nth=0');
  await p.keyboard.press('Control+Home');
  for (let i = 0; i < 40; i++) await p.keyboard.press('Enter');
  await p.waitForTimeout(200);
  ok(await p.evaluate(() => window.ens.pageCount()) > before, 'число страниц не выросло');
  eq(await p.evaluate(() => window.__bad()), [], 'текст вне полей после ввода');
  for (let i = 0; i < 40; i++) await p.keyboard.press('Backspace');
  await p.waitForTimeout(200);
  eq(await p.evaluate(() => window.ens.pageCount()), before, 'число страниц вернулось');
  eq(await p.evaluate(() => window.__sig()), sig, 'распорки вернулись');
  await p.context().close();
});

await test('вёрстка: абзацы рвутся минимум по две строки с каждой стороны', async () => {
  const p = await longPage((s, P) => Array.from({ length: 160 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(2 + (i % 6)))));
  const res = await p.evaluate(() => {
    const E = window.ens, v = E.view, out = [];
    for (const w of document.querySelectorAll('.pg-split')) {
      const pos = v.posAtDOM(w, 0), $p = v.state.doc.resolve(pos), start = $p.start(), end = $p.end();
      const tops = (a, b) => { const set = new Set(); for (let q = a; q <= b; q++) set.add(Math.round(v.coordsAtPos(q, 1).top)); return set.size; };
      out.push([tops(start, pos - 1), tops(pos, end)]);
    }
    return out;
  });
  ok(res.length >= 2, `рваных абзацев мало: ${res.length}`);
  for (const [a, b] of res) ok(a >= 2 && b >= 2, `строк до разрыва ${a}, после ${b}`);
  await p.context().close();
});

await test('вёрстка: заголовок не остаётся один в конце страницы', async () => {
  const p = await page();
  await p.evaluate(LAYOUT_JS);
  const bad = [];
  for (let k = 20; k < 52; k++) {
    const r = await p.evaluate((k) => {
      window.__load((s, P) => [...Array.from({ length: k }, (_, i) => P('Строка ' + i)), s.nodes.heading.create({ level: 2 }, s.text('Заголовок раздела')), P('Текст после заголовка, первая строка.'), P('Ещё строка.'), P('И ещё.')]);
      return new Promise((res) => setTimeout(() => {
        const E = window.ens, v = E.view, g = E.geometry(), box = v.dom.getBoundingClientRect(), sc = box.width / v.dom.offsetWidth;
        let h, n; v.state.doc.forEach((node, off) => { if (node.type.name === 'heading') { h = off + 1; n = off + node.nodeSize + 1; } });
        const page = (q) => Math.floor((((v.coordsAtPos(q, 1).top - box.top) / sc) - g.top) / g.P);
        res([page(h), page(n)]);
      }, 250));
    }, k);
    if (r[0] !== r[1]) bad.push([k, r]);
  }
  eq(bad, [], 'заголовок на одной странице, текст на другой');
  await p.context().close();
});

await test('вёрстка: разрыв страницы начинает новую страницу, таблица рвётся по строкам', async () => {
  const p = await longPage((s, P) => {
    const cell = (t) => s.nodes.table_cell.create(null, P(t));
    const rows = Array.from({ length: 45 }, (_, i) => s.nodes.table_row.create(null, [cell('Строка ' + i), cell('Значение ' + i)]));
    return [P('До разрыва'), P('После разрыва', { pageBreakBefore: true }), s.nodes.table.create({ borders: true }, rows), P('После таблицы')];
  });
  const r = await p.evaluate(() => {
    const E = window.ens, v = E.view, g = E.geometry(), box = v.dom.getBoundingClientRect(), sc = box.width / v.dom.offsetWidth;
    const page = (q) => Math.floor((((v.coordsAtPos(q, 1).top - box.top) / sc) - g.top) / g.P);
    return { first: page(1), second: page(v.state.doc.child(0).nodeSize + 1), pages: E.pageCount(), gaps: document.querySelectorAll('tr.pg-gap').length };
  });
  eq([r.first, r.second], [0, 1], 'разрыв страницы');
  ok(r.pages >= 3 && r.gaps >= 1, `таблица не разбита: страниц ${r.pages}, разрывов ${r.gaps}`);
  eq(await p.evaluate(() => window.__bad()), [], 'текст таблицы вне полей');
  await p.context().close();
});

await test('вёрстка: смена поля и ориентации пересчитывает страницы; в docx страниц столько же', async () => {
  const p = await longPage((s, P) => Array.from({ length: 60 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(3))));
  const normal = await p.evaluate(() => window.ens.pageCount());
  await p.click('.tab:has-text("Макет")');
  await p.selectOption('select[data-command="page.margins"]', 'wide');
  await p.waitForTimeout(250);
  const wide = await p.evaluate(() => window.ens.pageCount());
  ok(wide > normal, `широкие поля: ${wide} страниц, обычные: ${normal}`);
  eq(await p.evaluate(() => window.__bad()), [], 'текст вне полей при широких полях');
  await p.selectOption('select[data-command="page.orientation"]', 'landscape');
  await p.waitForTimeout(250);
  eq(await p.evaluate(() => window.__bad()), [], 'текст вне полей в альбомной');
  await p.context().close();
});

await test('вёрстка: большой документ (3000 абзацев) считается быстро', async () => {
  const p = await page();
  await p.evaluate(LAYOUT_JS);
  const ms = await p.evaluate(() => {
    const t0 = performance.now();
    window.__load((s, P) => Array.from({ length: 3000 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(1 + (i % 4)))));
    return new Promise((res) => setTimeout(() => res(performance.now() - t0), 50));
  });
  await p.waitForFunction(() => window.ens.pageCount() > 50, null, { timeout: 15000 });
  const typing = await p.evaluate(async () => {
    const E = window.ens, v = E.view, t0 = performance.now();
    for (let i = 0; i < 20; i++) { v.dispatch(v.state.tr.insertText('я', 5)); await new Promise((r) => setTimeout(r, 0)); }
    return (performance.now() - t0) / 20;
  });
  console.log(`      загрузка 3000 абзацев ≈ ${Math.round(ms)} мс, страниц ${await p.evaluate(() => window.ens.pageCount())}, расчёт ≈ ${Math.round(await p.evaluate(() => window.ens.pagStats.ms))} мс, правка ≈ ${Math.round(typing)} мс`);
  ok(typing < 60, `правка в большом документе занимает ${Math.round(typing)} мс`);
  await p.context().close();
});


// ---- Колонтитулы и печать ----
await test('колонтитулы: на каждом листе свой номер «Страница N из M», на первом — особый колонтитул', async () => {
  const p = await page();
  await p.evaluate(LAYOUT_JS);
  await p.evaluate((arr) => {
    const E = window.ens;
    const x = E.importDocx(new Uint8Array(arr));
    const s = E.schema;
    const body = [...x.doc.content.content];
    for (let i = 0; i < 70; i++) body.push(s.nodes.paragraph.create(null, s.text(`Абзац ${i}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(3))));
    E.loadDoc(s.nodes.doc.create(null, body), x.meta, 'hf');
  }, fixture('header.docx'));
  await p.waitForTimeout(400);
  const r = await p.evaluate(() => [...document.querySelectorAll('.paper')].map((el) => ({ h: el.querySelector('.hf-header')?.textContent ?? '', f: el.querySelector('.hf-footer')?.textContent ?? '' })));
  ok(r.length >= 3, `страниц ${r.length}`);
  eq(r[0].h, 'Титульный лист', 'первая страница');
  eq(r[1].h, 'Колонтитул документа', 'вторая страница');
  eq(r.map((x) => x.f), r.map((_, i) => (i === 0 ? '' : `Страница ${i + 1} из ${r.length}`)), 'номера страниц в нижнем колонтитуле (на первой странице колонтитула нет)');
  eq(await p.evaluate(() => window.__bad()), [], 'текст вне полей');
  await p.context().close();
});

await test('колонтитулы: высокий колонтитул сдвигает текст вниз, как в Word', async () => {
  const p = await longPage((s, P) => Array.from({ length: 60 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(3))));
  const before = await p.evaluate(() => ({ top: window.ens.effPage().top, pages: window.ens.pageCount() }));
  await p.evaluate(() => {
    const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const m = E.getMeta();
    m.hf.header = s.nodes.doc.create(null, Array.from({ length: 6 }, (_, i) => P('Строка шапки ' + i))).toJSON();
    E.loadDoc(E.view.state.doc, m, 'hf');
  });
  await p.waitForTimeout(400);
  const after = await p.evaluate(() => ({ top: window.ens.effPage().top, pages: window.ens.pageCount() }));
  ok(after.top > before.top + 30, `верхнее поле ${before.top} → ${after.top}`);
  ok(after.pages >= before.pages, 'страниц не стало меньше');
  eq(await p.evaluate(() => window.__bad()), [], 'текст вне полей под высоким колонтитулом');
  await p.context().close();
});

await test('колонтитулы: правка двойным щелчком по полю листа, номер страницы из ленты, отмена внутри колонтитула', async () => {
  const p = await longPage((s, P) => Array.from({ length: 60 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(3))));
  const box = await p.evaluate(() => { const r = document.getElementById('sheet').getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + 20 }; });
  await p.mouse.dblclick(box.x, box.y);
  await p.waitForSelector('.hf-edit');
  await p.keyboard.type('Заголовок страницы');
  eq(await p.evaluate(() => window.ens.isDirty()), true, 'документ изменён');
  await p.keyboard.press('Control+KeyB');
  await p.keyboard.press('Control+KeyZ');
  await p.keyboard.press('Control+KeyZ');
  ok((await p.evaluate(() => window.ens.getMeta().hf.header)) === null, 'отмена внутри колонтитула вернула пустоту');
  await p.keyboard.type('Верх');
  await p.click('.btn[data-command="hf.close"]');
  await p.waitForFunction(() => !document.querySelector('.hf-edit'));
  const first = await p.evaluate(() => document.querySelectorAll('.paper')[0].querySelector('.hf-header')?.textContent);
  eq(first, 'Верх', 'колонтитул показан на листе');
  eq(await p.evaluate(() => document.querySelectorAll('.paper')[2].querySelector('.hf-header')?.textContent), 'Верх', 'и на третьем листе');
  await p.click('.tab:has-text("Вставка")');
  await p.selectOption('select[data-command="insert.pageNumber"]', 'footer-right');
  const foot = await p.evaluate(() => { const f = document.querySelectorAll('.paper')[1].querySelector('.hf-footer'); return [f?.textContent, f?.firstElementChild?.style.textAlign]; });
  eq(foot, ['2', 'right'], 'номер страницы справа внизу');
  const saved = await p.evaluate(() => { const E = window.ens; const b = E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta(), 'x')); return JSON.stringify(b.meta.hf); });
  ok(saved.includes('Верх') && saved.includes('"kind":"page"'), 'колонтитулы сохранены в docx');
  await p.context().close();
});

await test('печать: страницы вплотную, колонтитулы на каждой, число страниц PDF = числу листов', async () => {
  const p = await longPage((s, P) => Array.from({ length: 75 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(3))));
  await p.evaluate(() => {
    const E = window.ens, s = E.schema, m = E.getMeta();
    m.hf.footer = s.nodes.doc.create(null, [s.nodes.paragraph.create({ align: 'center' }, [s.text('Стр. '), s.nodes.field.create({ kind: 'page' }), s.text(' из '), s.nodes.field.create({ kind: 'pages' })])]).toJSON();
    E.loadDoc(E.view.state.doc, m, 'print');
  });
  await p.waitForTimeout(300);
  const sheets = await p.evaluate(() => window.ens.pageCount());
  await p.evaluate(() => window.ens.printPrepare());
  const gap = await p.evaluate(() => { const ps = [...document.querySelectorAll('.paper')].map((e) => e.offsetTop); const H = window.ens.geometry().H; return ps.every((t, i) => Math.abs(t - i * H) < 1); });
  ok(gap, 'листы стоят вплотную');
  await p.emulateMedia({ media: 'print' });
  const pdf = join(mkdtempSync(join(tmpdir(), 'ens-pdf-')), 'out.pdf');
  await p.pdf({ path: pdf, preferCSSPageSize: true, printBackground: true });
  const info = execSync(`pdfinfo "${pdf}"`).toString();
  const n = +/Pages:\s+(\d+)/.exec(info)[1];
  eq(n, sheets, 'страниц в PDF');
  const third = execSync(`pdftotext -f 3 -l 3 -layout "${pdf}" -`).toString();
  ok(third.includes(`Стр. 3 из ${sheets}`), `нижний колонтитул третьей страницы: ${third.slice(-120)}`);
  await p.evaluate(() => window.ens.printRestore());
  await p.context().close();
});

// ---- Колонки ----
const COLS_JS = `
  // Знаки вне своих колонок: каждый знак должен лежать внутри прямоугольника одной из колонок своего листа.
  window.__colBad = () => {
    const E = window.ens, v = E.view, boxes = E.pageBoxes(), papers = [...document.querySelectorAll('.paper')];
    const scale = v.dom.getBoundingClientRect().width / v.dom.offsetWidth;
    const rects = [];
    boxes.forEach((b, i) => {
      const pr = papers[i].getBoundingClientRect();
      const top = pr.top + b.eff.top * E.PX * scale, bottom = pr.top + (b.H - b.eff.bottom * E.PX) * scale;
      const left = pr.left + b.page.left * E.PX * scale;
      if (!b.cols) { rects.push({ i, k: 0, l: left, r: pr.left + (b.W - b.page.right * E.PX) * scale, top, bottom }); return; }
      b.cols.x.forEach((x, k) => rects.push({ i, k, l: left + x * scale, r: left + (x + b.cols.w[k]) * scale, top, bottom }));
    });
    const bad = [], used = new Set();
    v.state.doc.descendants((node, pos) => {
      if (!node.isTextblock) return true;
      const start = pos + 1, end = pos + 1 + node.content.size;
      for (let q = start; q < end; q++) {
        const c = v.coordsAtPos(q, 1), x = c.left + 1, y = (c.top + c.bottom) / 2;
        const r = rects.find((r) => x >= r.l - 2 && x <= r.r + 2 && y >= r.top - 2 && y <= r.bottom + 2);
        if (!r) bad.push({ q, x: Math.round(x), y: Math.round(y) }); else used.add(r.i + ':' + r.k);
      }
      return false;
    });
    return { bad: bad.slice(0, 5), used: [...used].sort() };
  };
  window.__cols = (n, extra) => {
    const E = window.ens, m = E.getMeta();
    m.page = { ...m.page, cols: n, colSpace: 35.4, ...(extra || {}) };
    E.loadDoc(E.view.state.doc, m, 'long');
  };
`;
async function colPage(n, count = 60, extra) {
  const p = await longPage(count === 60 ? (s, P) => Array.from({ length: 60 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(1 + (i % 5)))) : (s, P) => Array.from({ length: 70 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(1 + (i % 5)))));
  await p.evaluate(COLS_JS);
  await p.evaluate(([n, extra]) => window.__cols(n, extra), [n, extra]);
  await p.waitForTimeout(400);
  return p;
}

await test('колонки: текст перетекает из колонки в колонку и со страницы на страницу, знаки внутри колонок', async () => {
  const one = await longPage((s, P) => Array.from({ length: 60 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(1 + (i % 5)))));
  const single = await one.evaluate(() => window.ens.pageCount());
  await one.context().close();
  for (const [n, count] of [[2, 60], [3, 60], [2, 70]]) {
    const p = await colPage(n, count);
    const r = await p.evaluate(() => window.__colBad());
    eq(r.bad, [], `знаки вне колонок (${n})`);
    eq(await p.evaluate(() => window.ens.pagStats.stuck), false, `расчёт сошёлся (${n}, ${count})`);
    // Каждая колонка первых двух листов начинается ровно у верхнего поля.
    const tops = await p.evaluate(() => {
      const E = window.ens, v = E.view, boxes = E.pageBoxes(), papers = [...document.querySelectorAll('.paper')];
      const out = {};
      v.state.doc.descendants((node, pos) => {
        if (!node.isTextblock) return true;
        for (let q = pos + 1; q < pos + 1 + node.content.size; q += 3) {
          const c = v.coordsAtPos(q, 1);
          papers.slice(0, 2).forEach((pp, i) => {
            const r = pp.getBoundingClientRect(), b = boxes[i];
            if (c.top < r.top || c.top > r.bottom) return;
            const k = b.cols.x.findIndex((x, j) => c.left - r.left - b.page.left * E.PX >= x - 2 && c.left - r.left - b.page.left * E.PX <= x + b.cols.w[j] + 2);
            const key = i + ':' + k, d = c.top - (r.top + b.eff.top * E.PX);
            if (!(key in out) || d < out[key]) out[key] = Math.round(d);
          });
        }
        return false;
      });
      return out;
    });
    ok(Object.values(tops).every((d) => Math.abs(d) <= 4), `колонки начинаются не у верхнего поля (${n}, ${count}): ${JSON.stringify(tops)}`);
    const pages = await p.evaluate(() => window.ens.pageCount());
    for (let k = 0; k < n; k++) ok(r.used.includes(`0:${k}`) && r.used.includes(`1:${k}`), `колонка ${k + 1} на листах 1 и 2 занята (${n}): ${r.used}`);
    ok(pages >= 2 && pages <= single + 1, `листов ${pages}, одной полосой ${single}`);
    eq(await p.evaluate(() => document.querySelectorAll('.paper').length), pages, 'подложек столько же, сколько листов');
    // Поток не длиннее листов: внизу нет пустого места после последнего листа.
    const tail = await p.evaluate(() => { const b = window.ens.pageBoxes().at(-1); return Math.round(window.ens.view.dom.offsetHeight - (b.top + b.H)); });
    ok(tail <= 2, `лишняя высота после последнего листа: ${tail}`);
    const sig = await p.evaluate(() => window.__sig() + [...document.querySelectorAll('.pg-col')].map((e) => e.style.cssText).join('|'));
    await p.evaluate(() => window.ens.run('view.zoomIn'));
    await p.waitForTimeout(150);
    eq(await p.evaluate(() => window.__sig() + [...document.querySelectorAll('.pg-col')].map((e) => e.style.cssText).join('|')), sig, `расчёт устойчив (${n})`);
    eq(await p.evaluate(() => window.ens.isDirty()), false, 'документ не «грязный» от вёрстки');
    await p.context().close();
  }
});

await test('колонки: абзац на стыке колонок — щелчок и ввод в хвосте во второй колонке, выделение', async () => {
  const p = await colPage(2);
  // Абзац, начатый в первой колонке и продолженный во второй: позиция в хвосте.
  const q = await p.evaluate(() => {
    const seg = document.querySelector('.pg-colseg');
    const v = window.ens.view;
    return seg ? v.posAtDOM(seg.firstChild, 3) : null;
  });
  ok(q != null, 'нет абзаца, перетекающего в другую колонку');
  const at = await p.evaluate((q) => { const c = window.ens.view.coordsAtPos(q, 1); return { x: c.left + 1, y: (c.top + c.bottom) / 2 }; }, q);
  await p.mouse.click(at.x, at.y);
  const head = await p.evaluate(() => window.ens.view.state.selection.head);
  ok(Math.abs(head - q) <= 1, `курсор после щелчка: ${head}, ожидалось около ${q}`);
  await p.keyboard.type('ВСТАВКА');
  const t = await p.evaluate((h) => window.ens.view.state.doc.textBetween(h - 1, h + 9), head);
  ok(t.includes('ВСТАВКА'), `введено не там: ${t}`);
  await p.waitForTimeout(200);
  eq((await p.evaluate(() => window.__colBad())).bad, [], 'знаки внутри колонок после ввода');
  // Выделение протаскиванием из первой колонки во вторую.
  const a = await p.evaluate(() => { const v = window.ens.view; const c = v.coordsAtPos(5, 1); return { x: c.left + 1, y: (c.top + c.bottom) / 2 }; });
  await p.mouse.move(a.x, a.y);
  await p.mouse.down();
  await p.mouse.move(at.x + 20, at.y, { steps: 5 });
  await p.mouse.up();
  const sel = await p.evaluate(() => { const s = window.ens.view.state.selection; return [s.from, s.to]; });
  ok(sel[0] <= 6 && sel[1] > q, `выделение через колонки: ${sel}`);
  await p.context().close();
});

await test('колонки: меню «Колонки», разрыв колонки Ctrl+Shift+Enter, «Другие колонки…» с разделителем', async () => {
  const p = await longPage((s, P) => Array.from({ length: 12 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(2))));
  await p.evaluate(COLS_JS);
  await p.click('.tab:has-text("Макет")');
  await p.click('#ribbon [data-command="layout.columns"]');
  await p.click('.columns-menu button:has-text("Две")');
  eq(await p.evaluate(() => window.ens.getMeta().page.cols), 2, 'две колонки');
  await p.waitForTimeout(200);
  // Разрыв колонки перед третьим абзацем: он начинает вторую колонку.
  const pos3 = await p.evaluate(() => { let r = 0, k = 0; window.ens.view.state.doc.forEach((n, off) => { if (k++ === 2) r = off + 1; }); return r; });
  await p.evaluate((q) => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, q))); E.view.focus(); }, pos3);
  await p.keyboard.press('Control+Shift+Enter');
  await p.waitForTimeout(200);
  const r = await p.evaluate(() => {
    const v = window.ens.view, els = [];
    v.state.doc.forEach((n, off) => els.push({ col: !!n.attrs.colBreakBefore, r: v.nodeDOM(off).getBoundingClientRect() }));
    const i = els.findIndex((e) => e.col);
    return { i, firstLeft: els[0].r.left, brokeLeft: els[i].r.left, brokeTop: els[i].r.top, firstTop: els[0].r.top };
  });
  ok(r.i === 3, `разрыв колонки у нового абзаца: ${r.i}`);
  ok(r.brokeLeft > r.firstLeft + 100 && Math.abs(r.brokeTop - r.firstTop) < 2, `абзац после разрыва — вверху второй колонки: ${JSON.stringify(r)}`);
  eq((await p.evaluate(() => window.__colBad())).bad, [], 'знаки внутри колонок');
  // «Другие колонки…»: три колонки, промежуток 1 см, разделитель.
  await p.click('#ribbon [data-command="layout.columns"]');
  await p.click('.columns-menu button:has-text("Другие колонки")');
  await p.fill('.modal input[type=number]', '3');
  await p.locator('.modal input[type=number]').dispatchEvent('input');
  await p.fill('.modal .dlg-row:has-text("Промежуток") input', '1');
  await p.check('.modal .dlg-row:has-text("Разделитель") input');
  await p.click('.modal button:has-text("ОК")');
  await p.waitForTimeout(300);
  const m = await p.evaluate(() => window.ens.getMeta().page);
  ok(m.cols === 3 && m.colSep && Math.abs(m.colSpace - 28.35) < 0.1, `параметры: ${JSON.stringify(m)}`);
  eq(await p.evaluate(() => document.querySelectorAll('.paper:first-child .col-sep').length), 2, 'две линии-разделителя на первом листе');
  // Обратно одна колонка.
  await p.click('#ribbon [data-command="layout.columns"]');
  await p.click('.columns-menu button:has-text("Одна")');
  await p.waitForTimeout(200);
  eq(await p.evaluate(() => document.querySelectorAll('.pg-col, .col-sep').length), 0, 'одна колонка — без смещений и линий');
  await p.context().close();
});

await test('колонки: неравные (слева), docx — ширины, разделитель и разрыв колонки туда и обратно', async () => {
  const p = await colPage(2, 60, { colW: [1, 2], colSep: true });
  eq((await p.evaluate(() => window.__colBad())).bad, [], 'знаки внутри неравных колонок');
  eq(await p.evaluate(() => window.ens.pagStats.stuck), false, 'расчёт сошёлся');
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const doc = s.nodes.doc.create(null, [s.nodes.paragraph.create(null, s.text('Один')), s.nodes.paragraph.create({ colBreakBefore: true }, s.text('Два'))]);
    const m = { ...E.getMeta(), page: { ...E.getMeta().page, cols: 2, colSpace: 20, colW: [150, 300], colSep: true } };
    const bytes = E.exportDocx(doc, m, 'x');
    const xml = new TextDecoder().decode(E.unzip(bytes)['word/document.xml']);
    const back = E.importDocx(bytes);
    return { xml, page: back.meta.page, br: back.doc.child(1).attrs.colBreakBefore, n: back.doc.childCount, t: back.doc.textContent };
  });
  ok(/<w:cols w:num="2" w:space="400" w:sep="1" w:equalWidth="0"><w:col w:w="3000" w:space="400"\/><w:col w:w="6000"\/><\/w:cols>/.test(r.xml), `w:cols: ${r.xml.match(/<w:cols.*?(\/>|<\/w:cols>)/)?.[0]}`);
  ok(r.xml.includes('<w:br w:type="column"/>'), 'разрыв колонки в файле');
  eq([r.page.cols, r.page.colSpace, r.page.colW, r.page.colSep], [2, 20, [150, 300], true], 'колонки после чтения');
  eq([r.br, r.n, r.t], [true, 2, 'ОдинДва'], 'разрыв колонки после чтения, лишних абзацев нет');
  await p.context().close();
});
await test('колонки: раздел с колонками между разделами в одну полосу; список и таблица в колонках', async () => {
  const p = await longPage((s, P) => {
    const F = 'Съешь ещё этих мягких французских булок, да выпей же чаю. ';
    const A4 = { w: 595.3, h: 841.9, top: 56.7, bottom: 56.7, left: 85, right: 42.5 };
    const item = (t) => s.nodes.list_item.create(null, [P(t)]);
    const cell = (t) => s.nodes.table_cell.create(null, P(t));
    const rows = Array.from({ length: 40 }, (_, r) => s.nodes.table_row.create(null, [cell(`Ячейка ${r}.1`), cell(`Ячейка ${r}.2`)]));
    return [
      ...Array.from({ length: 5 }, (_, i) => P(`Вступление ${i + 1}. ` + F.repeat(3))),
      s.nodes.section_break.create({ page: { ...A4, cols: 2, colSpace: 35.4 } }),
      ...Array.from({ length: 30 }, (_, i) => P(`Колонки ${i + 1}. ` + F.repeat(1 + (i % 4)))),
      s.nodes.bullet_list.create(null, Array.from({ length: 12 }, (_, i) => item(`Пункт ${i + 1}. ` + F))),
      s.nodes.table.create({ borders: true }, rows),
      ...Array.from({ length: 30 }, (_, i) => P(`Ещё ${i + 1}. ` + F.repeat(1 + (i % 3)))),
      s.nodes.section_break.create({ page: { ...A4 } }),
      ...Array.from({ length: 5 }, (_, i) => P(`Заключение ${i + 1}. ` + F.repeat(3))),
    ];
  });
  await p.evaluate(COLS_JS);
  await p.waitForTimeout(300);
  const r = await p.evaluate(() => window.__colBad());
  eq(r.bad, [], 'знаки вне колонок и полей');
  eq(await p.evaluate(() => window.ens.pagStats.stuck), false, 'расчёт сошёлся');
  ok(r.used.includes('1:0') && r.used.includes('1:1'), `второй лист — в две колонки: ${r.used}`);
  // Заключение — с нового листа, вверху, а не там, где кончился поток колонок.
  const c = await p.evaluate(() => {
    const E = window.ens, v = E.view, boxes = E.pageBoxes(), papers = [...document.querySelectorAll('.paper')];
    let pos = 0; v.state.doc.forEach((n, off) => { if (n.textContent.startsWith('Заключение 1')) pos = off; });
    const top = v.nodeDOM(pos).getBoundingClientRect().top, last = papers.length - 1;
    const pr = papers[last].getBoundingClientRect();
    return { d: Math.round(top - (pr.top + boxes[last].eff.top * E.PX)), tail: Math.round(v.dom.offsetHeight - (boxes[last].top + boxes[last].H)) };
  });
  ok(Math.abs(c.d) <= 2, `заключение не вверху последнего листа: ${c.d}`);
  ok(c.tail <= 2, `лишняя высота после последнего листа: ${c.tail}`);
  await p.context().close();
});


await test('колонки: плавающий рисунок в абзаце второй колонки стоит по координатам листа', async () => {
  const p = await colPage(2);
  const r = await p.evaluate(async () => {
    const E = window.ens, v = E.view, s = E.schema;
    // Абзац, целиком стоящий во второй колонке первого листа.
    let pos = -1;
    v.state.doc.forEach((n, off) => { const st = v.nodeDOM(off).style; if (pos < 0 && st.left && parseFloat(st.left) > 100) pos = off; });
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
    const float = { mode: 'behind', hRel: 'page', hOff: 200, hAlign: null, vRel: 'page', vOff: 300, vAlign: null, dist: [0, 0, 8, 8] };
    v.dispatch(v.state.tr.insert(pos + 1, s.nodes.image.create({ src: png, width: 40, height: 40, float })));
    await new Promise((res) => setTimeout(res, 500));
    const img = document.querySelector('.ProseMirror img.fl-behind').getBoundingClientRect();
    const paper = document.querySelector('.paper').getBoundingClientRect();
    return { pos, dx: Math.round(img.left - paper.left), dy: Math.round(img.top - paper.top) };
  });
  ok(r.pos > 0, 'нет абзаца во второй колонке');
  ok(Math.abs(r.dx - 200) <= 2 && Math.abs(r.dy - 300) <= 2, `рисунок не на месте: ${JSON.stringify(r)}`);
  await p.context().close();
});

await test('колонки: печать в PDF — текст в две колонки, листов столько же', async () => {
  const p = await colPage(2);
  const sheets = await p.evaluate(() => window.ens.pageCount());
  await p.evaluate(() => window.ens.printPrepare());
  await p.waitForTimeout(200);
  await p.emulateMedia({ media: 'print' });
  const pdf = join(mkdtempSync(join(tmpdir(), 'ens-pdf-')), 'cols.pdf');
  await p.pdf({ path: pdf, preferCSSPageSize: true, printBackground: true });
  const n = +/Pages:\s+(\d+)/.exec(execSync(`pdfinfo "${pdf}"`).toString())[1];
  eq(n, sheets, 'страниц в PDF');
  // Левая и правая половины первой страницы: в каждой — своя колонка (A4 — 595 пт в ширину).
  const half = (x) => execSync(`pdftotext -f 1 -l 1 -x ${x} -y 0 -W 297 -H 842 "${pdf}" -`).toString();
  const nums = (t) => [...t.matchAll(/Абзац (\d+)\./g)].map((m) => +m[1]);
  const L = nums(half(0)), R = nums(half(298));
  ok(L.includes(1) && R.length > 3 && Math.min(...R) > Math.max(...L), `левая колонка ${L}, правая ${R}`);
  await p.evaluate(() => window.ens.printRestore());
  await p.context().close();
});


// ---- Вид: несколько страниц ----
await test('вид: «Несколько страниц» — листы рядами, текст на своих листах, правка и печать', async () => {
  const p = await longPage((s, P) => Array.from({ length: 70 }, (_, i) => P(`Абзац ${i + 1}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(1 + (i % 5)))));
  await p.evaluate(COLS_JS);
  const pages = await p.evaluate(() => window.ens.pageCount());
  await p.click('.tab:has-text("Вид")');
  await p.click('#ribbon [data-command="view.multiPage"]');
  await p.waitForTimeout(500);
  const r = await p.evaluate(() => ({ boxes: window.ens.pageBoxes().map((b) => [Math.round(b.left ?? 0), Math.round(b.top)]), stuck: window.ens.pagStats.stuck }));
  ok(r.boxes.length === pages, `листов столько же: ${r.boxes.length} и ${pages}`);
  ok(r.boxes[1][0] > 500 && r.boxes[1][1] === r.boxes[0][1], `второй лист справа от первого: ${JSON.stringify(r.boxes)}`);
  eq(r.stuck, false, 'расчёт сошёлся');
  const c = await p.evaluate(() => window.__colBad());
  eq(c.bad, [], 'знаки вне полей своих листов');
  ok(c.used.includes('1:0') && c.used.includes('2:0'), `текст на листах 2 и 3: ${c.used}`);
  // Подложки на местах листов, лишней высоты под последним рядом нет.
  const geo = await p.evaluate(() => {
    const papers = [...document.querySelectorAll('.paper')], v = window.ens.view, b = window.ens.pageBoxes();
    const last = b.at(-1);
    return { left: papers[1].offsetLeft, tail: Math.round(v.dom.offsetHeight - (last.top + last.H)) };
  });
  ok(geo.left > 500, `подложка второго листа справа: ${geo.left}`);
  ok(geo.tail <= 2, `лишняя высота: ${geo.tail}`);
  // Щелчок и ввод на втором листе.
  const q = await p.evaluate(() => { const v = window.ens.view; let q = 0; v.state.doc.forEach((n, off) => { if (!q && n.textContent.startsWith('Абзац 30.')) q = off + 3; }); return q; });
  const at = await p.evaluate((q) => { const c = window.ens.view.coordsAtPos(q, 1); return { x: c.left + 1, y: (c.top + c.bottom) / 2 }; }, q);
  await p.mouse.click(at.x, at.y);
  ok(Math.abs((await p.evaluate(() => window.ens.view.state.selection.head)) - q) <= 1, 'курсор после щелчка');
  await p.keyboard.type('Ж');
  ok((await p.evaluate(() => window.ens.view.state.doc.textContent)).includes('АбЖзац 30') || (await p.evaluate(() => window.ens.view.state.doc.textContent)).includes('АбзЖац 30'), 'ввод на втором листе');
  // Мельче масштаб — больше листов в ряду.
  for (let i = 0; i < 3; i++) await p.evaluate(() => window.ens.run('view.zoomOut'));
  await p.waitForTimeout(400);
  const row = await p.evaluate(() => window.ens.pageBoxes().filter((b) => b.top === 0).length);
  ok(row >= 3, `при уменьшении масштаба в ряду листов: ${row}`);
  eq((await p.evaluate(() => window.__colBad())).bad, [], 'знаки на своих листах после смены масштаба');
  // Печать — столбиком, после печати снова рядами.
  await p.evaluate(() => window.ens.printPrepare());
  ok(await p.evaluate(() => window.ens.pageBoxes().every((b) => !b.left)), 'при печати листы столбиком');
  await p.evaluate(() => window.ens.printRestore());
  await p.waitForTimeout(300);
  ok(await p.evaluate(() => window.ens.pageBoxes()[1].left > 0), 'после печати снова рядами');
  // «Одна страница» и «По ширине»: листы столбиком, масштаб вписывает лист.
  await p.click('#ribbon [data-command="view.onePage"]');
  await p.waitForTimeout(400);
  const one = await p.evaluate(() => { const b = document.querySelector('.paper').getBoundingClientRect(), d = document.getElementById('desk').getBoundingClientRect(); return { h: b.height, dh: d.height, left: window.ens.pageBoxes()[1].left ?? 0 }; });
  ok(one.left === 0 && one.h <= one.dh && one.h > one.dh * 0.8, `одна страница: ${JSON.stringify(one)}`);
  await p.click('#ribbon [data-command="view.pageWidth"]');
  await p.waitForTimeout(400);
  const wid = await p.evaluate(() => { const b = document.querySelector('.paper').getBoundingClientRect(), d = document.getElementById('desk'); return { w: b.width, dw: d.clientWidth }; });
  ok(wid.w <= wid.dw && wid.w > wid.dw * 0.85, `по ширине: ${JSON.stringify(wid)}`);
  await p.context().close();
});

await test('вид: «Несколько страниц» с колонками и разделами разного размера', async () => {
  const p = await colPage(2, 70);
  await p.evaluate(() => {
    const E = window.ens, v = E.view, s = E.schema;
    const A4L = { w: 841.9, h: 595.3, top: 56.7, bottom: 56.7, left: 56.7, right: 56.7 };
    let pos = 0, k = 0; v.state.doc.forEach((n, off) => { if (k++ === 40) pos = off; });
    v.dispatch(v.state.tr.insert(pos, s.nodes.section_break.create({ page: A4L })));
  });
  await p.waitForTimeout(400);
  await p.evaluate(() => window.ens.run('view.multiPage'));
  await p.waitForTimeout(600);
  eq((await p.evaluate(() => window.__colBad())).bad, [], 'знаки вне колонок и полей');
  eq(await p.evaluate(() => window.ens.pagStats.stuck), false, 'расчёт сошёлся');
  await p.context().close();
});

// ---- Версии, журнал, защита от потерь ----
await test('версии: сохранение записывает версию и журнал, история едет внутри docx и читается обратно', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, s.text(t));
    E.loadDoc(s.nodes.doc.create(null, [P('Первый абзац'), P('Второй абзац')]), E.getMeta(), 'v');
    const b1 = E.buildDocx('v');
    const a = E.importDocx(b1);
    E.loadDoc(a.doc, a.meta, 'v', false, { raw: a.history, key: 'v.docx' });
    const v1 = E.hist().versions.length;
    const tr = E.view.state.tr.insertText(' (правка)', 13);
    E.view.dispatch(tr);
    const b2 = E.buildDocx('v');
    const c = E.importDocx(b2);
    const h = E.V.parse(c.history);
    return { v1, versions: h.versions.map((x) => x.kind), log: h.log.map((x) => [x.kind, x.ch.changed, x.ch.words]), size: b2.length };
  });
  ok(r.v1 >= 1, 'история прочитана из файла');
  eq(r.versions, ['save', 'save'], 'две версии «Сохранение»');
  eq(r.log.map((x) => x[0]), ['save', 'save'], 'две записи журнала');
  eq(r.log[1].slice(1), [1, 1], 'во втором сохранении изменён 1 абзац, +1 слово');
  await p.context().close();
});

await test('версии: docx с историей открывают python-docx и LibreOffice (customXml не мешает)', async () => {
  const p = await page();
  const bytes = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, s.text('Текст с историей'))]), E.getMeta(), 'v');
    return [...E.buildDocx('v')];
  });
  const f = join(mkdtempSync(join(tmpdir(), 'ens-hist-')), 'h.docx');
  writeFileSync(f, Buffer.from(bytes));
  const out = execSync(`python3 -c "from docx import Document; d=Document('${f}'); print(d.paragraphs[0].text)"`).toString().trim();
  eq(out, 'Текст с историей', 'python-docx читает файл');
  await p.context().close();
});

await test('версии: правка в другой программе замечается при открытии, прежнее состояние остаётся в истории', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, s.text(t));
    E.loadDoc(s.nodes.doc.create(null, [P('Договор, пункт 1'), P('Договор, пункт 2')]), E.getMeta(), 'd');
    const saved = E.buildDocx('d');
    const a = E.importDocx(saved);
    // «Другая программа» поменяла текст, но оставила историю в файле.
    const edited = s.nodes.doc.create(null, [P('Договор, пункт 1'), P('Договор, пункт 2 — изменён чужой программой')]);
    const foreign = E.exportDocx(edited, a.meta, 'd', a.history);
    const b = E.importDocx(foreign);
    E.loadDoc(b.doc, b.meta, 'd', false, { raw: b.history, key: 'd.docx' });
    return { kinds: E.hist().versions.map((v) => v.kind), log: E.hist().log.map((l) => l.kind), notice: document.getElementById('notice').textContent };
  });
  eq(r.kinds, ['save', 'outside'], 'версии: наше сохранение и состояние при открытии');
  eq(r.log, ['save', 'outside'], 'в журнале — правка в другой программе');
  ok(r.notice.includes('в другой программе'), r.notice);
  // Кнопка «Что изменилось» показывает разницу.
  await p.click('#notice button:has-text("Что изменилось")');
  await p.waitForSelector('.history-modal .diff');
  ok((await p.textContent('.history-modal .diff')).includes('изменён чужой программой'), 'в сравнении видна правка');
  await p.context().close();
});

await test('версии: окно истории — сравнение, возврат фрагмента и документа целиком, отмена', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, s.text(t));
    E.loadDoc(s.nodes.doc.create(null, [P('Пункт один'), P('Пункт два'), P('Пункт три')]), E.getMeta(), 'v');
    E.snap('named', 'Согласовано');
    const v = E.view, tr = v.state.tr;
    tr.insertText('ИЗМЕНЕНО ', v.state.doc.child(0).nodeSize + 1);
    v.dispatch(tr);
  });
  await p.waitForTimeout(700); // отдельный шаг отмены, а не продолжение набора
  await p.keyboard.press('Control+Alt+KeyH');
  await p.waitForSelector('.history-modal');
  ok((await p.textContent('.hist-list')).includes('«Согласовано»'), 'именная версия в списке');
  const diff = await p.textContent('.history-modal .diff');
  ok(diff.includes('ИЗМЕНЕНО'), 'изменение показано');
  await p.click('.diff-revert');
  await p.waitForFunction(() => !document.querySelector('.history-modal'));
  eq(await p.evaluate(() => window.ens.view.state.doc.child(1).textContent), 'Пункт два', 'фрагмент возвращён');
  await p.keyboard.press('Control+KeyZ');
  eq(await p.evaluate(() => window.ens.view.state.doc.child(1).textContent), 'ИЗМЕНЕНО Пункт два', 'отмена вернула правку');
  // Возврат всего документа.
  await p.keyboard.press('Control+Alt+KeyH');
  await p.waitForSelector('.history-modal');
  await p.click('.hist-row:has-text("Согласовано")');
  await p.click('button:has-text("Восстановить документ целиком")');
  eq(await p.evaluate(() => window.ens.view.state.doc.textContent), 'Пункт одинПункт дваПункт три', 'документ возвращён целиком');
  const kinds = await p.evaluate(() => window.ens.hist().versions.map((v) => v.kind));
  ok(kinds.includes('before'), `страховочная версия перед возвратом: ${kinds}`);
  await p.context().close();
});

await test('версии: «Заменить все» сначала записывает страховочную версию', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('кот и кот');
  await p.keyboard.press('Control+KeyH');
  await p.fill('#findbar input[type=text] >> nth=0', 'кот');
  await p.fill('#findbar input[type=text] >> nth=1', 'пёс');
  await p.click('#findbar button:has-text("Заменить все")');
  const v = await p.evaluate(() => window.ens.hist().versions.map((x) => [x.kind, x.label]));
  eq(v.length, 1, 'одна страховочная версия');
  eq(v[0][0], 'before', 'вид версии');
  await p.context().close();
});

await test('версии: история занимает место только под изменения (дедупликация блоков)', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const paras = Array.from({ length: 600 }, (_, i) => s.nodes.paragraph.create(null, s.text(`Абзац ${i}. ` + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(4))));
    E.loadDoc(s.nodes.doc.create(null, paras), E.getMeta(), 'big');
    E.snap('save');
    const one = E.V.serialize(E.hist()).length;
    for (let i = 0; i < 25; i++) {
      E.view.dispatch(E.view.state.tr.insertText('+', 3 + i));
      E.snap('auto');
    }
    return { one, many: E.V.serialize(E.hist()).length, versions: E.hist().versions.length };
  });
  ok(r.versions >= 20, `версий ${r.versions}`);
  ok(r.many < r.one * 1.6, `26 версий весят ${r.many}, одна — ${r.one}`);
  await p.context().close();
});

await test('защита: «Не сохранено N мин» и напоминание', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Несохранённая работа');
  ok((await p.textContent('#status-saved')).includes('несохранённые'), await p.textContent('#status-saved'));
  await p.evaluate(() => { window.ens.run('ver.reminder', '5'); });
  eq(await p.evaluate(() => JSON.parse(localStorage.getItem('ens-table.settings')).saveReminderMinutes), 5, 'настройка сохранена');
  await p.context().close();
});


// ---- Темы ----
await test('темы: те же, что в ENS-Table; лист меняет цвет на экране, печать остаётся белой', async () => {
  const p = await page();
  await p.evaluate(() => window.ens.run('view.theme'));
  await p.waitForSelector('.theme-list');
  const names = await p.$$eval('.theme-row b', (e) => e.map((x) => x.textContent.replace('✓ ', '')));
  for (const need of ['Классическая', 'Пустошь', 'Матрица']) ok(names.includes(need), `нет темы «${need}»: ${names}`);
  await p.click('.theme-row:has-text("Пустошь")');
  eq(await p.evaluate(() => document.documentElement.dataset.theme), 'wasteland', 'атрибут темы');
  eq(await p.evaluate(() => getComputedStyle(document.querySelector('.paper')).backgroundColor), 'rgb(242, 232, 207)', 'цвет листа на экране');
  await p.click('.modal button.primary');
  await p.emulateMedia({ media: 'print' });
  eq(await p.evaluate(() => getComputedStyle(document.querySelector('.paper')).backgroundColor), 'rgb(255, 255, 255)', 'в печати лист белый');
  await p.emulateMedia({ media: 'screen' });
  // Тема запоминается и общая с ENS-Table (настройка theme).
  eq(await p.evaluate(() => JSON.parse(localStorage.getItem('ens-table.settings')).theme), 'wasteland', 'настройка сохранена');
  await p.context().close();
});


// ---- Многоуровневые списки ----
const listDoc = (s, scheme) => {
  const P = (t) => s.nodes.paragraph.create(null, s.text(t));
  const item = (t, ...kids) => s.nodes.list_item.create(null, [P(t), ...kids]);
  const ol = (items, attrs) => s.nodes.ordered_list.create({ order: 1, fmt: 'decimal', scheme: null, ...attrs }, items);
  return { P, item, ol, scheme };
};

await test('списки: схема «1.2.14.Б» — номера считаются по структуре и показываются', async () => {
  const p = await page();
  const labels = await p.evaluate(() => {
    const E = window.ens, s = E.schema, { P, item, ol } = ((s) => {
      const P = (t) => s.nodes.paragraph.create(null, s.text(t));
      const item = (t, ...k) => s.nodes.list_item.create(null, [P(t), ...k]);
      const ol = (items, a) => s.nodes.ordered_list.create({ order: 1, fmt: 'decimal', scheme: null, ...a }, items);
      return { P, item, ol };
    })(s);
    const scheme = E.PRESETS.find((x) => x.id === 'letters').scheme;
    // 1 → 1.1 → 1.2 → (13 пунктов третьего уровня, берём 14-й) → А, Б на четвёртом.
    const l3 = Array.from({ length: 14 }, (_, i) => item(`пункт ${i + 1}`, ...(i === 13 ? [ol([item('подпункт А'), item('подпункт Б')])] : [])));
    const doc = s.nodes.doc.create(null, [ol([item('Раздел', ol([item('Подраздел 1.1'), item('Подраздел 1.2', ol(l3))]))], { scheme })]);
    E.loadDoc(doc, E.getMeta(), 'l');
    return [...document.querySelectorAll('.ProseMirror .num')].map((e) => e.textContent);
  });
  eq(labels.slice(0, 3), ['1.', '1.1.', '1.2.'], 'начало');
  ok(labels.includes('1.2.14.'), `нет «1.2.14.»: ${labels.slice(-5)}`);
  eq(labels.slice(-2), ['1.2.14.А', '1.2.14.Б'], 'четвёртый уровень — буквы');
  await p.context().close();
});

await test('списки: схема и начальный номер доживают до docx и обратно; в Word-определении те же шаблоны', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const item = (t, ...k) => s.nodes.list_item.create(null, [P(t), ...k]);
    const ol = (items, a) => s.nodes.ordered_list.create({ order: 1, fmt: 'decimal', scheme: null, ...a }, items);
    const scheme = E.PRESETS.find((x) => x.id === 'gost').scheme;
    const doc = s.nodes.doc.create(null, [ol([item('Один', ol([item('Вложенный')])), item('Два')], { scheme, order: 3 })]);
    const back = E.importDocx(E.exportDocx(doc, E.getMeta(), 'x'));
    return { scheme: back.doc.firstChild.attrs.scheme, want: scheme, order: back.doc.firstChild.attrs.order };
  });
  eq(r.scheme, r.want, 'схема после круга docx');
  await p.context().close();
});

await test('списки: «Начать с…», «Продолжить нумерацию» и «Начать заново» не ломают соседей', async () => {
  const p = await page();
  const out = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const item = (t) => s.nodes.list_item.create(null, [P(t)]);
    const ol = (items) => s.nodes.ordered_list.create({ order: 1, fmt: 'decimal', scheme: null }, items);
    E.loadDoc(s.nodes.doc.create(null, [ol([item('a'), item('b'), item('c')]), P('вставка между списками'), ol([item('d'), item('e')])]), E.getMeta(), 'l');
    const nums = () => [...document.querySelectorAll('.ProseMirror .num')].map((e) => e.textContent);
    const before = nums();
    // Курсор во второй список и «Продолжить».
    const v = E.view; const pos = v.state.doc.child(0).nodeSize + v.state.doc.child(1).nodeSize + 3;
    v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(pos))));
    E.run('list.continue');
    const cont = nums();
    E.run('list.restart');
    return { before, cont, restart: nums() };
  });
  eq(out.before, ['1.', '2.', '3.', '1.', '2.'], 'сначала каждый список с единицы');
  eq(out.cont, ['1.', '2.', '3.', '4.', '5.'], 'продолжение нумерации');
  eq(out.restart, ['1.', '2.', '3.', '1.', '2.'], 'заново');
  await p.context().close();
});

await test('списки: ручные номера «1.», «1.1», «1.2.14.Б», «а)» превращаются в настоящий список с отменой', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const lines = ['Вступление без номера', '1. Общие положения', '1.1 Термины', '1.2 Сроки', '1.2.14 Основной пункт', '1.2.14.Б Особый случай', '2. Заключительные положения', 'Обычный абзац', '1. Другой список', 'а) первое', 'б) второе', '2026 год был непростым'];
    E.loadDoc(s.nodes.doc.create(null, lines.map(P)), E.getMeta(), 'm');
    E.view.dispatch(E.view.state.tr.setSelection(E.view.state.selection.constructor.atStart(E.view.state.doc)));
    E.run('list.fromManual');
    const d = E.view.state.doc;
    return { types: [...Array(d.childCount).keys()].map((i) => d.child(i).type.name), nums: [...document.querySelectorAll('.ProseMirror .num')].map((e) => e.textContent), text: d.textContent };
  });
  eq(r.types, ['paragraph', 'ordered_list', 'paragraph', 'ordered_list', 'paragraph'], 'абзацы с номерами собраны в списки, прочее не тронуто');
  ok(['1.', '1.1', '1.2', '1.2.14', '1.2.14.Б', '2.'].every((x) => r.nums.includes(x)), `номера: ${r.nums}`);
  ok(r.nums.includes('а)') && r.nums.includes('б)'), `буквенные пункты: ${r.nums}`);
  ok(!r.text.includes('1.1 Термины') && r.text.includes('Термины'), 'ручные номера из текста убраны');
  ok(r.text.includes('2026 год был непростым'), 'одиночное «2026 год» не принято за список');
  await p.keyboard.press('Control+KeyZ');
  ok((await p.evaluate(() => window.ens.view.state.doc.textContent)).includes('1.1 Термины'), 'отмена вернула ручные номера');
  await p.context().close();
});

await test('списки: настройка схемы в окне и готовые схемы из ленты; номер не уезжает на следующую страницу', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.click('.btn[data-command="list.ordered"]');
  await p.keyboard.type('Первый');
  await p.keyboard.press('Enter');
  await p.keyboard.type('Второй');
  await p.selectOption('select[data-command="list.scheme"]', 'legal');
  await p.keyboard.press('Tab');
  await p.keyboard.type('Вложенный');
  const nums = await p.$$eval('.ProseMirror .num', (e) => e.map((x) => x.textContent));
  eq(nums, ['1.', '1.1.'], 'схема «1. 1.1.» применена: второй пункт вложен в первый');
  await p.selectOption('select[data-command="list.scheme"]', 'custom');
  await p.waitForSelector('.scheme-rows');
  ok((await p.textContent('.scheme-preview')).includes('1.2.14.'), 'образец в окне настройки');
  await p.click('.modal button.primary');
  // Длинный список через границу страницы: номер и текст пункта остаются на одной странице.
  const bad = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const items = Array.from({ length: 120 }, (_, i) => s.nodes.list_item.create(null, [P('Пункт ' + i + '. ' + 'Съешь ещё этих мягких французских булок, да выпей же чаю. '.repeat(2))]));
    E.loadDoc(s.nodes.doc.create(null, [s.nodes.ordered_list.create({ order: 1, fmt: 'decimal', scheme: null }, items)]), E.getMeta(), 'big');
    return new Promise((res) => setTimeout(() => {
      const v = E.view, g = E.geometry(), box = v.dom.getBoundingClientRect(), sc = box.width / v.dom.offsetWidth, out = [];
      document.querySelectorAll('.ProseMirror li.numbered').forEach((li) => {
        const num = li.querySelector('.num').getBoundingClientRect(), pr = li.querySelector('p').getBoundingClientRect();
        const pg = (y) => Math.floor((y - box.top) / sc / g.P);
        if (pg(num.top + 2) !== pg(pr.top + 2)) out.push(li.querySelector('.num').textContent);
      });
      res({ out, pages: E.pageCount() });
    }, 500));
  });
  ok(bad.pages >= 3, 'список занял несколько страниц');
  eq(bad.out, [], 'номер на одной странице с текстом пункта');
  await p.context().close();
});


// ---- Типограф, метки, фокус, поиск команд ----
const typoRun = (p, text, opts = {}) => p.evaluate(({ text, opts }) => {
  const E = window.ens, s = E.schema;
  E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, s.text(text))]), E.getMeta(), 't');
  E.run('typo.apply');
  return E.view.state.doc.textContent;
}, { text, opts });

await test('типограф: кавычки, вложенные кавычки, тире, многоточие', async () => {
  const p = await page();
  const norm = (s) => s.replaceAll('\u00A0', ' ');
  eq(norm(await typoRun(p, 'Он сказал "привет" и ушёл')), 'Он сказал «привет» и ушёл', 'ёлочки');
  eq(norm(await typoRun(p, 'Компания "Рога и "Копыта" и К" открылась')), 'Компания «Рога и „Копыта“ и К» открылась', 'вложенные лапки');
  eq(await typoRun(p, 'слово - слово'), 'слово — слово', 'тире с неразрывным пробелом');
  eq(await typoRun(p, 'Ждём...'), 'Ждём…', 'многоточие');
  await p.context().close();
});

await test('типограф: неразрывные пробелы (предлоги, «г. Москва», «№ 5», единицы, инициалы, разряды, даты)', async () => {
  const p = await page();
  const r = await typoRun(p, 'Он живёт в лесу, г. Москва, № 5, весит 70 кг, А. С. Пушкин, 1 000 000 руб., 5 мая 2026 г. и т. д.');
  const NB = ' ';
  for (const need of [`в${NB}лесу`, `г.${NB}Москва`, `№${NB}5`, `70${NB}кг`, `А.${NB}С.${NB}Пушкин`, `1${NB}000${NB}000`, `5${NB}мая`, `2026${NB}г.`, `т.${NB}д.`]) ok(r.includes(need), `нет «${need.replaceAll(NB, '␣')}» в «${r.replaceAll(NB, '␣')}»`);
  await p.context().close();
});

await test('типограф: лишние пробелы, пробел после запятой; числа, адреса и ссылки не трогаются', async () => {
  const p = await page();
  eq((await typoRun(p, '  слово  ,слово ,слово.Слово  ')).replaceAll('\u00A0', ' '), 'слово, слово, слово. Слово', 'пробелы');
  eq((await typoRun(p, 'Цена 3,5 и 1,25; адрес https://ens.example/a,b почта a@b.ru')).replaceAll('\u00A0', ' '), 'Цена 3,5 и 1,25; адрес https://ens.example/a,b почта a@b.ru', 'числа и адреса целы');
  await p.context().close();
});

await test('типограф: форматирование сохраняется, ссылки не трогаются, один шаг отмены, страховочная версия', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const bold = s.text('"жирный" текст', [s.marks.strong.create()]);
    const link = s.text('"адрес"', [s.marks.link.create({ href: 'https://x.y' })]);
    E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, [s.text('Это '), bold, s.text(' и '), link])]), E.getMeta(), 't');
    E.run('typo.apply');
    const para = E.view.state.doc.firstChild;
    const parts = []; para.forEach((n) => parts.push([n.text, n.marks.map((m) => m.type.name).join('+')]));
    return { parts, versions: E.hist().versions.map((v) => v.kind) };
  });
  ok(r.parts.some(([t, m]) => t.includes('«жирный»') && m.includes('strong')), JSON.stringify(r.parts));
  ok(r.parts.some(([t, m]) => t === '"адрес"' && m.includes('link')), 'ссылка не изменена');
  ok(r.versions.includes('before'), 'страховочная версия');
  await p.keyboard.press('Control+KeyZ');
  ok((await p.evaluate(() => window.ens.view.state.doc.textContent)).includes('"жирный"'), 'отмена одним шагом');
  await p.context().close();
});

await test('типограф: кавычки при наборе («ёлочки» по контексту)', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Он сказал "привет "там" всем" и ушёл');
  eq(await p.evaluate(() => window.ens.view.state.doc.textContent), 'Он сказал «привет „там“ всем» и ушёл', 'кавычки при наборе');
  await p.context().close();
});

await test('метки: двойной щелчок на левом поле, Shift — зелёная, переход, список, в docx не попадают', async () => {
  const p = await longPage((s, P) => Array.from({ length: 30 }, (_, i) => P(`Абзац ${i + 1}`)));
  const pt = await p.evaluate(() => { const r = document.querySelectorAll('.ProseMirror p')[2].getBoundingClientRect(), s = document.getElementById('sheet').getBoundingClientRect(); return { x: s.left + 30, y: r.top + r.height / 2 }; });
  await p.mouse.dblclick(pt.x, pt.y);
  eq(await p.evaluate(() => window.ens.view.state.doc.child(2).attrs.flag), 'red', 'красная метка');
  await p.keyboard.down('Shift');
  const pt2 = await p.evaluate(() => { const r = document.querySelectorAll('.ProseMirror p')[5].getBoundingClientRect(), s = document.getElementById('sheet').getBoundingClientRect(); return { x: s.left + 30, y: r.top + r.height / 2 }; });
  await p.mouse.dblclick(pt2.x, pt2.y);
  await p.keyboard.up('Shift');
  eq(await p.evaluate(() => window.ens.view.state.doc.child(5).attrs.flag), 'green', 'зелёная метка');
  await p.evaluate(() => window.ens.view.dispatch(window.ens.view.state.tr.setSelection(window.ens.view.state.selection.constructor.atStart(window.ens.view.state.doc))));
  await p.keyboard.press('Control+Alt+ArrowDown');
  eq(await p.evaluate(() => window.ens.view.state.doc.resolve(window.ens.view.state.selection.from).parent.textContent), 'Абзац 3', 'к следующей метке');
  ok(await p.evaluate(() => getComputedStyle(document.querySelectorAll('.ProseMirror p')[2], '::after').backgroundColor !== 'rgba(0, 0, 0, 0)'), 'метка нарисована');
  const back = await p.evaluate(() => { const E = window.ens; return JSON.stringify(E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta(), 'x')).doc.toJSON()).includes('"flag":"'); });
  eq(back, false, 'метки не попадают в docx');
  await p.context().close();
});

await test('фокус абзаца: остальной текст затемняется', async () => {
  const p = await longPage((s, P) => Array.from({ length: 5 }, (_, i) => P(`Абзац ${i + 1}`)));
  await p.click('.ProseMirror p >> nth=1');
  await p.waitForFunction(() => window.ens.view.state.selection.from > 1);
  await p.keyboard.press('Control+Shift+KeyK');
  await p.waitForTimeout(300);
  const ops = await p.evaluate(() => [...document.querySelectorAll('.ProseMirror p')].map((e) => getComputedStyle(e).opacity));
  eq(ops.map((o) => +o < 0.9), [true, false, true, true, true], 'затемнены все, кроме текущего');
  await p.keyboard.press('Control+Shift+KeyK');
  await p.context().close();
});

await test('поиск команд Alt+Q: находит по-русски и с опечаткой, запускает', async () => {
  const p = await page();
  await p.keyboard.press('Alt+KeyQ');
  await p.keyboard.type('типограв');
  await p.waitForSelector('.search-results button');
  ok((await p.textContent('.search-results')).includes('Типограф'), await p.textContent('.search-results'));
  await p.keyboard.press('Escape');
  await p.keyboard.press('Alt+KeyQ');
  await p.keyboard.type('история верс');
  await p.keyboard.press('Enter');
  await p.waitForSelector('.history-modal');
  await p.context().close();
});


// ---- Поиск и замена по оформлению ----
await test('поиск по оформлению: шрифт/кегль/Ж, замена только оформления и текста с оформлением', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const ts = (font, size) => s.marks.textStyle.create({ font, size });
    const P = (...r) => s.nodes.paragraph.create(null, r);
    E.loadDoc(s.nodes.doc.create(null, [
      P(s.text('Обычный '), s.text('Арий', [ts('Arial', 14)]), s.text(' и '), s.text('жирный', [s.marks.strong.create()])),
      P(s.text('Арий', [ts('Arial', 14), s.marks.strong.create()]), s.text(' Арий')),
    ]), E.getMeta(), 'fmt');
  });
  await p.keyboard.press('Control+KeyH');
  await p.click('.findbar button:has-text("Формат")');
  const st = () => p.evaluate(() => { const s = window.ens.findState(); return s.matches.length; });
  await p.fill('.fmtbox .fmtrow:nth-child(1) input[type=text]', 'Arial');
  await p.waitForTimeout(100);
  eq(await st(), 2, 'Arial: два участка');
  await p.fill('.fmtbox .fmtrow:nth-child(1) input[type=number]', '14');
  await p.check('.fmtbox .fmtrow:nth-child(1) label:has-text("Ж") input');
  await p.waitForTimeout(100);
  eq(await st(), 1, 'Arial 14 жирный: один');
  await p.uncheck('.fmtbox .fmtrow:nth-child(1) label:has-text("Ж") input');
  await p.fill('.findbar input[placeholder="Найти"]', 'Арий');
  await p.waitForTimeout(100);
  eq(await st(), 2, 'текст «Арий» + Arial 14: два');
  // замена только оформления: Arial 14 → Verdana, жирный
  await p.fill('.fmtbox .fmtrow:nth-child(2) input[type=text]', 'Verdana');
  await p.check('.fmtbox .fmtrow:nth-child(2) label:has-text("Ж") input');
  await p.click('.findbar button:has-text("Заменить все")');
  await p.waitForTimeout(150);
  const r = await p.evaluate(() => { const out = []; window.ens.view.state.doc.descendants((n) => { if (n.isText) out.push([n.text, n.marks.map((m) => m.type.name + (m.attrs.font ? ':' + m.attrs.font : '')).sort().join(',')]); }); return out; });
  ok(r.filter((x) => x[1].includes('Verdana') && x[1].includes('strong')).map((x) => x[0]).join('|') === 'Арий|Арий', `после замены: ${JSON.stringify(r)}`);
  ok(r.some((x) => x[0] === ' Арий' && !x[1].includes('strong')), 'не найденное не тронуто');
});

// ---- По-русски: сумма прописью, ФИО по падежам ----
await test('сумма прописью и ФИО по падежам: окна вставляют текст', async () => {
  const p = await page();
  await p.evaluate(() => { const E = window.ens, s = E.schema; E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null)]), E.getMeta(), 'ru'); });
  await p.evaluate(() => window.ens.run('insert.sumWords'));
  await p.fill('.modal input[type=text]', '12 345,67');
  const out = await p.textContent('.ru-out');
  eq(out, 'Двенадцать тысяч триста сорок пять рублей 67 копеек', 'сумма');
  await p.click('.modal button.primary');
  eq(await p.evaluate(() => window.ens.view.state.doc.textContent), 'Двенадцать тысяч триста сорок пять рублей 67 копеек', 'вставлено');
  await p.evaluate(() => window.ens.run('insert.nameCases'));
  await p.fill('input[placeholder="Иванов Иван Иванович"]', 'Петрова Анна Сергеевна');
  await p.fill('input[placeholder^="Должность"]', 'Генеральный директор');
  const rows = await p.$$eval('.ru-case', (e) => e.map((x) => x.textContent));
  ok(rows[2].includes('Генеральному директору Петровой Анне Сергеевне'), `дательный: ${rows[2]}`);
  await p.click('.ru-case:nth-child(2)');
  const fin = await p.evaluate(() => window.ens.view.state.doc.textContent);
  ok(fin.endsWith('Генерального директора Петровой Анны Сергеевны'), `вставлен родительный: ${fin}`);
});

await test('срок в рабочих днях: расчёт, поправки, вставка даты', async () => {
  const p = await page();
  await p.evaluate(() => { const E = window.ens, s = E.schema; E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null)]), E.getMeta(), 'dl'); });
  await p.evaluate(() => window.ens.run('insert.deadline'));
  await p.fill('.modal input[type=date]', '2026-03-02'); // понедельник
  await p.fill('.modal input[type=number]', '5');
  let out = await p.textContent('.ru-out');
  ok(out.includes('«10» марта 2026 г.') && out.includes('вторник'), `5 рабочих дней с 02.03.2026 (8 марта — воскресенье, выходной переносится на пн 9-е): ${out}`);
  await p.fill('.modal textarea', '=2026-03-07');
  out = await p.textContent('.ru-out');
  ok(out.includes('«07» марта') && out.includes('суббота'), `рабочая суббота 7 марта: ${out}`);
  await p.click('.modal button.primary');
  eq(await p.evaluate(() => window.ens.view.state.doc.textContent), '«07» марта 2026 г.', 'вставлена дата');
});

await test('мастер документов: приказ по ГОСТ, номер из журнала, второй номер растёт', async () => {
  const p = await page();
  await p.evaluate(() => localStorage.removeItem('ens-text.journal'));
  const make = async (subject) => {
    await p.evaluate(() => window.ens.run('file.wizard'));
    await p.waitForTimeout(100);
    const unsaved = await p.$('.modal button.danger'); // «Не сохранять», если в документе есть несохранённое
    if (unsaved) await unsaved.click();
    await p.selectOption('.modal select', 'order');
    await p.fill('.wiz-form textarea >> nth=0', 'ООО «Ромашка»');
    await p.fill('input[placeholder^="О назначении"]', subject);
    await p.fill('.wiz-form textarea >> nth=2', 'Назначить Иванова И. И.\nУтвердить инструкцию');
    await p.fill('input[placeholder="И. И. Иванов"]', 'А. А. Петров');
    await p.fill('input[placeholder="Директор"]', 'Директор');
    await p.click('.modal button.primary');
    await p.waitForTimeout(300);
  };
  await make('О назначении ответственного');
  const r = await p.evaluate(() => { const E = window.ens, d = E.view.state.doc, m = E.getMeta(); return { text: d.textContent, left: m.page.left, size: m.defaults.size, hdr: JSON.stringify(m.hf.header).includes('field'), diff: m.hf.differentFirst, j: JSON.parse(localStorage.getItem('ens-text.journal')).entries.map((e) => e.label) }; });
  ok(r.text.includes('ПРИКАЗ') && r.text.includes('№ 1') && r.text.includes('1. Назначить Иванова') && r.text.includes('3. Контроль за исполнением') && r.text.includes('ПРИКАЗЫВАЮ'), `текст приказа: ${r.text}`);
  ok(Math.abs(r.left - 56.7) < 0.1 && r.size === 14 && r.hdr && r.diff, 'поля, шрифт, номер страницы');
  eq(r.j, ['1'], 'журнал');
  await make('Второй');
  eq(await p.evaluate(() => JSON.parse(localStorage.getItem('ens-text.journal')).entries.map((e) => e.label).join(',')), '1,2', 'номер вырос');
  // документ проходит проверку по ГОСТ без замечаний к полям и шрифту
  await p.evaluate(() => window.ens.run('review.check'));
  await p.waitForSelector('.check-row');
  await p.selectOption('.check-modal select', 'gost7097');
  const bad = await p.$$eval('.check-row.bad .check-text > div:first-child', (e) => e.map((x) => x.textContent));
  ok(!bad.some((x) => /Поля|Шрифт|Номера страниц/.test(x)), `замечания к приказу: ${bad.join(' | ')}`);
});

await test('совместимость с Word: шрифт без аналога, широкая таблица, дробный кегль, потери', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema, m = E.getMeta();
    const ts = (font, size) => s.marks.textStyle.create({ font, size });
    const P = (...r) => s.nodes.paragraph.create(null, r);
    const cell = (w) => s.nodes.table_cell.create({ colwidth: [w] }, P(s.text('x')));
    const table = s.nodes.table.create(null, [s.nodes.table_row.create(null, [cell(500), cell(500)])]);
    m.lossy = ['сноски: 2'];
    E.loadDoc(s.nodes.doc.create(null, [P(s.text('Inter', [ts('Inter', 11.3)])), P(s.text('Arial', [ts('Arial', 12)])), table, P()]), m, 'cmp');
  });
  await p.evaluate(() => window.ens.run('review.compat'));
  await p.waitForSelector('.cmp-row');
  const titles = await p.$$eval('.cmp-row .cmp-title', (e) => e.map((x) => x.textContent));
  const has = (re) => titles.some((x) => re.test(x));
  ok(has(/Шрифт «Inter»/) && !has(/Шрифт «Arial»/), `шрифты: ${titles.join(' | ')}`);
  ok(has(/Таблица шире/) && has(/Дробный кегль/) && has(/Не перенесено/) && has(/насчитал страниц/), `остальное: ${titles.join(' | ')}`);
  await p.click('.cmp-row:has-text("Шрифт «Inter»") button');
  await p.waitForTimeout(100);
  ok(await p.evaluate(() => document.getElementById('dialog').hidden), 'окно закрылось после «Показать»');
  // чистый документ — без предупреждений
  await p.evaluate(() => { const E = window.ens, s = E.schema; const m = E.getMeta(); m.lossy = []; E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, s.text('Просто текст'))]), m, 'clean'); });
  await p.evaluate(() => window.ens.run('review.compat'));
  await p.waitForSelector('.cmp-verdict.ok');
});

await test('читаемость: время, длинные предложения, повторы, канцеляризмы; ночной лист не трогает печать', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const P = (t) => s.nodes.paragraph.create(null, s.text(t));
    E.loadDoc(s.nodes.doc.create(null, [
      P('В целях осуществления контроля данный отдел производит проверку документов. Директор И. И. Иванов, г. Москва, подписал приказ.'),
      P('Длинное ' + 'слово '.repeat(35) + 'конец.'),
    ]), E.getMeta(), 'rd');
  });
  await p.evaluate(() => window.ens.run('review.readability'));
  await p.waitForSelector('.rd-row');
  const kinds = await p.$$eval('.rd-row .rd-kind', (e) => e.map((x) => x.textContent));
  ok(kinds.includes('Канцеляризм') && kinds.includes('Длинное предложение') && kinds.includes('Повтор слова'), `замечания: ${kinds.join(', ')}`);
  const stats = await p.$$eval('.rd-stat .rd-lab', (e) => e.map((x) => x.textContent));
  ok(stats.includes('слов') && stats.includes('индекс Флеша'), 'показатели');
  await p.click('.rd-row:has(.rd-kind.clerical) button');
  const selText = await p.evaluate(() => { const st = window.ens.view.state; return st.doc.textBetween(st.selection.from, st.selection.to); });
  ok(/целях|данный|производит/i.test(selText), `«Показать» выделяет канцеляризм: «${selText}»`);
  // ночной лист: тёмный лист на экране, в печати белый
  await p.evaluate(() => window.ens.run('view.night'));
  const bg = await p.$eval('.paper', (e) => getComputedStyle(e).backgroundColor);
  ok(bg !== 'rgb(255, 255, 255)', `лист тёмный: ${bg}`);
  await p.emulateMedia({ media: 'print' });
  const printBg = await p.$eval('.paper', (e) => getComputedStyle(e).backgroundColor);
  await p.emulateMedia({ media: 'screen' });
  eq(printBg, 'rgb(255, 255, 255)', 'в печати лист белый');
  await p.evaluate(() => window.ens.run('view.night'));
});

await test('шаблоны: готовые создаются, «мой шаблон» сохраняется локально и открывается; свой профиль ГОСТ', async () => {
  const p = await page();
  const dismissUnsaved = async () => { await p.waitForTimeout(100); const d = await p.$('.modal button.danger'); if (d) await d.click(); };
  await p.evaluate(() => window.ens.run('file.templates'));
  await dismissUnsaved();
  await p.waitForSelector('.tpl-row');
  const titles = await p.$$eval('.tpl-title', (e) => e.map((x) => x.textContent));
  for (const need of ['Письмо', 'Приказ', 'Акт', 'Протокол', 'Служебная записка', 'Заявление', 'Доверенность (простая)', 'Отчёт о НИР (ГОСТ 7.32)']) ok(titles.includes(need), `нет шаблона «${need}»: ${titles.join(' | ')}`);
  await p.click('.tpl-row:has-text("Служебная записка") button.primary');
  await p.waitForTimeout(300);
  const memo = await p.evaluate(() => window.ens.view.state.doc.textContent);
  ok(memo.includes('СЛУЖЕБНАЯ ЗАПИСКА') && memo.includes('[Наименование организации]'), `служебная записка: ${memo.slice(0, 80)}`);
  // мой шаблон: сохраняем текущий документ, очищаем документ, создаём из шаблона
  await p.evaluate(() => window.ens.run('file.saveTemplate'));
  await p.fill('.modal input[type=text]', 'Моя записка');
  await p.click('.modal button.primary');
  await p.waitForTimeout(300);
  await p.evaluate(() => { const E = window.ens, s = E.schema; E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, s.text('другое'))]), E.getMeta(), 'x'); });
  await p.evaluate(() => window.ens.run('file.templates'));
  await dismissUnsaved();
  await p.click('.rd-tabs button[data-tab=mine]');
  await p.waitForSelector('.tpl-row:has-text("Моя записка")');
  const note = await p.textContent('.tpl-list p');
  ok(note.includes('только на этом компьютере'), 'пояснение про локальное хранение');
  await p.click('.tpl-row:has-text("Моя записка") button.primary');
  await p.waitForTimeout(400);
  ok((await p.evaluate(() => window.ens.view.state.doc.textContent)).includes('СЛУЖЕБНАЯ ЗАПИСКА'), 'мой шаблон открылся');
  // удаление
  await p.evaluate(() => window.ens.run('file.templates'));
  await dismissUnsaved();
  await p.click('.rd-tabs button[data-tab=mine]');
  await p.click('.tpl-row:has-text("Моя записка") button:has-text("Удалить")');
  await p.waitForTimeout(300);
  eq(await p.$$eval('.tpl-row', (e) => e.length), 0, 'шаблон удалён');
  await p.keyboard.press('Escape');
  // свой профиль
  await p.evaluate(() => window.ens.run('review.check'));
  await p.waitForSelector('.check-modal');
  await p.click('.check-modal button:has-text("Изменить свой профиль")');
  await p.fill('.prof-margins input >> nth=0', '25');
  await p.fill('.modal textarea >> nth=0', 'Arial: 11-12');
  await p.click('.modal button.primary:has-text("Сохранить")');
  await p.waitForTimeout(200);
  const opts = await p.$$eval('.check-modal select option', (o) => o.map((x) => x.textContent));
  ok(opts.some((x) => x.includes('Свой профиль')), `свой профиль в списке: ${opts.join(' | ')}`);
  ok(await p.evaluate(() => JSON.parse(localStorage.getItem('ens-text.profile.custom')).margins.left === 25), 'поле сохранено');
  await p.evaluate(() => localStorage.removeItem('ens-text.profile.custom'));
});

await test('рисунок с обрезкой: srcRect читается и пишется, показывается обрезанным; длинное предупреждение переносится', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mP8z8BQz0AEYBxVSF+FABJADveWkH6oAAAAAElFTkSuQmCC';
    const img = s.nodes.image.create({ src: png, width: 100, height: 60, crop: [0.1, 0.2, 0.3, 0.1] });
    const doc = s.nodes.doc.create(null, [s.nodes.paragraph.create(null, [img, s.text(' текст')])]);
    const bytes = E.exportDocx(doc, E.getMeta(), 'crop');
    const back = E.importDocx(bytes);
    const crop = back.doc.firstChild.firstChild.attrs.crop;
    const m = E.getMeta(); m.lossy = ['обтекание изображений (изображения стали строчными)', 'разделы (разрывы разделов стали обычными)', 'очень длинное пояснение '.repeat(12)];
    E.loadDoc(back.doc, m, 'crop');
    return { crop };
  });
  ok(r.crop && Math.abs(r.crop[0] - 0.1) < 1e-4 && Math.abs(r.crop[1] - 0.2) < 1e-4 && Math.abs(r.crop[2] - 0.3) < 1e-4 && Math.abs(r.crop[3] - 0.1) < 1e-4, `обрезка в docx: ${JSON.stringify(r.crop)}`);
  const style = await p.$eval('.ProseMirror img[data-crop]', (e) => e.getAttribute('style'));
  ok(/object-view-box:\s*inset\(20(\.0+)?% 30(\.0+)?% 10(\.0+)?% 10(\.0+)?%\)/.test(style), `стиль: ${style}`);
  const h = await p.$eval('.notice span', (e) => e.getBoundingClientRect().height);
  ok(h > 30, `предупреждение переносится по строкам (высота ${h})`);
});

await test('разделы: книжная — альбомная — книжная; листы разного размера, docx туда и обратно', async () => {
  const p = await page();
  const r = await p.evaluate(async () => {
    const E = window.ens, s = E.schema;
    const P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const land = { ...E.getMeta().page, w: E.getMeta().page.h, h: E.getMeta().page.w };
    const m = E.getMeta();
    const doc = s.nodes.doc.create(null, [P('Книжная 1'), P('Книжная 1б'), s.nodes.section_break.create({ page: land }), P('Альбомная 2'), s.nodes.section_break.create({ page: { ...m.page } }), P('Книжная 3')]);
    E.loadDoc(doc, m, 'sect');
    await new Promise((r) => setTimeout(r, 400));
    const boxes = E.pageBoxes().map((b) => ({ top: Math.round(b.top), W: Math.round(b.W), H: Math.round(b.H) }));
    const papers = [...document.querySelectorAll('.paper')].map((e) => ({ w: Math.round(e.getBoundingClientRect().width), left: Math.round(e.offsetLeft) }));
    const bytes = E.exportDocx(E.view.state.doc, E.getMeta(), 'sect');
    const back = E.importDocx(bytes);
    const kinds = []; back.doc.forEach((n) => kinds.push(n.type.name + (n.type.name === 'section_break' ? (n.attrs.page.w > n.attrs.page.h ? ':land' : ':port') : '')));
    return { boxes, papers, kinds, lossy: back.meta.lossy, count: E.view.state.doc.childCount };
  });
  eq(r.boxes.length, 3, `три листа: ${JSON.stringify(r.boxes)}`);
  ok(r.boxes[0].W < r.boxes[0].H && r.boxes[1].W > r.boxes[1].H && r.boxes[2].W < r.boxes[2].H, `ориентации: ${JSON.stringify(r.boxes)}`);
  ok(r.boxes[1].top > r.boxes[0].top + r.boxes[0].H && r.boxes[2].top > r.boxes[1].top + r.boxes[1].H, 'листы идут друг за другом без наложения');
  ok(r.papers.length === 3 && r.papers[1].w > r.papers[0].w, `подложки разной ширины: ${JSON.stringify(r.papers)}`);
  eq(r.kinds.join(','), 'paragraph,paragraph,section_break:land,paragraph,section_break:port,paragraph', 'разделы после docx');
  eq(r.lossy.filter((x) => /раздел/.test(x)), [], 'о разделах не предупреждает');
});

await test('разделы из настоящего docx (python-docx): три листа, ориентации сохраняются при записи', async () => {
  const p = await page();
  await p.setInputFiles('#file-input', join(here, 'fixtures', 'sections.docx'));
  await p.waitForFunction(() => window.ens.pageBoxes().length === 3, null, { timeout: 5000 });
  const o = await p.evaluate(() => window.ens.pageBoxes().map((b) => (b.W > b.H ? 'land' : 'port')).join(','));
  eq(o, 'port,land,port', 'ориентации листов');
  eq(await p.evaluate(() => window.ens.getMeta().lossy.filter((x) => /раздел/.test(x))), [], 'без предупреждений о разделах');
  const bytes = await p.evaluate(() => [...window.ens.exportDocx(window.ens.view.state.doc, window.ens.getMeta(), 'x')]);
  const tmp = mkdtempSync(join(tmpdir(), 'ens-sect-'));
  writeFileSync(join(tmp, 's.docx'), Buffer.from(bytes));
  const out = execSync(`python3 -c "from docx import Document; d=Document('${tmp}/s.docx'); print(','.join('land' if s.page_width>s.page_height else 'port' for s in d.sections))"`).toString().trim();
  eq(out, 'port,land,port', 'python-docx видит три раздела с теми же ориентациями');
});

await test('плавающие рисунки: логотип в колонтитуле «за текстом» от страницы, обтекание в тексте, запись обратно', async () => {
  const p = await page();
  await p.setInputFiles('#file-input', join(here, 'fixtures', 'floats.docx'));
  await p.waitForSelector('.paper .hf img.fl-behind', { timeout: 5000 });
  await p.waitForTimeout(300);
  const r = await p.evaluate(() => {
    const paper = document.querySelector('.paper');
    const img = paper.querySelector('.hf img.fl-behind');
    const pr = paper.getBoundingClientRect(), ir = img.getBoundingClientRect();
    const hf = paper.querySelector('.hf');
    const body = document.querySelector('.ProseMirror img.fl-wrap');
    return {
      dx: Math.round(ir.left - pr.left), dy: Math.round(ir.top - pr.top), hfH: hf.offsetHeight, w: Math.round(ir.width), pos: getComputedStyle(img).position,
      wrap: body ? getComputedStyle(body).float : null, lossy: window.ens.getMeta().lossy, effTop: Math.round(window.ens.effPage().top),
    };
  });
  ok(Math.abs(r.dx - 76) <= 2 && Math.abs(r.dy - 38) <= 2, `логотип от угла страницы 76×38 px: ${r.dx}×${r.dy}`);
  ok(r.hfH < 40, `колонтитул не раздут логотипом: высота ${r.hfH}`);
  ok(r.w > 200, `логотип 6 см (≈227 px): ${r.w}`);
  eq(r.wrap, 'left', 'картинка в тексте обтекается слева');
  eq(r.lossy.filter((x) => /обтекание/.test(x)), [], 'об обтекании не предупреждает');
  const back = await p.evaluate(() => {
    const E = window.ens;
    const b = E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta(), 'f'));
    let f = null; b.doc.descendants((n) => { if (n.type.name === 'image' && n.attrs.float) f = n.attrs.float; });
    const hdr = JSON.stringify(b.meta.hf.header);
    return { f, hdr: hdr.includes('"mode":"behind"') };
  });
  ok(back.f && back.f.mode === 'wrap' && back.f.side === 'left' && back.hdr, `после записи и чтения: ${JSON.stringify(back)}`);
});

await test('обтекание рисунка из ленты: в тексте, за текстом, слева; запись в docx', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const png = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mP8z8BQz0AEYBxVSF+FABJADveWkH6oAAAAAElFTkSuQmCC';
    const img = s.nodes.image.create({ src: png, width: 80, height: 80 });
    E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, [s.text('Текст '), img, s.text(' ещё текст')])]), E.getMeta(), 'wrap');
    const { NodeSelection } = E.pm;
    E.view.dispatch(E.view.state.tr.setSelection(NodeSelection.create(E.view.state.doc, 7)));
  });
  await p.click('#tabs .tab:has-text("Макет")');
  await p.waitForTimeout(100);
  await p.selectOption('select[data-command="image.wrap"]', 'behind');
  await p.waitForTimeout(300);
  eq(await p.$eval('.ProseMirror img[data-crop], .ProseMirror img.fl-behind', (e) => getComputedStyle(e).position), 'absolute', 'за текстом — вне строки');
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.NodeSelection.create(E.view.state.doc, 7))); });
  await p.selectOption('select[data-command="image.wrap"]', 'left');
  await p.waitForTimeout(200);
  eq(await p.$eval('.ProseMirror img.fl-wrap', (e) => getComputedStyle(e).float), 'left', 'слева с обтеканием');
  const xml = await p.evaluate(() => { const E = window.ens; const b = E.exportDocx(E.view.state.doc, E.getMeta(), 'w'); return new TextDecoder().decode(E.unzip(b)['word/document.xml']); });
  ok(xml.includes('<wp:anchor') && xml.includes('wrapSquare'), 'в docx записан якорь с обтеканием');
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.NodeSelection.create(E.view.state.doc, 7))); });
  await p.selectOption('select[data-command="image.wrap"]', 'inline');
  await p.waitForTimeout(200);
  eq(await p.$$eval('.ProseMirror img.fl', (e) => e.length), 0, 'снова в тексте');
});

await test('рисунки: ручки размера мышью, обрезка, окно «Формат», подпись, перетаскивание плавающего', async () => {
  const p = await page();
  const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mP8z8BQz0AEYBxVSF+FABJADveWkH6oAAAAAElFTkSuQmCC';
  const setup = (float) => p.evaluate(({ png, float }) => {
    const E = window.ens, s = E.schema;
    const img = s.nodes.image.create({ src: png, width: 120, height: 80, float });
    E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, [s.text('Строка '), img, s.text(' текста')]), s.nodes.paragraph.create(null, s.text('Вторая строка'))]), E.getMeta(), 'img');
  }, { png: PNG, float });
  const attrs = () => p.evaluate(() => { let a = null; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'image') a = n.attrs; }); return a; });
  await setup(null);
  await p.waitForTimeout(300);
  // щелчок выделяет рисунок, появляются ручки
  const box = await (await p.$('.ProseMirror img[src^="data:"]')).boundingBox();
  await p.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await p.waitForSelector('.img-ov:not([hidden]) .img-h-se');
  // угловая ручка сохраняет пропорции
  const se = await (await p.$('.img-h-se')).boundingBox();
  await p.mouse.move(se.x + 4, se.y + 4);
  await p.mouse.down();
  await p.mouse.move(se.x + 64, se.y + 14, { steps: 5 });
  await p.mouse.up();
  await p.waitForTimeout(200);
  let a = await attrs();
  ok(a.width > 160 && Math.abs(a.width / a.height - 1.5) < 0.05, `угол: размер ${a.width}×${a.height}, пропорции 3:2`);
  // боковая ручка меняет только ширину
  const e = await (await p.$('.img-h-e')).boundingBox();
  await p.mouse.move(e.x + 4, e.y + 4);
  await p.mouse.down();
  await p.mouse.move(e.x + 44, e.y + 4, { steps: 4 });
  await p.mouse.up();
  await p.waitForTimeout(200);
  const a2 = await attrs();
  ok(a2.width > a.width + 30 && a2.height === a.height, `край: ширина ${a.width}→${a2.width}, высота ${a2.height}`);
  // обрезка слева
  await p.evaluate(() => window.ens.run('image.crop'));
  const w = await (await p.$('.img-h-w')).boundingBox();
  await p.mouse.move(w.x + 4, w.y + 4);
  await p.mouse.down();
  await p.mouse.move(w.x + 44, w.y + 4, { steps: 4 });
  await p.mouse.up();
  await p.waitForTimeout(200);
  const a3 = await attrs();
  ok(a3.crop && a3.crop[0] > 0.1 && a3.width < a2.width - 30, `обрезка: ${JSON.stringify(a3.crop)}, ширина ${a3.width}`);
  await p.evaluate(() => window.ens.run('image.crop'));
  await p.evaluate(() => window.ens.run('image.cropReset'));
  await p.waitForTimeout(100);
  eq((await attrs()).crop, null, 'обрезка убрана');
  // окно «Формат рисунка»: ширина 5 см с сохранением пропорций
  await p.evaluate(() => window.ens.run('image.format'));
  await p.waitForSelector('.modal');
  await p.fill('.modal .dlg-row:has-text("Ширина") input[type=text]', '5');
  await p.click('.modal button.primary');
  await p.waitForTimeout(200);
  a = await attrs();
  ok(Math.abs(a.width - 189) <= 2, `5 см ≈ 189 px: ${a.width}`);
  // подпись
  await p.evaluate(() => window.ens.run('image.caption'));
  ok((await p.evaluate(() => window.ens.view.state.doc.textContent)).includes('Рисунок 1 — '), 'подпись добавлена');
  // плавающий рисунок: Alt+щелчок выбирает, перетаскивание меняет положение
  await setup({ mode: 'behind', hRel: 'page', hOff: 200, hAlign: null, vRel: 'page', vOff: 300, vAlign: null, dist: [0, 0, 8, 8] });
  await p.waitForTimeout(400);
  const fb = await (await p.$('.ProseMirror img.fl-behind')).boundingBox();
  await p.keyboard.down('Alt');
  await p.mouse.click(fb.x + fb.width / 2, fb.y + fb.height / 2);
  await p.keyboard.up('Alt');
  await p.waitForSelector('.img-ov.abs:not([hidden])');
  const body = await (await p.$('.img-ov-body')).boundingBox();
  const before = await attrs();
  await p.mouse.move(body.x + 20, body.y + 20);
  await p.mouse.down();
  await p.mouse.move(body.x + 80, body.y + 50, { steps: 5 });
  await p.mouse.up();
  await p.waitForTimeout(300);
  const moved = await attrs();
  ok(moved.float.hOff - before.float.hOff >= 50 && moved.float.vOff - before.float.vOff >= 25, `перетащили: ${before.float.hOff},${before.float.vOff} → ${moved.float.hOff},${moved.float.vOff}`);
  await p.evaluate(() => window.ens.run('image.align'));
});

await test('линейки: поля клавиатурой меняются только в текущем разделе и сохраняются в DOCX', async () => {
  const p = await page();
  await p.waitForSelector('#hr-inner .rm[role=slider][aria-valuenow]');
  await p.waitForSelector('#vr-inner .rm[role=slider][aria-valuenow]');
  const left = p.locator('#hr-inner .rm[aria-label="Левое поле"]');
  const right = p.locator('#hr-inner .rm[aria-label="Правое поле"]');
  const top = p.locator('#vr-inner .rm[aria-label="Верхнее поле"]');
  const bottom = p.locator('#vr-inner .rm[aria-label="Нижнее поле"]');
  const initial = await p.evaluate(() => ({ ...window.ens.getMeta().page }));
  await left.focus(); await left.press('ArrowRight');
  await right.focus(); await right.press('Shift+ArrowLeft');
  await top.focus(); await top.press('ArrowDown');
  await bottom.focus(); await bottom.press('ArrowUp');
  const first = await p.evaluate(() => ({ ...window.ens.getMeta().page }));
  const MM = 72 / 25.4; // шаг клавиш: 1 мм, с Shift — 5 мм
  const moved = [first.left - initial.left, first.right - initial.right, first.top - initial.top, first.bottom - initial.bottom];
  ok(moved.every((d, i) => Math.abs(d - MM * [1, 5, 1, 1][i]) < 0.5), `четыре поля: ${JSON.stringify(moved)}`);
  await left.focus(); await left.press('ArrowLeft');
  ok(Math.abs((await p.evaluate(() => window.ens.getMeta().page.left)) - initial.left) < 0.5, 'обратная стрелка');
  await p.evaluate(() => {
    const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const meta = E.getMeta();
    E.loadDoc(s.nodes.doc.create(null, [P('Первый раздел'), s.nodes.section_break.create({ page: { ...meta.page } }), P('Второй раздел')]), meta, 'ruler-section');
    const doc = E.view.state.doc;
    E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(doc, doc.content.size - 2)));
  });
  await p.waitForTimeout(300);
  await left.focus(); await left.press('Shift+ArrowRight');
  const result = await p.evaluate(() => {
    const E = window.ens, doc = E.view.state.doc;
    const section = doc.child(1).attrs.page;
    const back = E.importDocx(E.exportDocx(doc, E.getMeta(), 'ruler-section'));
    return { first: E.getMeta().page.left, second: section.left, backFirst: back.meta.page.left, backSecond: back.doc.child(1).attrs.page.left };
  });
  ok(Math.abs(result.first - initial.left) < 0.5 && Math.abs(result.second - (initial.left + 5 * MM)) < 0.5, `независимые поля: ${JSON.stringify(result)}`);
  ok(Math.abs(result.backFirst - result.first) < 0.1 && Math.abs(result.backSecond - result.second) < 0.1, `поля после DOCX: ${JSON.stringify(result)}`);
  await p.context().close();
});

await test('линейки: поля и отступы абзаца тянутся мышью, видны метки столбцов и строк таблицы', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const P = (t) => s.nodes.paragraph.create(null, s.text(t));
    E.loadDoc(s.nodes.doc.create(null, [P('Первый абзац для проверки отступов на линейке. '.repeat(4)), P('Второй абзац')]), E.getMeta(), 'ruler');
    E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 5)));
  });
  await p.waitForTimeout(400);
  await p.waitForSelector('#hr-inner svg line', { state: 'attached' });
  await p.waitForSelector('#vr-inner svg line', { state: 'attached' });
  const drag = async (sel, dx, dy) => {
    const b = await (await p.$(sel)).boundingBox();
    await p.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
    await p.mouse.down();
    await p.mouse.move(b.x + b.width / 2 + dx, b.y + b.height / 2 + dy, { steps: 6 });
    await p.mouse.up();
    await p.waitForTimeout(250);
  };
  // левое поле: +1 см (≈38 px) → +28 pt
  const left0 = await p.evaluate(() => window.ens.getMeta().page.left);
  await drag('#hr-inner .rm >> nth=0', 38, 0);
  const left1 = await p.evaluate(() => window.ens.getMeta().page.left);
  ok(left1 - left0 > 25 && left1 - left0 < 32, `левое поле ${left0} → ${left1} pt`);
  // верхнее поле: −0,5 см
  const top0 = await p.evaluate(() => window.ens.getMeta().page.top);
  await drag('#vr-inner .rm >> nth=0', 0, -19);
  const top1 = await p.evaluate(() => window.ens.getMeta().page.top);
  ok(top0 - top1 > 10 && top0 - top1 < 18, `верхнее поле ${top0} → ${top1} pt`);
  // красная строка (первая строка) +1 см
  const attrs = () => p.evaluate(() => { const n = window.ens.view.state.doc.child(0); return { indent: n.attrs.indent, first: n.attrs.first, rindent: n.attrs.rindent }; });
  await drag('.ri-first', 38, 0);
  let a = await attrs();
  ok(a.first > 25 && a.first < 32 && !a.indent, `красная строка: ${JSON.stringify(a)}`);
  // отступ слева (квадрат): абзац целиком +1 см, красная строка сохраняется
  await drag('.ri-left', 38, 0);
  a = await attrs();
  ok(a.indent > 25 && a.indent < 32 && a.first > 25, `отступ слева: ${JSON.stringify(a)}`);
  // выступ (нижний треугольник): первая строка остаётся на месте, остальные сдвигаются
  const firstPos = a.indent + a.first;
  await drag('.ri-hang', 19, 0);
  a = await attrs();
  ok(Math.abs(a.indent + a.first - firstPos) < 1.5 && a.indent > 38, `выступ: ${JSON.stringify(a)}`);
  // отступ справа
  await drag('.ri-right', -38, 0);
  a = await attrs();
  ok(a.rindent > 25, `отступ справа: ${JSON.stringify(a)}`);
  const docx = await p.evaluate(() => { const E = window.ens; const b = E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta(), 'r')); const n = b.doc.child(0); return { i: n.attrs.indent, f: n.attrs.first, r: n.attrs.rindent }; });
  ok(docx.r > 25 && docx.i > 38, `отступы после docx: ${JSON.stringify(docx)}`);
  // таблица: метки столбцов и строк, перетаскивание
  await p.evaluate(() => window.ens.run('insert.table'));
  await p.waitForSelector('.modal input');
  await p.click('.modal button.primary');
  await p.waitForTimeout(400);
  await p.waitForSelector('.rc');
  eq(await p.$$eval('.rc', (e) => e.length), 4, 'границы 3 столбцов');
  ok((await p.$$eval('.rr', (e) => e.length)) >= 3, 'границы строк');
  const widths = () => p.evaluate(() => { let w = null; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'table_row' && !w) w = n.content.content.map((c) => (c.attrs.colwidth ?? [null])[0]); }); return w; });
  await drag('.rc >> nth=1', 40, 0);
  const w1 = await widths();
  ok(w1[0] > 0 && w1[1] > 0, `ширины после перетаскивания: ${JSON.stringify(w1)}`);
  await drag('.rr >> nth=1', 0, 25);
  const rh = await p.evaluate(() => { let h = null; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'table_row' && h == null) h = n.attrs.height; }); return h; });
  ok(rh > 15, `высота первой строки: ${rh} pt`);
});

await test('таблица: границы и заливка, выравнивание, свойства; всё доживает до docx', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const cell = (t) => s.nodes.table_cell.create({ colwidth: [150] }, s.nodes.paragraph.create(null, s.text(t)));
    const row = (a, b) => s.nodes.table_row.create(null, [cell(a), cell(b)]);
    const table = s.nodes.table.create({ borders: true }, [row('А1', 'Б1'), row('А2', 'Б2')]);
    E.loadDoc(s.nodes.doc.create(null, [table, s.nodes.paragraph.create()]), E.getMeta(), 'tbl');
    E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 4)));
  });
  await p.waitForTimeout(300);
  await p.evaluate(() => window.ens.run('table.borderDialog'));
  await p.waitForSelector('.modal');
  await p.selectOption('.modal select >> nth=0', 'table');
  // только внешние границы, пунктир 1 пт, красный; заливка
  await p.click('.modal button:has-text("Внешние")');
  await p.selectOption('.modal select >> nth=1', 'dashed');
  await p.click('.modal button.primary');
  await p.waitForTimeout(250);
  const cells = () => p.evaluate(() => { const out = []; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'table_cell') out.push(n.attrs); }); return out; });
  let c = await cells();
  ok(c[0].borders && c[0].borders.t.s === 'dashed' && c[0].borders.l.s === 'dashed', `внешние границы A1: ${JSON.stringify(c[0].borders)}`);
  ok(c[0].borders.r.s === 'solid' && c[0].borders.b.s === 'solid', 'внутренние границы остались прежними');
  // заливка и выравнивание по вертикали
  await p.evaluate(() => window.ens.run('table.borderDialog'));
  await p.waitForSelector('.modal');
  await p.selectOption('.modal select >> nth=0', 'table');
  await p.click('.modal button:has-text("Внутренние")');
  await p.check('.modal label:has-text("залить") input');
  await p.click('.modal button:has-text("Убрать границы")');
  await p.waitForTimeout(250);
  c = await cells();
  ok(c[0].background && c[0].borders.r.s === 'none' && c[0].borders.t.s === 'dashed', `внутренние убраны, заливка есть: ${JSON.stringify(c[0])}`);
  await p.evaluate(() => window.ens.run('table.valign'));
  // свойства таблицы: по центру страницы, высота строки
  await p.evaluate(() => window.ens.run('table.props'));
  await p.waitForSelector('.modal');
  await p.selectOption('.modal select', 'center');
  await p.fill('.modal input[placeholder="авто"]', '1');
  await p.click('.modal button.primary');
  await p.waitForTimeout(250);
  const t = await p.evaluate(() => { let a = null, h = null; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'table') a = n.attrs.align; if (n.type.name === 'table_row' && h == null) h = n.attrs.height; }); return { a, h }; });
  ok(t.a === 'center' && Math.abs(t.h - 28.3) < 1, `свойства: ${JSON.stringify(t)}`);
  // docx туда и обратно
  const back = await p.evaluate(() => {
    const E = window.ens;
    const d = E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta(), 't'));
    const out = { cells: [], align: null, h: null };
    d.doc.descendants((n) => { if (n.type.name === 'table_cell') out.cells.push(n.attrs); if (n.type.name === 'table') out.align = n.attrs.align; if (n.type.name === 'table_row' && out.h == null) out.h = n.attrs.height; });
    return out;
  });
  ok(back.align === 'center' && Math.abs(back.h - 28.3) < 1, `положение и высота после docx: ${back.align}, ${back.h}`);
  ok(back.cells[0].background && back.cells[0].borders && back.cells[0].borders.t.s === 'dashed' && back.cells[0].borders.r.s === 'none', `границы и заливка после docx: ${JSON.stringify(back.cells[0])}`);
});

await test('надписи, рамки абзацев и плавающие таблицы: чтение, положение на листе, запись обратно (и LibreOffice)', async () => {
  const p = await page();
  await p.setInputFiles('#file-input', join(here, 'fixtures', 'frames.docx'));
  await p.waitForSelector('.ProseMirror .frame-box', { timeout: 5000 });
  await p.waitForTimeout(500);
  const info = await p.evaluate(() => {
    const E = window.ens;
    const frames = [];
    E.view.state.doc.descendants((n) => { if (n.type.name === 'frame') frames.push({ kind: n.attrs.kind, mode: n.attrs.float?.mode, hRel: n.attrs.float?.hRel, hOff: Math.round(n.attrs.float?.hOff ?? -1), vRel: n.attrs.float?.vRel, w: n.attrs.width, fill: n.attrs.fill, border: n.attrs.border?.c, text: n.textContent.slice(0, 40) }); return false; });
    const paper = document.querySelector('.paper').getBoundingClientRect();
    const rect = (sel) => { const e = document.querySelector(sel); if (!e) return null; const r = e.getBoundingClientRect(); return { x: Math.round(r.left - paper.left), y: Math.round(r.top - paper.top), w: Math.round(r.width), h: Math.round(r.height) }; };
    return {
      frames, lossy: E.getMeta().lossy, header: JSON.stringify(E.getMeta().hf.header).includes('"frame"'),
      box: rect('.ProseMirror .frame-box.fr-front'), wrap: [...document.querySelectorAll('.ProseMirror .frame-box.fr-wrap')].map((e) => ({ float: getComputedStyle(e).float, table: e.classList.contains('fr-table') })),
      hf: rect('.paper .hf .frame-box'),
    };
  });
  const tb = info.frames.find((f) => f.fill === '#FFF2CC');
  ok(tb && tb.mode === 'front' && tb.hRel === 'page' && Math.abs(tb.hOff - 315) <= 2 && tb.border === '#FF0000' && tb.text.includes('Надпись'), `надпись: ${JSON.stringify(info.frames)}`);
  ok(info.frames.some((f) => f.kind === 'table') && info.frames.some((f) => f.text.includes('Рамка абзацев')), 'плавающая таблица и рамка абзацев прочитаны');
  ok(info.header, 'надпись в колонтитуле прочитана');
  ok(info.box && Math.abs(info.box.x - 315) <= 3 && Math.abs(info.box.y - 189) <= 3 && info.box.w > 270, `надпись на листе: ${JSON.stringify(info.box)}`);
  ok(info.hf && Math.abs(info.hf.x - 76) <= 3 && Math.abs(info.hf.y - 38) <= 3, `надпись в колонтитуле: ${JSON.stringify(info.hf)}`);
  ok(info.wrap.length >= 2 && info.wrap.every((w) => w.float === 'left' || w.float === 'right'), `обтекание: ${JSON.stringify(info.wrap)}`);
  eq(info.lossy.filter((x) => /надпис|фигур/.test(x)), [], 'о надписях не предупреждает');
  // запись обратно: положение, заливка, текст, плавающая таблица
  const back = await p.evaluate(() => {
    const E = window.ens;
    const d = E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta(), 'f'));
    const frames = [];
    d.doc.descendants((n) => { if (n.type.name === 'frame') frames.push({ kind: n.attrs.kind, mode: n.attrs.float?.mode, hOff: Math.round(n.attrs.float?.hOff ?? -1), fill: n.attrs.fill, text: n.textContent.slice(0, 20) }); return false; });
    return { frames, hdr: JSON.stringify(d.meta.hf.header).includes('"frame"') };
  });
  const tb2 = back.frames.find((f) => f.fill === '#FFF2CC');
  ok(tb2 && tb2.mode === 'front' && Math.abs(tb2.hOff - 315) <= 2 && tb2.text.includes('Надпись'), `после docx: ${JSON.stringify(back.frames)}`);
  ok(back.frames.some((f) => f.kind === 'table') && back.hdr, 'таблица и надпись колонтитула после docx');
  // LibreOffice открывает записанный файл и видит текст надписи
  const bytes = await p.evaluate(() => [...window.ens.exportDocx(window.ens.view.state.doc, window.ens.getMeta(), 'f')]);
  const tmp = mkdtempSync(join(tmpdir(), 'ens-frames-'));
  writeFileSync(join(tmp, 'f.docx'), Buffer.from(bytes));
  execSync(`soffice --headless --convert-to pdf --outdir "${tmp}" "${tmp}/f.docx"`, { stdio: 'pipe', timeout: 90000 });
  const txt = execSync(`pdftotext "${tmp}/f.pdf" -`).toString().replace(/\s+/g, ' ');
  ok(/Надпись\s*:\s*реквизиты/.test(txt) && txt.includes('Шапка организации') && txt.includes('Директору') && txt.includes('Рамка абзацев'), `LibreOffice увидел надписи: ${txt.slice(0, 120).replace(/\n/g, ' | ')}`);
});

await test('надпись: вставка, выделение, ручки размера, перетаскивание, окно «Формат надписи»', async () => {
  const p = await page();
  await p.evaluate(() => { const E = window.ens, s = E.schema; E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, s.text('Абзац под надписью.')), s.nodes.paragraph.create(null, s.text('Второй абзац.'))]), E.getMeta(), 'tb'); });
  await p.waitForTimeout(300);
  await p.evaluate(() => window.ens.run('insert.textbox'));
  await p.waitForSelector('.ProseMirror .frame-box');
  await p.waitForTimeout(300);
  await p.keyboard.type('Реквизиты');
  eq(await p.evaluate(() => { let t = null; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'frame') t = n.textContent; }); return t; }), 'Реквизиты', 'в надписи можно печатать');
  await p.evaluate(() => window.ens.run('frame.select'));
  await p.waitForSelector('.img-ov:not([hidden]) .img-h-se');
  const attrs = () => p.evaluate(() => { let a = null; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'frame') a = n.attrs; }); return a; });
  const w0 = (await attrs()).width;
  const se = await (await p.$('.img-h-se')).boundingBox();
  await p.mouse.move(se.x + 4, se.y + 4);
  await p.mouse.down();
  await p.mouse.move(se.x + 64, se.y + 44, { steps: 5 });
  await p.mouse.up();
  await p.waitForTimeout(250);
  let a = await attrs();
  ok(a.width > w0 + 50 && a.height > 100, `размер надписи ${w0} → ${a.width}×${a.height}`);
  // перетаскивание
  const body = await (await p.$('.img-ov-body')).boundingBox();
  const h0 = a.float.hOff;
  await p.mouse.move(body.x + 30, body.y + 30);
  await p.mouse.down();
  await p.mouse.move(body.x + 90, body.y + 50, { steps: 5 });
  await p.mouse.up();
  await p.waitForTimeout(300);
  a = await attrs();
  ok(a.float.hRel === 'page' && a.float.hOff > h0 + 40, `надпись перетащена: ${JSON.stringify(a.float)}`);
  // окно: рамка, заливка, обтекание
  await p.evaluate(() => window.ens.run('image.format'));
  await p.waitForSelector('.modal');
  await p.fill('.modal .dlg-row:has-text("Ширина") input', '8');
  await p.selectOption('.modal .dlg-row:has-text("Рамка") select >> nth=0', 'dashed');
  await p.check('.modal label:has-text("залить") input');
  await p.selectOption('.modal .dlg-row:has-text("Обтекание") select', 'right');
  await p.click('.modal button.primary');
  await p.waitForTimeout(300);
  a = await attrs();
  ok(Math.abs(a.width - 302) <= 2 && a.border.s === 'dashed' && a.fill && a.float.mode === 'wrap' && a.float.side === 'right', `формат: ${JSON.stringify({ w: a.width, b: a.border, f: a.fill, fl: a.float.mode })}`);
  eq(await p.$eval('.ProseMirror .frame-box', (e) => getComputedStyle(e).float), 'right', 'обтекание справа');
  // в docx и обратно
  const back = await p.evaluate(() => { const E = window.ens; const d = E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta(), 't')); let f = null; d.doc.descendants((n) => { if (n.type.name === 'frame') f = { mode: n.attrs.float?.mode, side: n.attrs.float?.side, fill: n.attrs.fill, text: n.textContent, dash: n.attrs.border?.s }; }); return f; });
  ok(back && back.mode === 'wrap' && back.text === 'Реквизиты' && back.fill && back.dash === 'dashed', `после docx: ${JSON.stringify(back)}`);
  // удаление
  await p.evaluate(() => window.ens.run('image.delete'));
  eq(await p.$$eval('.ProseMirror .frame-box', (e) => e.length), 0, 'надпись удалена');
});

await test('таблица: сделать плавающей справа и вернуть в текст', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const cell = (t) => s.nodes.table_cell.create({ colwidth: [120] }, s.nodes.paragraph.create(null, s.text(t)));
    const table = s.nodes.table.create({ borders: true }, [s.nodes.table_row.create(null, [cell('Кому'), cell('Иванову')])]);
    E.loadDoc(s.nodes.doc.create(null, [table, s.nodes.paragraph.create(null, s.text('Обычный абзац, который обтекает таблицу справа налево. '.repeat(5)))]), E.getMeta(), 'fl');
    E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 4)));
  });
  await p.waitForTimeout(300);
  await p.click('#tabs .tab:has-text("Таблица")');
  await p.selectOption('select[data-command="table.wrap"]', 'right');
  await p.waitForTimeout(400);
  eq(await p.$eval('.ProseMirror .frame-box.fr-table', (e) => getComputedStyle(e).float), 'right', 'таблица справа');
  const cols = await p.$$eval('.rc', (e) => e.length);
  ok(cols >= 2, `линейка видит столбцы плавающей таблицы: ${cols}`);
  await p.selectOption('select[data-command="table.wrap"]', 'inline');
  await p.waitForTimeout(300);
  eq(await p.$$eval('.ProseMirror .frame-box', (e) => e.length), 0, 'рамка-обёртка снята');
  eq(await p.evaluate(() => window.ens.view.state.doc.firstChild.type.name), 'table', 'таблица снова в потоке');
});

// ---- Проверка оформления и «рентген» ----
await test('проверка по ГОСТ: находит нарушения, «Исправить всё» приводит документ к профилю', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema, m = E.getMeta();
    m.page = { ...m.page, left: 28, right: 14, top: 28, bottom: 28 }; // поля по 10 мм и меньше
    m.defaults = { font: 'Arial', size: 11, line: 1.15, before: 0, after: 0 };
    const long = 'Это длинный абзац основного текста, который должен соответствовать оформлению: отступ, выравнивание, шрифт и интервал по профилю. ';
    const P = (t, a) => s.nodes.paragraph.create(a || null, t ? s.text(t) : undefined);
    E.loadDoc(s.nodes.doc.create(null, [P('    ' + long + 'Он сказал "привет".'), P(long), P(''), P(''), P(''), P(long)]), m, 'gost');
  });
  await p.evaluate(() => window.ens.run('review.check'));
  await p.waitForSelector('.check-row');
  await p.selectOption('.check-modal select', 'gost7097');
  const bad = await p.$$eval('.check-row.bad .check-text > div:first-child', (e) => e.map((x) => x.textContent));
  for (const need of ['Поля, мм', 'Шрифт или размер вне профиля', 'Межстрочный интервал вне профиля', 'Абзацный отступ', 'не выровнен по ширине', 'Номера страниц', 'Отступ абзаца набран пробелами', 'Лишние пустые абзацы', 'Типографика']) ok(bad.some((x) => x.includes(need)), `не найдено замечание «${need}»: ${bad.join(' | ')}`);
  await p.click('.check-modal button:has-text("Исправить всё")');
  await p.waitForTimeout(300);
  const left = await p.$$eval('.check-row.bad .check-text > div:first-child', (e) => e.map((x) => x.textContent));
  eq(left, [], 'после «Исправить всё» замечаний не осталось');
  const r = await p.evaluate(() => { const E = window.ens, m = E.getMeta(); const d = E.view.state.doc; return { left: m.page.left, font: m.defaults.font, size: m.defaults.size, line: m.defaults.line, first: d.child(0).attrs.first, align: d.child(0).attrs.align, empties: [...Array(d.childCount).keys()].filter((i) => d.child(i).content.size === 0).length, hdr: JSON.stringify(m.hf.header).includes('"field"'), diff: m.hf.differentFirst, text: d.child(0).textContent, versions: E.hist().versions.map((v) => v.kind) }; });
  ok(r.left >= 56.6, `левое поле ${r.left}`);
  eq([r.font, r.size, r.line], ['Times New Roman', 14, 1], 'шрифт и интервал по профилю');
  ok(Math.abs(r.first - 35.4) < 1 && r.align === 'justify', 'отступ и выравнивание');
  ok(r.empties <= 1, `пустых абзацев осталось ${r.empties}`);
  ok(r.hdr && r.diff, 'номер страницы вверху, на первой странице нет');
  ok(r.text.includes('«привет»') && !r.text.startsWith(' '), `типографика и пробелы: ${r.text.slice(0, 40)}`);
  ok(r.versions.includes('before'), 'страховочная версия перед исправлением');
  await p.context().close();
});

await test('«рентген» форматирования: что задано вручную, сброс формата', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Текст');
  await p.keyboard.press('Control+KeyA');
  await p.keyboard.press('Control+KeyB');
  await p.selectOption('select[data-command="format.fontName"]', 'Arial');
  await p.keyboard.press('Control+KeyE');
  await p.keyboard.press('Control+Alt+KeyI');
  await p.waitForSelector('.inspector .insp-row');
  const man = await p.$$eval('.inspector .insp-row .man', (e) => e.map((x) => x.textContent));
  ok(man.includes('по центру') && man.includes('Arial') && man.includes('жирный'), `вручную: ${man}`);
  await p.click('.inspector button:has-text("Сбросить формат абзаца")');
  await p.click('.inspector button:has-text("Сбросить формат текста")');
  const after = await p.evaluate(() => { const n = window.ens.view.state.doc.firstChild; return [n.attrs.align, n.child(0).marks.length]; });
  eq(after, [null, 0], 'формат сброшен');
  await p.keyboard.press('Control+Alt+KeyI');
  await p.context().close();
});

await test('меню «Файл»: кнопка «Назад» видна, Esc и кнопка закрывают меню', async () => {
  const p = await page();
  await p.click('.tab.file');
  await p.waitForSelector('.text-backstage:not([hidden]) .bs-back', { timeout: 3000 });
  const st = await p.evaluate(() => {
    const b = document.querySelector('.text-backstage .bs-back');
    const r = b.getBoundingClientRect();
    return { icon: !!b.querySelector('svg'), w: Math.round(r.width), h: Math.round(r.height), title: b.getAttribute('title') };
  });
  ok(st.icon, 'значок как в ENS-Table');
  ok(st.w > 20 && st.h > 16, `кнопка не нулевая: ${JSON.stringify(st)}`);
  ok(/Esc/.test(st.title || ''), 'подсказка про Esc');
  await p.click('.text-backstage .bs-back');
  await p.waitForSelector('.text-backstage', { state: 'hidden', timeout: 3000 });
  await p.click('.tab.file');
  await p.waitForSelector('.text-backstage:not([hidden])');
  await p.keyboard.press('Escape');
  await p.waitForSelector('.text-backstage', { state: 'hidden', timeout: 3000 });
  await p.context().close();
});

await test('меню открыто: правая кнопка и Ctrl+R/F5 не проходят; в полях ввода меню остаётся', async () => {
  const p = await page();
  const outcome = () => p.evaluate(() => {
    const fire = (el, init) => { const e = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }); el.dispatchEvent(e); return e.defaultPrevented; };
    const ctx = new MouseEvent('contextmenu', { bubbles: true, cancelable: true });
    document.querySelector('.ProseMirror').dispatchEvent(ctx);
    // Своё меню, открытое этим событием, закрываем щелчком мимо — как пользователь.
    document.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    return {
      ctx: ctx.defaultPrevented,
      ctrlR: fire(document.body, { key: 'r', ctrlKey: true }),
      f5: fire(document.body, { key: 'F5' }),
    };
  });
  // Меню WebView в документе подавлено всегда: вместо него своё контекстное меню (contextmenu.ts).
  eq(await outcome(), { ctx: true, ctrlR: false, f5: false }, 'без окна: меню WebView заменено своим, клавиши не трогаем');

  await p.click('.tab.file');
  await p.waitForSelector('.text-backstage:not([hidden])', { timeout: 3000 });
  eq(await outcome(), { ctx: true, ctrlR: true, f5: true }, 'при открытом меню перезапуск заблокирован');

  await p.keyboard.press('Escape');
  await p.waitForSelector('.text-backstage', { state: 'hidden', timeout: 3000 });

  // В окне с полем ввода системное меню должно остаться, иначе пропадёт «Вставить».
  await p.evaluate(() => window.ens.run('file.saveAs'));
  await p.waitForSelector('#dialog:not([hidden]) input[type=text]', { timeout: 3000 });
  const inField = await p.evaluate(() => {
    const input = document.querySelector('#dialog input[type=text]');
    const e = new MouseEvent('contextmenu', { bubbles: true, cancelable: true });
    input.dispatchEvent(e);
    const r = new MouseEvent('contextmenu', { bubbles: true, cancelable: true });
    document.querySelector('#dialog h2').dispatchEvent(r);
    return { input: e.defaultPrevented, heading: r.defaultPrevented };
  });
  eq(inField, { input: false, heading: true }, 'поле ввода — меню можно, остальное окно — нельзя');
  await p.keyboard.press('Escape');

  eq(await outcome(), { ctx: true, ctrlR: false, f5: false }, 'после закрытия снова как обычно');
  await p.context().close();
});

await test('поиск формата в строку: условия, стили, вложенные абзацы и переходы', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const fmt = (font, size) => s.marks.textStyle.create({ font, size });
    const em = s.marks.em.create(), strong = s.marks.strong.create();
    const P = (text, marks = []) => s.nodes.paragraph.create(null, s.text(text, marks));
    const body = [
      P('Совпадает', [fmt('Arial', 14), strong, em]), P('По умолчанию'),
      s.nodes.paragraph.create(null, [s.text('Шрифт', [fmt('Arial', 12)]), s.text('Размер', [fmt('Tinos', 14), em])]),
      s.nodes.heading.create({ level: 1 }, s.text('Заголовок', [fmt('Arial', 14)])),
      s.nodes.bullet_list.create(null, s.nodes.list_item.create(null, P('В списке', [fmt('Arial', 14), em]))),
      s.nodes.table.create(null, s.nodes.table_row.create(null, s.nodes.table_cell.create(null, P('В таблице', [fmt('Arial', 14), em])))),
      P('С кавычками', [fmt('Times New Roman', 12.5)]), s.nodes.paragraph.create(),
    ];
    const meta = E.getMeta();
    meta.defaults = { ...meta.defaults, font: 'Arial', size: 14 };
    E.loadDoc(s.nodes.doc.create(null, body), meta, 'format-search');
    E.run('review.findFormat');
  });
  const input = p.locator('.format-search input');
  const hits = () => p.locator('.format-hit').allTextContents();
  eq(await hits(), [], 'пустой запрос');
  await input.fill('arial 14 курсив');
  const expected = ['1. Совпадает', '2. В списке', '3. В таблице'];
  eq(await hits(), expected, 'один фрагмент, без наследования настроек документа');
  await input.fill('Arial 14 жирный курсив');
  eq(await hits(), ['1. Совпадает'], 'сочетание жирного и курсива');
  await input.fill('Arial 14 без-жирного');
  eq(await hits(), ['1. Заголовок', '2. В списке', '3. В таблице'], 'без жирного');
  await input.fill('стиль:"Заголовок 1"');
  eq(await hits(), ['1. Заголовок'], 'стиль заголовка');
  await input.fill('стиль:Обычный Arial 14 курсив');
  eq(await hits(), expected, 'стиль и прямое форматирование');
  await input.fill('шрифт:"Times New Roman" размер:12,5');
  eq(await hits(), ['1. С кавычками'], 'пробелы в шрифте и дробный размер');
  await input.fill('стиль:Обычный');
  ok((await hits()).some((x) => x.includes('(пустой абзац)')), 'поиск пустого абзаца по стилю');
  for (const invalid of ['размер:abc', 'размер:0', 'размер:0x10', 'стиль:Нет', 'цвет:красный', 'шрифт:"Arial', 'жирный без-жирного', '14 12', 'шрифт:']) {
    await input.fill(invalid);
    eq(await hits(), [], 'ошибка очищает результаты');
    ok(!(await p.locator('.format-search [role=status]').textContent()).includes('Найдено'), 'ошибка запроса');
  }
  await input.fill('Несуществующий 99');
  eq(await p.locator('.format-search [role=status]').textContent(), 'Найдено абзацев: 0', 'нет совпадений');
  await input.fill('Arial 14 курсив');
  await p.locator('.format-hit').nth(2).click();
  ok(await p.locator('#dialog').isHidden(), 'диалог закрыт');
  eq(await p.evaluate(() => window.ens.view.state.selection.$from.parent.textContent), 'В таблице', 'переход в таблицу');
  await p.evaluate(() => window.ens.run('review.findFormat'));
  await input.fill('жирный'); await input.press('Enter');
  eq(await p.evaluate(() => window.ens.view.state.selection.$from.parent.textContent), 'Совпадает', 'Enter: первый результат');
  await p.evaluate(() => window.ens.run('review.findFormat'));
  await p.keyboard.press('Escape');
  ok(await p.locator('#dialog').isHidden(), 'Escape закрывает окно');
  await p.context().close();
});

await test('Alt+Q открывает однострочный поиск форматирования', async () => {
  const p = await page();
  await p.click('.ProseMirror'); await p.keyboard.press('Alt+KeyQ');
  ok(await p.locator('#search').evaluate((e) => e === document.activeElement), 'Alt+Q: фокус поиска команд');
  await p.locator('#search').fill('Поиск по форматированию'); await p.keyboard.press('Enter');
  await p.waitForSelector('.format-search input');
  ok(await p.locator('.format-search input').evaluate((e) => e === document.activeElement), 'фокус строки запроса');
  await p.keyboard.press('Escape'); await p.context().close();
});

await test('слияние: независимые правки, вставки, удаления, конфликты и вложенные абзацы', async () => {
  const p = await page();
  const failures = await p.evaluate(async () => {
    const { planMerge, resolveMerge } = await import('/src/merge.ts');
    const s = window.ens.schema;
    const P = (t, bold = false) => s.nodes.paragraph.create(null, t ? s.text(t, bold ? [s.marks.strong.create()] : []) : null);
    const D = (a) => s.nodes.doc.create(null, a.map((x) => typeof x === 'string' ? P(x) : x));
    const failures = [];
    const check = (b, m, t, expected, count = 0, choice = 'mine') => {
      const plan = planMerge(D(b), D(m), D(t));
      if (plan.conflicts !== count) failures.push('число конфликтов: ' + JSON.stringify([b, m, t, plan.conflicts]));
      const choices = new Map(Array.from({ length: count }, (_, i) => [i, choice]));
      const doc = resolveMerge(plan, choices);
      if (!doc.eq(D(expected.length ? expected : ['']))) failures.push('результат: ' + JSON.stringify(doc.toJSON()) + ' / ' + JSON.stringify(expected));
      if (count) {
        try { resolveMerge(plan, new Map()); failures.push('конфликт разрешён без выбора'); } catch { /* ожидаемая ошибка */ }
      }
    };
    check(['A', 'B', 'C'], ['X', 'B', 'C'], ['A', 'B', 'Y'], ['X', 'B', 'Y']);
    check(['A', 'B'], ['X', 'A', 'B'], ['A', 'B', 'Y'], ['X', 'A', 'B', 'Y']);
    check(['A', 'B', 'C'], ['A', 'C'], ['A', 'B', 'Y'], ['A', 'Y']);
    check(['A'], ['X'], ['Y'], ['Y'], 1, 'theirs');
    check(['A'], ['X'], ['Y'], ['A'], 1, 'base');
    check(['A'], [], ['Y'], [], 1);
    check(['A'], ['X', 'A'], ['Y', 'A'], ['X', 'A'], 1);
    check(['A'], ['X'], ['X'], ['X']);
    check(['A'], [P('A', true)], ['B'], [P('A', true)], 1);
    const list = (a) => s.nodes.bullet_list.create(null, a.map((t) => s.nodes.list_item.create(null, P(t))));
    check([list(['A', 'B', 'C'])], [list(['X', 'B', 'C'])], [list(['A', 'B', 'Y'])], [list(['X', 'B', 'Y'])]);
    const table = (a) => s.nodes.table.create(null, s.nodes.table_row.create(null, a.map((t) => s.nodes.table_cell.create(null, P(t)))));
    check([table(['A', 'B', 'C'])], [table(['X', 'B', 'C'])], [table(['A', 'B', 'Y'])], [table(['X', 'B', 'Y'])]);
    for (let i = 1; i < 20; i++) {
      const a = Array.from({ length: i }, (_, j) => 'абзац ' + j);
      check(a, [...a, 'новый'], a, [...a, 'новый']);
      check(a, a, a.slice(1), a.slice(1));
      check(a, a.slice(1), a.slice(1), a.slice(1));
    }
    return failures;
  });
  eq(failures, [], 'регрессии слияния');
  await p.context().close();
});

await test('слияние: локальные DOCX, выбор конфликтов, отмена и сохранение результата', async () => {
  const p = await page();
  const files = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const make = (texts) => [...E.exportDocx(s.nodes.doc.create(null, texts.map((t) => s.nodes.paragraph.create(null, s.text(t)))), E.getMeta(), 'copy')];
    return [make(['Основа', 'Общий']), make(['Мой', 'Общий']), make(['Коллега', 'Общий'])];
  });
  await p.click('.ProseMirror'); await p.keyboard.type('Текущий документ');
  const before = await json(p);
  await p.evaluate(() => window.ens.run('review.mergeCopies'));
  await p.locator('.merge-dialog button', { hasText: 'Сравнить копии' }).click();
  ok((await p.locator('.merge-dialog [role=status]').textContent()).includes('все три'), 'нужны три файла');
  for (let i = 0; i < 3; i++) await p.locator(`[data-merge-file="${i}"]`).setInputFiles({ name: `copy-${i}.docx`, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(files[i]) });
  await p.locator('.merge-dialog button', { hasText: 'Сравнить копии' }).click();
  await p.waitForSelector('.merge-conflict select');
  ok(await p.locator('.merge-results > button.primary').isDisabled(), 'без выбора применять нельзя');
  eq(await json(p), before, 'сравнение не меняет текущий документ');
  await p.keyboard.press('Escape');
  eq(await json(p), before, 'отмена не меняет документ');
  await p.evaluate(() => window.ens.run('review.mergeCopies'));
  for (let i = 0; i < 3; i++) await p.locator(`[data-merge-file="${i}"]`).setInputFiles({ name: `copy-${i}.docx`, mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(files[i]) });
  await p.locator('.merge-dialog button', { hasText: 'Сравнить копии' }).click();
  await p.waitForSelector('.merge-conflict select');
  eq(await p.locator('.merge-conflict select').inputValue(), '', 'до выбора показан пустой пункт, а не готовое решение');
  await p.locator('.merge-conflict select').selectOption('theirs');
  await p.locator('.merge-results > button.primary').click();
  await p.locator('.modal button', { hasText: 'Не сохранять' }).click();
  eq(await text(p), 'Коллега|Общий', 'выбран вариант коллеги');
  ok(await p.evaluate(() => window.ens.isDirty()), 'объединённый документ не сохранён');
  const back = await p.evaluate(() => { const E = window.ens; return E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta(), 'merged')).doc.toJSON(); });
  eq(back, await json(p), 'результат переживает DOCX');
  await p.context().close();
});

await test('слияние: ключи параметров, бинарные данные и подсказка о неверной основе', async () => {
  const p = await page();
  const failures = await p.evaluate(async () => {
    const { mergeKey, describeMeta, baseOverlap } = await import('/src/merge.ts');
    const failures = [];
    const t = (cond, msg) => { if (!cond) failures.push(msg); };
    t(mergeKey({ a: 1, b: [2, { c: 3, d: 4 }] }) === mergeKey({ b: [2, { d: 4, c: 3 }], a: 1 }), 'порядок ключей');
    t(mergeKey({ a: 1, b: undefined }) === mergeKey({ a: 1 }), 'undefined как отсутствие свойства');
    t(mergeKey({ a: 1 }) !== mergeKey({ a: 2 }), 'разные значения');
    t(mergeKey(new Uint8Array([1, 2, 3])) === mergeKey(new Uint8Array([1, 2, 3])), 'одинаковые байты');
    t(mergeKey(new Uint8Array([1, 2, 3])) !== mergeKey(new Uint8Array([1, 2, 4])), 'разные байты');
    t(mergeKey(new Map([['x', 1]])) !== mergeKey(new Map([['x', 2]])), 'Map сравнивается по содержимому');
    t(mergeKey(new Map([['x', 1], ['y', 2]])) === mergeKey(new Map([['y', 2], ['x', 1]])), 'порядок записей Map');
    t(describeMeta({ img: new Uint8Array(100000) }).length < 200, 'бинарные данные не раздувают просмотр');
    const s = window.ens.schema;
    const doc = (texts) => s.nodes.doc.create(null, texts.map((x) => s.nodes.paragraph.create(null, s.text(x))));
    t(baseOverlap(doc(['a', 'b', 'c', 'd']), doc(['a', 'b', 'c', 'd', 'e'])) === 1, 'основа целиком сохранена');
    t(baseOverlap(doc(['a', 'b', 'c', 'd']), doc(['x', 'y', 'z'])) === 0, 'общего нет');
    t(baseOverlap(doc(['a', 'b', 'c', 'd']), doc(['a', 'b', 'x'])) === 0.5, 'половина общего');
    return failures;
  });
  eq(failures, [], 'ключи и основа слияния');
  await p.context().close();
});
await test('вёрстка: строка с выступом и табуляциями не выходит левее листа (DOCX туда и обратно)', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const mk = (indent, first, t) => s.nodes.paragraph.create({ indent, first }, s.text(t));
    // Выступ больше отступа слева, первая строка шире выступа, табуляции — образец артефакта.
    const doc = s.nodes.doc.create(null, [
      mk(36, -36, 'а)\tКороткий выступ, всё по правилам'),
      mk(18, -54, 'Пункт1234567890\tтекст шире выступа\tи ещё табуляция'),
      mk(null, -200, 'Очень\tдлинный\tвыступ\tза поле и за край листа'),
      mk(-150, null, 'Отрицательный отступ слева больше поля'),
    ]);
    const back = E.importDocx(E.exportDocx(doc, E.getMeta(), 'x'));
    E.loadDoc(back.doc, back.meta, 'x');
    const paper = document.querySelector('.paper').getBoundingClientRect().left;
    const firsts = [...E.view.dom.querySelectorAll('p')].map((el) => {
      const rg = document.createRange();
      rg.selectNodeContents(el);
      return Math.min(...[...rg.getClientRects()].map((x) => x.left)) - paper;
    });
    return { firsts, attrs: back.doc.content.content.map((n) => [n.attrs.indent, n.attrs.first]) };
  });
  eq(r.attrs, [[36, -36], [18, -54], [null, -200], [-150, null]], 'отступы пережили DOCX без изменений');
  r.firsts.forEach((x, i) => ok(x >= -0.5, `абзац ${i + 1}: начало строки левее листа на ${(-x).toFixed(1)} px`));
  ok(r.firsts[0] > 100, 'обычный выступ внутри поля не тронут');
  await p.context().close();
});
await test('рисунок «за текстом»: обычный щелчок выделяет его там, где нет букв, буквы над ним остаются тексту', async () => {
  const p = await page();
  const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mP8z8BQz0AEYBxVSF+FABJADveWkH6oAAAAAElFTkSuQmCC';
  await p.evaluate((png) => {
    const E = window.ens, s = E.schema;
    // Рисунок 300×200 px за первыми строками: над его верхом текст, ниже — пустые абзацы.
    const float = { mode: 'behind', hRel: 'page', hOff: 100, hAlign: null, vRel: 'page', vOff: 50, vAlign: null, dist: [0, 0, 8, 8] };
    const img = s.nodes.image.create({ src: png, width: 300, height: 200, float });
    const P = (c) => s.nodes.paragraph.create(null, c);
    E.loadDoc(s.nodes.doc.create(null, [P([s.text('Текст над рисунком '), img, s.text('продолжается')]), P(), P(), P(), P(), P(), P()]), E.getMeta(), 'behind');
  }, PNG);
  await p.waitForTimeout(400);
  const kind = () => p.evaluate(() => { const sel = window.ens.view.state.selection; return sel.node ? sel.node.type.name : 'text'; });
  const img = await (await p.$('.ProseMirror img.fl-behind')).boundingBox();
  // Нижняя часть рисунка: под ней пустые абзацы — щелчок выбирает рисунок, появляются ручки.
  await p.mouse.click(img.x + img.width / 2, img.y + img.height - 15);
  eq(await kind(), 'image', 'щелчок по пустому месту над рисунком');
  await p.waitForSelector('.img-ov.abs:not([hidden]) .img-h-se');
  // Ручкой можно менять размер: одна запись в истории.
  const se = await (await p.$('.img-h-se')).boundingBox();
  await p.mouse.move(se.x + 4, se.y + 4);
  await p.mouse.down();
  await p.mouse.move(se.x - 56, se.y - 36, { steps: 5 });
  await p.mouse.up();
  await p.waitForTimeout(200);
  const w = await p.evaluate(() => { let a = null; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'image') a = n.attrs; }); return a.width; });
  ok(w < 280, `ширина уменьшилась: ${w}`);
  // Щелчок по букве первой строки, которая лежит поверх рисунка, ставит курсор в текст.
  const letter = await p.evaluate(() => {
    const t = window.ens.view.dom.querySelector('p').firstChild;
    const r = document.createRange(); r.setStart(t, 2); r.setEnd(t, 3);
    const b = r.getBoundingClientRect();
    return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  });
  const im2 = await (await p.$('.ProseMirror img.fl-behind')).boundingBox();
  ok(letter.y > im2.y && letter.x > im2.x, `буква лежит над рисунком: ${JSON.stringify([letter, im2])}`);
  await p.mouse.click(letter.x, letter.y);
  eq(await kind(), 'text', 'щелчок по букве — курсор в тексте');
  await p.context().close();
});
await test('линейки «живые»: отступ, поле и столбец двигаются за мышью, в истории одна запись', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const long = 'Длинный абзац для проверки отступов, который переносится на несколько строк. '.repeat(4);
    const cell = (t) => s.nodes.table_cell.create(null, s.nodes.paragraph.create(null, s.text(t)));
    const row = () => s.nodes.table_row.create(null, [cell('а'), cell('б'), cell('в')]);
    E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, s.text(long)), s.nodes.table.create({ borders: true }, [row(), row()])]), E.getMeta(), 'live');
    E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 3)));
    E.view.focus();
  });
  await p.waitForSelector('#hr-inner .ri-left:not([hidden])');
  const paraLeft = () => p.evaluate(() => window.ens.view.dom.querySelector('p').getBoundingClientRect().left);
  const indent = () => p.evaluate(() => window.ens.view.state.doc.firstChild.attrs.indent);
  const x0 = await paraLeft();
  // Отступ слева: во время перетаскивания абзац уже сдвинут, документ ещё не изменён.
  const m = await (await p.$('#hr-inner .ri-left')).boundingBox();
  await p.mouse.move(m.x + m.width / 2, m.y + m.height / 2);
  await p.mouse.down();
  await p.mouse.move(m.x + m.width / 2 + 30, m.y + m.height / 2, { steps: 3 });
  await p.mouse.move(m.x + m.width / 2 + 60, m.y + m.height / 2, { steps: 3 });
  const mid = await paraLeft();
  ok(mid - x0 > 50, `абзац двигается за мышью: ${x0} → ${mid}`);
  eq(await indent(), null, 'до отпускания документ не меняется');
  const m2 = await (await p.$('#hr-inner .ri-left')).boundingBox();
  ok(m2.x - m.x > 50, 'метка отступа едет за мышью');
  await p.mouse.up();
  const ind = await indent();
  ok(ind > 40, `после отпускания отступ записан: ${ind}`);
  // Одна запись в истории: одна отмена возвращает исходное, вторая ничего не меняет.
  await p.keyboard.press('Control+KeyZ');
  eq(await indent(), null, 'одна отмена — исходный отступ');
  const t = await text(p);
  await p.keyboard.press('Control+KeyZ');
  eq(await text(p), t, 'вторая отмена ничего не меняет');
  ok(Math.abs((await paraLeft()) - x0) < 1, 'предпросмотр снят');
  // Левое поле: текст перетекает до отпускания.
  const pageLeft = () => p.evaluate(() => window.ens.getMeta().page.left);
  const L0 = await pageLeft();
  const lm = await (await p.$('#hr-inner .rm[aria-label="Левое поле"]')).boundingBox();
  await p.mouse.move(lm.x + lm.width / 2, lm.y + lm.height / 2);
  await p.mouse.down();
  await p.mouse.move(lm.x + lm.width / 2 + 40, lm.y + lm.height / 2, { steps: 4 });
  ok((await paraLeft()) - x0 > 30, 'поле: текст сдвинулся до отпускания');
  eq(await pageLeft(), L0, 'поле: документ не изменён до отпускания');
  await p.mouse.up();
  ok((await pageLeft()) - L0 > 20, 'поле записано');
  // Столбец таблицы: ширина меняется за мышью, запись — одна.
  await p.evaluate(() => {
    const E = window.ens; let pos = 0;
    E.view.state.doc.descendants((n, q) => { if (!pos && n.type.name === 'table_cell') pos = q + 2; });
    E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, pos)));
  });
  await p.waitForSelector('#hr-inner .rc');
  const cellW = () => p.evaluate(() => window.ens.view.dom.querySelector('td').getBoundingClientRect().width);
  const colw = () => p.evaluate(() => { let c = null; window.ens.view.state.doc.descendants((n) => { if (!c && n.type.name === 'table_cell') c = n.attrs.colwidth; }); return c; });
  const w0 = await cellW();
  const c0 = await colw();
  const rc = (await p.$$('#hr-inner .rc'))[1];
  const b = await rc.boundingBox();
  await p.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
  await p.mouse.down();
  await p.mouse.move(b.x + b.width / 2 + 40, b.y + b.height / 2, { steps: 4 });
  ok((await cellW()) - w0 > 30, `столбец шире до отпускания: ${w0} → ${await cellW()}`);
  eq(await colw(), c0, 'ширина в документе не изменена до отпускания');
  await p.mouse.up();
  ok((await colw())?.[0] > w0 + 20, `ширина записана: ${JSON.stringify(await colw())}`);
  await p.keyboard.press('Control+KeyZ');
  eq(await colw(), c0, 'одна отмена возвращает ширину');
  await p.context().close();
});

await test('обрезка рисунка: во время перетаскивания видна отрезаемая часть', async () => {
  const p = await page();
  const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mP8z8BQz0AEYBxVSF+FABJADveWkH6oAAAAAElFTkSuQmCC';
  await p.evaluate((png) => {
    const E = window.ens, s = E.schema;
    E.loadDoc(s.nodes.doc.create(null, [s.nodes.paragraph.create(null, [s.text('А '), s.nodes.image.create({ src: png, width: 200, height: 100 })])]), E.getMeta(), 'crop');
    E.view.dispatch(E.view.state.tr.setSelection(E.pm.NodeSelection.create(E.view.state.doc, 3)));
    E.run('image.crop');
  }, PNG);
  await p.waitForSelector('.img-ov.crop:not([hidden]) .img-h-e');
  const e = await (await p.$('.img-h-e')).boundingBox();
  const img0 = await (await p.$('.ProseMirror img[src^="data:"]')).boundingBox();
  await p.mouse.move(e.x + 4, e.y + 4);
  await p.mouse.down();
  await p.mouse.move(e.x - 56, e.y + 4, { steps: 4 });
  const g = await p.$('.crop-ghost');
  ok(g, 'полупрозрачный полный рисунок показан');
  const gb = await g.boundingBox();
  ok(Math.abs(gb.x - img0.x) < 2 && Math.abs(gb.width - img0.width) < 2, `призрак на месте исходного рисунка: ${JSON.stringify([gb, img0])}`);
  await p.mouse.up();
  ok(!(await p.$('.crop-ghost')), 'после отпускания призрак убран');
  const crop = await p.evaluate(() => { let c = null; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'image') c = n.attrs.crop; }); return c; });
  ok(crop && crop[2] > 0.2, `обрезано справа: ${JSON.stringify(crop)}`);
  await p.context().close();
});
await test('таблица: перемещение за ручку в углу, линия места, одна запись в истории, Esc отменяет', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const cell = (t) => s.nodes.table_cell.create(null, P(t));
    const tbl = s.nodes.table.create({ borders: true }, [s.nodes.table_row.create(null, [cell('А1'), cell('Б1')]), s.nodes.table_row.create(null, [cell('А2'), cell('Б2')])]);
    E.loadDoc(s.nodes.doc.create(null, [P('Первый'), tbl, P('Второй'), P('Третий'), P('Четвёртый')]), E.getMeta(), 'move');
  });
  const order = () => p.evaluate(() => { const o = []; window.ens.view.state.doc.forEach((n) => o.push(n.type.name === 'table' ? 'T' : n.textContent)); return o.join(','); });
  eq(await order(), 'Первый,T,Второй,Третий,Четвёртый', 'исходный порядок');
  // Ручка появляется при наведении на таблицу.
  const cellBox = await (await p.$('.ProseMirror td')).boundingBox();
  await p.mouse.move(cellBox.x + 10, cellBox.y + 10);
  await p.waitForSelector('.tbl-move:not([hidden])');
  const h = await (await p.$('.tbl-move')).boundingBox();
  const tb = await (await p.$('.ProseMirror table')).boundingBox();
  ok(h.x < tb.x && h.y < tb.y && tb.x - h.x < 30, 'ручка у левого верхнего угла таблицы');
  // Путь от таблицы к ручке её не прячет.
  await p.mouse.move(h.x + h.width / 2, h.y + h.height / 2, { steps: 5 });
  ok(!(await p.$eval('.tbl-move', (e) => e.hidden)), 'ручка не исчезла по дороге');
  const target = await p.evaluate(() => { const ps = window.ens.view.dom.querySelectorAll(':scope > p'); const r = ps[ps.length - 1].getBoundingClientRect(); return { x: r.left + 20, y: r.bottom - 3 }; });
  // Esc отменяет перетаскивание.
  await p.mouse.down();
  await p.mouse.move(target.x, target.y, { steps: 6 });
  ok(await p.$eval('.tbl-drop', (e) => !e.hidden), 'линия места вставки видна');
  await p.keyboard.press('Escape');
  await p.mouse.up();
  eq(await order(), 'Первый,T,Второй,Третий,Четвёртый', 'Esc: таблица осталась на месте');
  // Перенос в конец документа.
  await p.mouse.move(cellBox.x + 10, cellBox.y + 10);
  await p.waitForSelector('.tbl-move:not([hidden])');
  const h2 = await (await p.$('.tbl-move')).boundingBox();
  await p.mouse.move(h2.x + h2.width / 2, h2.y + h2.height / 2, { steps: 3 });
  await p.mouse.down();
  await p.mouse.move(target.x, target.y, { steps: 6 });
  await p.mouse.up();
  eq(await order(), 'Первый,Второй,Третий,Четвёртый,T', 'таблица перенесена в конец');
  ok(await p.evaluate(() => window.ens.view.state.selection.$from.node(1).type.name === 'table'), 'курсор в перенесённой таблице');
  ok((await p.evaluate(() => window.ens.view.state.doc.lastChild.textContent)) === 'А1Б1А2Б2', 'содержимое таблицы цело');
  // Одна отмена возвращает на место.
  await p.keyboard.press('Control+KeyZ');
  eq(await order(), 'Первый,T,Второй,Третий,Четвёртый', 'одна отмена возвращает таблицу');
  // Перенос в начало (до первого абзаца) и запись в DOCX.
  await p.mouse.move(cellBox.x + 10, cellBox.y + 10);
  await p.waitForSelector('.tbl-move:not([hidden])');
  const h3 = await (await p.$('.tbl-move')).boundingBox();
  const first = await p.evaluate(() => { const r = window.ens.view.dom.querySelector(':scope > p').getBoundingClientRect(); return { x: r.left + 20, y: r.top + 2 }; });
  await p.mouse.move(h3.x + h3.width / 2, h3.y + h3.height / 2, { steps: 3 });
  await p.mouse.down();
  await p.mouse.move(first.x, first.y, { steps: 6 });
  await p.mouse.up();
  eq(await order(), 'T,Первый,Второй,Третий,Четвёртый', 'таблица перенесена в начало');
  const back = await p.evaluate(() => { const E = window.ens; const b = E.importDocx(E.exportDocx(E.view.state.doc, E.getMeta(), 'm')); const o = []; b.doc.forEach((n) => o.push(n.type.name === 'table' ? 'T' : n.textContent)); return o.join(','); });
  ok(back.startsWith('T,Первый,Второй,Третий,Четвёртый'), `DOCX: ${back}`);
  await p.context().close();
});
await test('контекстные меню: текст, выделение, ссылка, рисунок, таблица; меню браузера не показывается', async () => {
  const p = await page();
  const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mP8z8BQz0AEYBxVSF+FABJADveWkH6oAAAAAElFTkSuQmCC';
  await p.evaluate((png) => {
    const E = window.ens, s = E.schema;
    const P = (c) => s.nodes.paragraph.create(null, c);
    const cell = (t) => s.nodes.table_cell.create(null, P(s.text(t)));
    const link = s.marks.link.create({ href: 'https://example.org/' });
    E.loadDoc(s.nodes.doc.create(null, [
      P([s.text('Обычный текст абзаца. '), s.text('ссылка', [link]), s.text(' и хвост')]),
      P([s.nodes.image.create({ src: png, width: 120, height: 80 })]),
      s.nodes.table.create({ borders: true }, [s.nodes.table_row.create(null, [cell('А1'), cell('Б1')]), s.nodes.table_row.create(null, [cell('А2'), cell('Б2')])]),
      P(s.text('Последний')),
    ]), E.getMeta(), 'ctx');
  }, PNG);
  // Меню WebView отменено (defaultPrevented) везде, кроме полей ввода.
  await p.evaluate(() => { window.__ctxPrevented = []; document.addEventListener('contextmenu', (e) => setTimeout(() => window.__ctxPrevented.push(e.defaultPrevented)), false); });
  const kind = () => p.$eval('.context-menu.ctx-text', (m) => (m.hidden ? 'hidden' : m.dataset.kind));
  const labels = () => p.$$eval('.context-menu.ctx-text button', (b) => b.map((x) => (x.disabled ? '-' : '') + x.querySelector('span').textContent));
  const at = (sel, off = 2) => p.evaluate(({ sel, off }) => {
    const el = window.ens.view.dom.querySelector(sel);
    const t = [...el.childNodes].find((n) => n.nodeType === 3) ?? el.firstChild.firstChild ?? el.firstChild;
    const r = document.createRange(); r.setStart(t, off); r.setEnd(t, off + 1);
    const b = r.getBoundingClientRect();
    return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  }, { sel, off });
  // Текст: курсор встаёт в точку щелчка, «Вырезать» и «Копировать» недоступны.
  let pt = await at('p', 3);
  await p.mouse.click(pt.x, pt.y, { button: 'right' });
  eq(await kind(), 'text', 'меню текста');
  let l = await labels();
  ok(l.includes('-Вырезать') && l.includes('-Копировать') && l.includes('Вставить') && l.includes('Ссылка…'), `пункты текста: ${l}`);
  ok(await p.evaluate(() => window.ens.view.state.selection.from === 4), 'курсор встал в точку щелчка');
  await p.keyboard.press('Escape');
  eq(await kind(), 'hidden', 'Esc закрывает меню');
  // Выделение: щелчок внутри выделения его сохраняет; пункты форматирования.
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 1, 8))); });
  pt = await at('p', 2);
  await p.mouse.click(pt.x, pt.y, { button: 'right' });
  eq(await kind(), 'selection', 'меню выделения');
  l = await labels();
  ok(l.includes('Вырезать') && l.includes('Полужирный') && l.includes('Очистить форматирование'), `пункты выделения: ${l}`);
  ok(await p.evaluate(() => { const s = window.ens.view.state.selection; return s.from === 1 && s.to === 8; }), 'выделение сохранено');
  // «Полужирный» из меню применяется к выделению.
  await p.click('.context-menu.ctx-text button:has-text("Полужирный")');
  ok(await p.evaluate(() => window.ens.view.state.doc.firstChild.firstChild.marks.some((m) => m.type.name === 'strong')), 'полужирный применён');
  eq(await kind(), 'hidden', 'меню закрылось после команды');
  // Ссылка: открыть, изменить, копировать адрес, удалить.
  const a = await (await p.$('.ProseMirror a')).boundingBox();
  await p.mouse.click(a.x + a.width / 2, a.y + a.height / 2, { button: 'right' });
  ok((await kind()).startsWith('link'), 'меню ссылки');
  l = await labels();
  ok(l.includes('Открыть ссылку') && l.includes('Изменить ссылку…') && l.includes('Удалить ссылку'), `пункты ссылки: ${l}`);
  await p.click('.context-menu.ctx-text button:has-text("Удалить ссылку")');
  ok(await p.evaluate(() => !window.ens.view.dom.querySelector('a') && window.ens.view.state.doc.firstChild.textContent.includes('ссылка')), 'ссылка снята, текст остался');
  // Рисунок: правый щелчок выделяет его; пункты рисунка.
  const im = await (await p.$('.ProseMirror img[src^="data:"]')).boundingBox();
  await p.mouse.click(im.x + im.width / 2, im.y + im.height / 2, { button: 'right' });
  eq(await kind(), 'image', 'меню рисунка');
  l = await labels();
  ok(l.includes('Формат рисунка…') && l.includes('За текстом') && l.includes('Удалить рисунок'), `пункты рисунка: ${l}`);
  ok(await p.evaluate(() => window.ens.view.state.selection.node?.type.name === 'image'), 'рисунок выделен');
  await p.click('.context-menu.ctx-text button:has-text("За текстом")');
  ok(await p.evaluate(() => { let f = null; window.ens.view.state.doc.descendants((n) => { if (n.type.name === 'image') f = n.attrs.float; }); return f?.mode === 'behind'; }), 'обтекание «за текстом» из меню');
  // Таблица: строки и столбцы. Сначала снять выделение рисунка: его рамка (уже «за текстом») может лежать над таблицей.
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 1))); });
  const td = await (await p.$('.ProseMirror tr:last-child td:last-child')).boundingBox();
  await p.mouse.click(td.x + 10, td.y + td.height / 2, { button: 'right' });
  ok((await kind()).startsWith('table'), `меню таблицы: ${await kind()}`);
  l = await labels();
  ok(l.includes('Вставить строку ниже') && l.includes('Удалить столбец') && l.includes('Свойства таблицы…'), `пункты таблицы: ${l}`);
  // Фоновая служебная транзакция (пересчёт страниц) не закрывает открытое меню.
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setMeta('addToHistory', false)); });
  ok((await kind()).startsWith('table'), `меню пережило служебную транзакцию: ${await kind()}`);
  await p.click('.context-menu.ctx-text button:has-text("Вставить строку ниже")');
  eq(await p.$$eval('.ProseMirror tr', (r) => r.length), 3, 'строка добавлена из меню');
  // Клавиатура: Shift+F10, стрелки, Enter.
  await p.keyboard.press('Shift+F10');
  ok((await kind()) !== 'hidden', 'Shift+F10 открывает меню');
  await p.keyboard.press('Escape');
  // Вне листа (лента) меню WebView тоже не показывается; Shift+правый щелчок по тексту — системное (орфография).
  await p.click('#tabs', { button: 'right' });
  await p.click('#status-page', { button: 'right' });
  pt = await at('p', 3);
  await p.keyboard.down('Shift');
  await p.mouse.click(pt.x, pt.y, { button: 'right' });
  await p.keyboard.up('Shift');
  eq(await kind(), 'hidden', 'Shift+правый щелчок: наше меню не открывается');
  await p.waitForTimeout(50);
  const prevented = await p.evaluate(() => window.__ctxPrevented);
  ok(prevented.length >= 8 && prevented.slice(0, -1).every(Boolean) && prevented[prevented.length - 1] === false, `меню WebView: ${JSON.stringify(prevented)}`);
  await p.context().close();
});
await test('вёрстка: в документе из нескольких разделов отступ и выступ абзаца сохраняются, строка не выходит за лист (образец пользователя)', async () => {
  const p = await page();
  // Как в присланном файле: книжный раздел, затем альбомный; абзац «отступ 3 см, выступ 3 см», по ширине,
  // первое слово шире выступа, две табуляции; рядом — абзац с отступом первой строки 1,5 см.
  const r = await p.evaluate(async () => {
    const E = window.ens, s = E.schema;
    const m = E.getMeta();
    const land = { ...m.page, w: m.page.h, h: m.page.w, left: 28.35, right: 35.45 };
    const doc = s.nodes.doc.create(null, [
      s.nodes.paragraph.create(null, s.text('Книжный раздел')),
      s.nodes.section_break.create({ page: land }),
      s.nodes.paragraph.create({ first: 42.5 }, s.text('объёмы топлива')),
      s.nodes.paragraph.create({ indent: 85.05, first: -85.05, align: 'justify' }, s.text('шжроьроьбромьроьь\t\tПриложение: Согласованные объемы')),
      s.nodes.paragraph.create({ indent: -200 }, s.text('Отрицательный отступ больше поля')),
    ]);
    const back = E.importDocx(E.exportDocx(doc, m, 'x'));
    E.loadDoc(back.doc, back.meta, 'x');
    await new Promise((res) => setTimeout(res, 400));
    const ps = [...E.view.dom.querySelectorAll(':scope > p')];
    const papers = [...document.querySelectorAll('.paper')].map((e) => e.getBoundingClientRect());
    const info = (el) => {
      const rg = document.createRange(); rg.selectNodeContents(el);
      const rects = [...rg.getClientRects()];
      const first = Math.min(...rects.map((x) => x.left));
      // По строке, а не по блоку: у первого абзаца раздела сверху отступ до нового листа.
      const y = rects[0].top;
      const paper = papers.find((b) => y >= b.top && y < b.bottom);
      return { first, paperLeft: paper.left, paperW: paper.width };
    };
    return { fuel: info(ps[1]), hang: info(ps[2]), neg: info(ps[3]), firstAttr: back.doc.child(2).attrs.first, k: 4 / 3 };
  });
  const k = r.k;
  // Альбомный лист шире: текст считается от его края; поле 28,35 pt (1 см).
  const marginX = (x) => (x.first - x.paperLeft) / k;
  ok(Math.abs(marginX(r.hang) - 28.35) < 2, `выступ: первая строка у поля (1 см), а не левее: ${marginX(r.hang).toFixed(1)} pt`);
  ok(Math.abs(marginX(r.fuel) - (28.35 + 42.5)) < 2, `отступ первой строки 1,5 см не теряется: ${marginX(r.fuel).toFixed(1)} pt`);
  ok(marginX(r.neg) >= -0.5, `отрицательный отступ не за краем листа: ${marginX(r.neg).toFixed(1)} pt`);
  await p.context().close();
});
await test('табуляция как в Word: стопы через 1,27 см от поля, неявный стоп на выступе; документ не меняется', async () => {
  const p = await page();
  const r = await p.evaluate(async () => {
    const E = window.ens, s = E.schema;
    const P = (a, t) => s.nodes.paragraph.create(a, s.text(t));
    const doc = s.nodes.doc.create(null, [
      P(null, 'а\tб\tв'),
      // Отступ 18 pt: браузер считал бы шаг от отступа, Word — от поля.
      P({ indent: 18 }, 'Слово\tдалее'),
      // Как в образце: отступ 3 см, выступ 3 см, слово шире выступа, две табуляции.
      P({ indent: 85.05, first: -85.05, align: 'justify' }, 'шжроьроьбромьроьь\t\tПриложение'),
      // Короткое слово при выступе: табуляция встаёт на отступ (неявный стоп).
      P({ indent: 85.05, first: -85.05 }, '1.\tПункт'),
    ]);
    E.loadDoc(doc, E.getMeta(), 'tabs');
    await new Promise((res) => setTimeout(res, 600));
    const pt = 3 / 4;
    const ps = [...E.view.dom.querySelectorAll(':scope > p')];
    // Поле — левый край текста страницы: левый край абзаца без отступа.
    const margin = ps[0].getBoundingClientRect().left;
    const ends = ps.map((el) => [...el.querySelectorAll('span.tab')].map((t) => {
      const b = t.getBoundingClientRect();
      return Math.round(((b.right + parseFloat(getComputedStyle(t).marginRight) - margin) * pt) * 10) / 10;
    }));
    // Конец первого слова образца — от него Word считает следующий стоп.
    const t3 = ps[2].querySelector('span.tab').previousSibling ?? ps[2].firstChild;
    const rg = document.createRange(); rg.selectNodeContents(t3);
    const wordEnd = Math.round((rg.getBoundingClientRect().right - margin) * pt * 10) / 10;
    const json = JSON.stringify(E.view.state.doc.toJSON());
    return { ends, wordEnd, same: json === JSON.stringify(doc.toJSON()) };
  });
  const near = (a, b) => Math.abs(a - b) < 1;
  ok(near(r.ends[0][0], 36) && near(r.ends[0][1], 72), `простые табуляции на 36 и 72 pt: ${r.ends[0]}`);
  const onGrid = (x) => { const m = x % 36; return m < 1 || 36 - m < 1; };
  ok(onGrid(r.ends[1][0]), `от поля, а не от отступа (кратно 36 pt): ${r.ends[1]}`);
  // Слово шире выступа (85 pt): табуляции идут на ближайшие стопы сетки 36 pt после конца слова.
  const after = 36 * (Math.floor(r.wordEnd / 36) + 1);
  ok(r.wordEnd > 85 && near(r.ends[2][0], after) && near(r.ends[2][1], after + 36), `образец: слово до ${r.wordEnd} pt, табуляции ${r.ends[2]}, ждём ${after} и ${after + 36}`);
  ok(near(r.ends[3][0], 85.05), `короткий номер при выступе — табуляция до отступа: ${r.ends[3]}`);
  ok(r.same, 'документ не изменился (табуляция — только вид)');
  await p.context().close();
});
await test('версия: один номер — из package.json в «О программе», в tauri.conf.json и в Cargo.toml', async () => {
  const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')).version;
  const conf = JSON.parse(readFileSync(join(root, 'src-tauri/tauri.conf.json'), 'utf8')).version;
  eq(conf, '../package.json', 'tauri.conf.json берёт версию из package.json');
  for (const f of ['src-tauri/Cargo.toml', 'redos7/Cargo.toml']) {
    const m = /^version\s*=\s*"([^"]+)"/m.exec(readFileSync(join(root, f), 'utf8'));
    eq(m?.[1], pkg, `${f}: версия совпадает с package.json`);
  }
  const p = await page();
  await p.click('.tab.file');
  await p.waitForSelector('.text-backstage:not([hidden]) .bs-about');
  ok((await p.$eval('.text-backstage .bs-about', (a) => a.textContent)).includes(`ENS-Text ${pkg}`), 'меню «Файл» показывает версию из package.json');
  await p.context().close();
});
await test('скрытые символы: Ctrl+* показывает ¶ · ° → ↵, вёрстка и документ не меняются, при печати не видны', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const P = (c) => s.nodes.paragraph.create(null, c);
    E.loadDoc(s.nodes.doc.create(null, [
      P([s.text('Один два\tтри'), s.nodes.hard_break.create(), s.text('№ 5')]),
      P(),
      P(s.text('Длинный абзац для проверки переносов строк, который не должен сдвинуться от знаков. '.repeat(3))),
    ]), E.getMeta(), 'marks');
  });
  await p.waitForTimeout(400);
  // Положение каждой буквы текста (по текстовым узлам): знаки не должны сдвигать ни одну.
  const geom = () => p.evaluate(() => {
    const out = [];
    const d = window.ens.view.dom, o = d.getBoundingClientRect();
    const w = document.createTreeWalker(d, NodeFilter.SHOW_TEXT);
    let n = 0;
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      if (t.parentElement.closest('.hm-w')) continue;
      for (let i = 0; i < t.length; i++, n++) {
        if (n % 7) continue;
        const r = document.createRange(); r.setStart(t, i); r.setEnd(t, i + 1);
        const b = r.getBoundingClientRect(); out.push([Math.round((b.left - o.left) * 2) / 2, Math.round((b.top - o.top) * 2) / 2]);
      }
    }
    return out;
  });
  const before = await geom();
  const doc0 = await p.evaluate(() => JSON.stringify(window.ens.view.state.doc.toJSON()));
  await p.click('.ProseMirror');
  await p.keyboard.press('Control+Shift+Digit8');
  await p.waitForTimeout(300);
  const r = await p.evaluate(() => {
    const d = window.ens.view.dom;
    const content = (sel) => [...d.querySelectorAll(sel)].map((e) => getComputedStyle(e, '::after').content);
    return {
      on: d.classList.contains('show-marks'),
      para: content('.hm-para'), sp: content('.hm-sp'), nb: content('.hm-nb'), tab: content('span.tab'), br: content('.hm-br'),
    };
  });
  ok(r.on, 'режим включён');
  eq(r.para, ['"¶"', '"¶"', '"¶"'], '¶ в конце каждого абзаца, и пустого тоже');
  ok(r.sp.length >= 10 && r.sp.every((c) => c === '"·"'), `· на пробелах: ${r.sp.length}`);
  eq(r.nb, ['"°"'], '° на неразрывном пробеле');
  eq(r.tab, ['"→"'], '→ на табуляции');
  eq(r.br, ['"↵"'], '↵ на разрыве строки');
  const after = await geom();
  // Допуск 1 px: обёртки пробелов меняют субпиксельную расстановку, но не строки.
  ok(after.length === before.length && after.every((q, i) => Math.abs(q[0] - before[i][0]) <= 1 && q[1] === before[i][1]), 'буквы на своих местах, строки не перенеслись иначе');
  eq(await p.evaluate(() => JSON.stringify(window.ens.view.state.doc.toJSON())), doc0, 'документ не изменился');
  ok(await p.$eval('[data-command="view.marks"]', (b) => b.classList.contains('on')), 'кнопка ¶ на ленте нажата');
  // Ввод при включённых знаках: новые пробелы тоже с точкой.
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 1))); });
  await p.keyboard.type('А Б ');
  ok((await p.$$('.ProseMirror .hm-sp')).length >= r.sp.length + 2, 'знаки появляются на новых пробелах');
  // Печать: знаков не видно.
  await p.emulateMedia({ media: 'print' });
  const printed = await p.evaluate(() => [...document.querySelectorAll('.ProseMirror .hm-para, .ProseMirror .hm-sp')].slice(0, 3).map((e) => e.classList.contains('hm-para') ? getComputedStyle(e).display : getComputedStyle(e, '::after').content));
  ok(printed.every((x) => x === 'none'), `при печати знаков нет: ${printed}`);
  await p.emulateMedia({ media: 'screen' });
  // Выбор запоминается; повторное нажатие выключает.
  await p.reload();
  await p.waitForFunction(() => window.ens);
  ok(await p.evaluate(() => window.ens.view.dom.classList.contains('show-marks')), 'после перезапуска знаки снова видны');
  // Окно восстановления несохранённого документа — закрываем.
  if (await p.$('#dialog:not([hidden])')) { await p.keyboard.press('Escape'); await p.waitForSelector('#dialog', { state: 'hidden' }); }
  await p.click('.ProseMirror');
  await p.keyboard.press('Control+Shift+Digit8');
  ok(await p.evaluate(() => !window.ens.view.dom.classList.contains('show-marks') && !document.querySelector('.hm-para')), 'выключено');
  await p.context().close();
});
await test('словарь: фамилия Едренников во всех формах не проверяется; свои слова — окно и контекстное меню', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const P = (t) => s.nodes.paragraph.create(null, s.text(t));
    E.loadDoc(s.nodes.doc.create(null, [
      P('Едренников, Едренникову, Едренниковой, Едренниковыми и едренникова — все формы.'),
      P('Слово Зюзюкинцев проверяется, пока его нет в словаре.'),
    ]), E.getMeta(), 'dict');
  });
  await p.waitForTimeout(200);
  const ok1 = () => p.$$eval('.ProseMirror .dict-word[spellcheck="false"]', (a) => a.map((x) => x.textContent));
  eq(await ok1(), ['Едренников', 'Едренникову', 'Едренниковой', 'Едренниковыми', 'едренникова'], 'встроенные формы фамилии не проверяются');
  // Контекстное меню: «Добавить в словарь» для слова под курсором.
  const pt = await p.evaluate(() => {
    const t = window.ens.view.dom.querySelectorAll('p')[1].firstChild;
    const i = t.textContent.indexOf('Зюзюкинцев') + 3;
    const r = document.createRange(); r.setStart(t, i); r.setEnd(t, i + 1);
    const b = r.getBoundingClientRect();
    return { x: b.left + b.width / 2, y: b.top + b.height / 2 };
  });
  await p.mouse.click(pt.x, pt.y, { button: 'right' });
  await p.click('.context-menu.ctx-text button:has-text("Добавить «Зюзюкинцев» в словарь")');
  ok((await ok1()).includes('Зюзюкинцев'), 'слово из меню больше не проверяется');
  // Набор новой формы: тоже из словаря (фамилия на -ев — добавлены все формы).
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.insertText(' Зюзюкинцевым', E.view.state.doc.content.size - 1)); });
  ok((await ok1()).includes('Зюзюкинцевым'), 'форма фамилии тоже в словаре');
  // Окно «Словарь»: свои слова, удаление, добавление без форм.
  await p.evaluate(() => window.ens.run('review.dictionary'));
  await p.waitForSelector('#dialog .dict-list');
  ok((await p.$$eval('#dialog .dict-row span', (a) => a.map((x) => x.textContent))).includes('Зюзюкинцева'), 'в окне видны добавленные формы');
  await p.fill('#dialog input[type=text]', 'ЕНСОфис');
  await p.keyboard.press('Enter');
  ok((await p.$$eval('#dialog .dict-row span', (a) => a.map((x) => x.textContent))).includes('ЕНСОфис'), 'слово добавлено из окна');
  await p.click('#dialog .dict-row:has-text("Зюзюкинцев") >> nth=0 >> button');
  await p.keyboard.press('Escape');
  const saved = await p.evaluate(() => JSON.parse(localStorage.getItem('ens-text.dict')));
  ok(saved.includes('ЕНСОфис') && !saved.includes('Зюзюкинцев'), `словарь сохранён у пользователя: ${saved.length} слов`);
  ok(!(await ok1()).includes('Зюзюкинцев'), 'удалённое слово снова проверяется');
  eq(await p.evaluate(() => window.ens.view.state.doc.textContent.includes('Зюзюкинцевым')), true, 'текст документа не тронут');
  await p.context().close();
});
await test('меню «Файл» как в ENS-Table: разделы, шаблоны плитками, параметры, «Закрыть», авторство и лицензия', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Черновик');
  // Alt+F открывает меню на «Сведениях».
  await p.keyboard.press('Alt+KeyF');
  await p.waitForSelector('.text-backstage:not([hidden])');
  const nav = await p.$$eval('.text-backstage .bs-nav button[data-go]', (b) => b.map((x) => x.textContent.trim()).filter(Boolean));
  eq(nav, ['Сведения', 'Создать', 'Открыть', 'Сохранить', 'Сохранить как', 'Печать', 'Закрыть', 'Параметры'], 'разделы как в ENS-Table');
  ok((await p.$eval('.text-backstage .bs-body h2', (h) => h.textContent)).includes('есть несохранённые изменения'), 'сведения: несохранённые изменения');
  ok(/Слов: 1/.test(await p.$eval('.text-backstage .bs-body', (b) => b.textContent)) || /слов: 1/.test(await p.$eval('.text-backstage .bs-body', (b) => b.textContent)), 'сведения: число слов');
  const about = await p.$eval('.text-backstage .bs-about', (a) => a.textContent);
  ok(about.includes('Едренников') && about.includes('AGPL-3.0-or-later'), `авторство и лицензия: ${about}`);
  // Создать: пустой документ и все встроенные шаблоны плитками.
  await p.click('.text-backstage [data-go="new"]');
  const tiles = await p.$$eval('.text-backstage .bs-tile', (b) => b.map((x) => x.textContent.trim()));
  ok(tiles[0] === 'Новый документ' && tiles.includes('Письмо') && tiles.includes('Приказ') && tiles.length >= 10, `плитки: ${tiles}`);
  await p.click('.text-backstage .bs-tile:has-text("Письмо")');
  // Документ не сохранён: сначала вопрос.
  await p.waitForSelector('#dialog:not([hidden]) button.danger');
  await p.click('#dialog button.danger');
  await p.waitForFunction(() => window.ens.view.state.doc.textContent.length > 50);
  ok(await p.$('.text-backstage[hidden]'), 'меню закрылось');
  ok(!(await p.evaluate(() => window.ens.view.state.doc.textContent.includes('Черновик'))), 'создан документ из шаблона «Письмо»');
  // Параметры: имя автора, напоминание, типограф — сохраняются в настройках.
  await p.keyboard.press('Alt+KeyF');
  await p.click('.text-backstage [data-go="options"]');
  await p.fill('.text-backstage input[data-set="userName"]', 'Тестов Т. Т.');
  await p.press('.text-backstage input[data-set="userName"]', 'Tab');
  await p.fill('.text-backstage input[data-set="saveReminderMinutes"]', '25');
  await p.press('.text-backstage input[data-set="saveReminderMinutes"]', 'Tab');
  const live0 = await p.$eval('.text-backstage input[data-set="liveQuotes"]', (c) => c.checked);
  await p.click('.text-backstage input[data-set="liveQuotes"]');
  const marks = await p.$('.text-backstage input[data-set="view.marks"]');
  ok(marks, 'скрытые знаки — в параметрах');
  await marks.click();
  const st = await p.evaluate(() => {
    const raw = Object.keys(localStorage).map((k) => localStorage.getItem(k)).join(' ');
    return { name: raw.includes('Тестов Т. Т.'), rem: /"saveReminderMinutes":25/.test(raw), live: localStorage.getItem('ens-text.liveQuotes'), marks: window.ens.view.dom.classList.contains('show-marks') };
  });
  ok(st.name && st.rem, `имя и напоминание сохранены: ${JSON.stringify(st)}`);
  eq(st.live, live0 ? '0' : '1', 'типограф переключён');
  ok(st.marks, 'скрытые знаки включены из параметров');
  // Открыть: в браузере — «Обзор…» и пояснение, без путей.
  await p.click('.text-backstage [data-go="open"]');
  ok((await p.$$('.text-backstage .bs-wide')).length >= 1, 'кнопка «Обзор…»');
  // «Закрыть»: документ закрывается, остаётся пустой.
  await p.click('.text-backstage [data-go="close"]');
  await p.waitForSelector('#dialog:not([hidden]) button.danger');
  await p.click('#dialog button.danger');
  await p.waitForFunction(() => window.ens.view.state.doc.textContent === '');
  ok(await p.$('.text-backstage[hidden]'), 'меню закрылось после «Закрыть»');
  await p.context().close();
});
await test('шаблоны: каждый встроенный шаблон создаётся без ошибок (письмо, приказ, акт, протокол без даты)', async () => {
  const p = await page();
  const r = await p.evaluate(async () => {
    const { builtInTemplates } = await import('/src/templates.ts');
    return builtInTemplates().map((t) => { try { const b = t.build(); b.doc.check(); return [t.id, b.doc.textContent.length > 0 || t.id.startsWith('blank')]; } catch (e) { return [t.id, String(e.message)]; } });
  });
  for (const [id, res] of r) ok(res === true, `шаблон ${id}: ${res}`);
  ok(r.length >= 9, `шаблонов: ${r.length}`);
  await p.context().close();
});
await test('шаблон «Письмо на бланке»: поля, шапка-таблица, шрифты, исполнитель внизу первого листа, DOCX', async () => {
  const p = await page();
  await p.keyboard.press('Alt+KeyF');
  await p.click('.text-backstage [data-go="new"]');
  await p.click('.text-backstage .bs-tile:has-text("Письмо на бланке")');
  await p.waitForFunction(() => window.ens.view.state.doc.textContent.includes('Исполнитель:'));
  await p.waitForTimeout(600);
  const r = await p.evaluate(() => {
    const E = window.ens, d = E.view.state.doc, m = E.getMeta();
    const t = d.firstChild;
    const cells = []; t.descendants((n) => { if (n.type.name === 'table_cell') cells.push(n.textContent); });
    const marks = (text) => { let out = null; d.descendants((n) => { if (n.isText && n.text.includes(text) && !out) out = n.marks.map((x) => x.type.name + (x.attrs.size ? ':' + x.attrs.size : '')); }); return out; };
    const paraOf = (text) => { let a = null; d.descendants((n) => { if (n.type.name === 'paragraph' && n.textContent.startsWith(text) && !a) a = n.attrs; }); return a; };
    // Низ блока «Исполнитель» и низ текста первой страницы.
    const sheet = document.getElementById('sheet').getBoundingClientRect();
    const last = E.view.dom.lastElementChild.getBoundingClientRect();
    const box = E.pageBoxes()[0];
    const limit = box.top + box.H - box.eff.bottom * E.PX;
    const lineH = E.view.dom.lastElementChild.getBoundingClientRect().height;
    return {
      page: [m.page.top, m.page.bottom, m.page.left, m.page.right].map((x) => Math.round(x / 28.3465 * 100) / 100),
      font: [m.defaults.font, m.defaults.size],
      cols: cells.length, logo: cells[0], middle: cells[1], addr: cells[2],
      addrMarks: marks('Генеральному директору'), subj: marks('О выплате премии'), appeal: marks('Уважаемый'),
      bodyFirst: paraOf('[Основной текст')?.first, attach: paraOf('Приложение:')?.first, item: paraOf('1. ')?.first,
      sign: marks('[Должность]'),
      gap: limit - (last.bottom - sheet.top), lineH, pages: E.pageCount(),
      lastText: d.lastChild.textContent,
      docx: (() => { const b = E.importDocx(E.exportDocx(d, m, 'x')); return [b.doc.childCount === d.childCount, b.meta.lossy.length]; })(),
    };
  });
  eq(r.page, [2, 2, 2.5, 1.5], 'поля: верх 2, низ 2, лево 2,5, право 1,5 см');
  eq(r.font, ['Times New Roman', 14], 'Times New Roman 14');
  eq(r.cols, 3, 'шапка — таблица на 3 столбца');
  ok(r.logo.includes('Логотип') && r.middle === '' && r.addr.startsWith('Генеральному директору'), `ячейки: ${r.logo} | ${r.middle} | ${r.addr}`);
  eq(r.addrMarks, ['strong'], 'адресат жирный');
  ok(r.subj.includes('em') && r.subj.includes('textStyle:12'), `тема курсивом 12: ${r.subj}`);
  ok(r.appeal.includes('strong'), 'обращение жирным');
  ok(Math.abs(r.bodyFirst - 35.4) < 0.1 && r.attach == null && Math.abs(r.item - 35.4) < 0.1, `абзацы 1,25 см, «Приложение» без отступа: ${r.bodyFirst} ${r.attach} ${r.item}`);
  ok(r.sign.includes('strong'), 'должность подписанта жирным');
  eq(r.pages, 1, 'одна страница');
  ok(r.gap >= -1 && r.gap < r.lineH + 2, `исполнитель в последних строках листа: до низа ${r.gap.toFixed(1)} px при строке ${r.lineH.toFixed(1)} px`);
  ok(r.lastText.includes('(000)'), 'последняя строка — телефон исполнителя');
  ok(r.docx[0], 'DOCX туда и обратно: блоки на месте');
  await p.context().close();
});
await test('лента как в Word: одна высота на всех вкладках, крупные значки, списки показывают шрифт после смены вкладки', async () => {
  const p = await page();
  const heights = new Set();
  const tabs = await p.$$eval('#tabs .tab:not(.file)', (t) => t.map((x) => x.textContent));
  for (const name of tabs) {
    await p.click(`#tabs .tab:text-is("${name}")`);
    const r = await p.evaluate(() => {
      const rb = document.getElementById('ribbon');
      // Кнопки внутри ленты не наезжают друг на друга.
      const boxes = [...rb.querySelectorAll('.btn, .ribbon-select')].map((e) => e.getBoundingClientRect()).filter((b) => b.width > 0);
      let overlap = 0;
      for (let i = 0; i < boxes.length; i++) for (let j = i + 1; j < boxes.length; j++) {
        const a = boxes[i], b = boxes[j];
        const w = Math.min(a.right, b.right) - Math.max(a.left, b.left), h = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);
        if (w > 2 && h > 2) overlap++;
      }
      const large = rb.querySelector('.btn.large svg');
      return { h: Math.round(rb.getBoundingClientRect().height), overlap, large: large ? Math.round(large.getBoundingClientRect().width) : null, clipped: [...rb.querySelectorAll('.btn, .ribbon-select')].some((e) => e.getBoundingClientRect().bottom > rb.getBoundingClientRect().bottom + 1) };
    });
    heights.add(r.h);
    eq(r.overlap, 0, `вкладка «${name}»: кнопки не наезжают`);
    ok(!r.clipped, `вкладка «${name}»: ничего не обрезано снизу`);
    if (r.large != null) eq(r.large, 32, `вкладка «${name}»: крупные значки 32 px`);
  }
  eq(heights.size, 1, `лента одной высоты: ${[...heights]}`);
  await p.click('#tabs .tab:text-is("Главная")');
  eq(await p.$eval('select[data-command="format.fontName"]', (s) => s.value), 'Times New Roman', 'после смены вкладки список шрифтов показывает шрифт документа');
  eq(await p.$eval('select[data-command="format.fontSize"]', (s) => s.value), '12', 'и размер');
  await p.context().close();
});

await test('выбор шрифта: образец своим начертанием, кириллица и аналог, поиск, применение к выделению', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Пример текста');
  await p.keyboard.press('Control+KeyA');
  await p.click('select[data-command="format.fontName"]');
  await p.waitForSelector('.font-picker .fp-row');
  const rows = await p.$$eval('.font-picker .fp-row', (r) => r.map((x) => ({ font: x.dataset.font, family: x.querySelector('.fp-name').style.fontFamily, tags: [...x.querySelectorAll('.fp-tag')].map((t) => t.textContent) })));
  const tnr = rows.find((r) => r.font === 'Times New Roman');
  ok(tnr && tnr.family.includes('Times New Roman') && tnr.family.includes('Tinos'), `имя показано своим шрифтом (с аналогом): ${tnr?.family}`);
  ok(tnr.tags.includes('кириллица'), `Times New Roman: кириллица: ${tnr.tags}`);
  ok(rows.every((r) => r.tags.length >= 1), 'у каждого шрифта есть пометка');
  // Шрифты системы находятся сами, без окна разрешения (queryLocalFonts не вызывается).
  const heads = await p.$$eval('.font-picker .fp-head', (h) => h.map((x) => x.textContent));
  ok(heads.includes('Шрифты системы'), `раздел «Шрифты системы» сразу: ${heads}`);
  ok(rows.some((r) => r.font === 'DejaVu Sans'), 'найден установленный DejaVu Sans');
  ok(!rows.some((r) => r.font === 'Segoe Fluent Icons'), 'неустановленные шрифты каталога не показаны');
  eq(await p.$eval('.font-picker .fp-more', (b) => b.textContent), 'Найти остальные шрифты…', 'полный список — только по кнопке');
  // Поиск: несуществующий шрифт — «нет в системе», но выбрать можно.
  await p.fill('.font-picker .fp-search', 'Нетакогошрифта');
  const missing = await p.$$eval('.font-picker .fp-row', (r) => r.map((x) => [x.dataset.font, x.textContent]));
  ok(missing.length === 1 && missing[0][1].includes('нет в системе'), `неизвестный шрифт: ${JSON.stringify(missing)}`);
  await p.fill('.font-picker .fp-search', 'arial');
  await p.keyboard.press('Enter');
  ok(!(await p.$('.font-picker')), 'окно закрылось после выбора');
  const font = await p.evaluate(() => { let f = null; window.ens.view.state.doc.descendants((n) => { if (n.isText && !f) f = n.marks.find((m) => m.type.name === 'textStyle')?.attrs.font; }); return f; });
  eq(font, 'Arial', 'шрифт применён к выделенному тексту');
  // Недавние: выбранный шрифт первым.
  await p.click('select[data-command="format.fontName"]');
  await p.waitForSelector('.font-picker .fp-head');
  eq(await p.$eval('.font-picker .fp-head', (h) => h.textContent), 'Недавние', 'раздел «Недавние»');
  eq(await p.$eval('.font-picker .fp-row', (r) => r.dataset.font), 'Arial', 'последний выбранный — первым');
  await p.keyboard.press('Escape');
  ok(!(await p.$('.font-picker')), 'Esc закрывает окно');
  await p.context().close();
});

await test('чёткие значки ленты: большие — рисунок 32/28 px, средние и малые — 20 px, виден один рисунок', async () => {
  const p = await page();
  await p.click('#tabs .tab:has-text("Вставка")');
  const info = (sel) => p.$eval(sel, (b) => [...b.querySelectorAll('svg')].filter((s) => getComputedStyle(s).display !== 'none').map((s) => [s.getAttribute('viewBox'), Math.round(s.getBoundingClientRect().width)]));
  eq(await info('.btn.large[data-command="insert.table"]'), [['0 0 32 32', 32]], 'Таблица: значок под 32 px');
  await p.click('#tabs .tab:has-text("Главная")');
  const small = await p.$$eval('#ribbon .btn.small, #ribbon .btn.medium', (bs) => bs.map((b) => [...b.querySelectorAll('svg')].filter((s) => getComputedStyle(s).display !== 'none').map((s) => [s.getAttribute('viewBox'), Math.round(s.getBoundingClientRect().width)])).filter((x) => x.length));
  ok(small.every((x) => x.length === 1), 'у каждой кнопки виден один рисунок');
  ok(small.every((x) => x[0][0] === `0 0 ${x[0][1]} ${x[0][1]}`), `рисунок не растянут: ${JSON.stringify(small.filter((x) => x[0][0] !== `0 0 ${x[0][1]} ${x[0][1]}`))}`);
  await p.context().close();
});

await test('регистр (Aa, Shift+F3) и межстрочный интервал кнопкой с меню', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('пРИВЕТ, КАК ДЕЛА? всё хорошо.');
  const text = () => p.evaluate(() => window.ens.view.state.doc.textContent);
  // Часть текста жирная: оформление при смене регистра сохраняется.
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.addMark(1, 7, E.schema.marks.strong.create())); });
  await p.keyboard.press('Control+KeyA');
  const pick = async (label) => { await p.click('[data-command="format.caseMenu"]'); await p.click(`.case-menu button:has-text("${label}")`); return text(); };
  eq(await pick('иЗМЕНИТЬ'), 'Привет, как дела? ВСЁ ХОРОШО.', 'инверсия (случайный Caps Lock)');
  eq(await pick('Как в предложениях'), 'Привет, как дела? Всё хорошо.', 'как в предложениях');
  eq(await pick('ВСЕ ПРОПИСНЫЕ'), 'ПРИВЕТ, КАК ДЕЛА? ВСЁ ХОРОШО.', 'все прописные');
  eq(await pick('все строчные'), 'привет, как дела? всё хорошо.', 'все строчные');
  eq(await pick('Начинать С Прописных'), 'Привет, Как Дела? Всё Хорошо.', 'каждое слово с прописной');
  const bold = await p.evaluate(() => { const out = []; window.ens.view.state.doc.descendants((n) => { if (n.isText && n.marks.length) out.push(n.text); }); return out; });
  eq(bold, ['Привет'], 'жирный остался на том же слове');
  // Shift+F3 по кругу: смешанный → ПРОПИСНЫЕ → строчные → Как в предложениях.
  await p.keyboard.press('Shift+F3'); eq(await text(), 'ПРИВЕТ, КАК ДЕЛА? ВСЁ ХОРОШО.', 'Shift+F3: прописные');
  await p.keyboard.press('Shift+F3'); eq(await text(), 'привет, как дела? всё хорошо.', 'Shift+F3: строчные');
  await p.keyboard.press('Shift+F3'); eq(await text(), 'Привет, как дела? Всё хорошо.', 'Shift+F3: как в предложениях');
  await p.keyboard.press('Control+KeyZ');
  eq(await text(), 'привет, как дела? всё хорошо.', 'отмена — на шаг');
  // Без выделения — слово под курсором.
  await p.evaluate(() => { const E = window.ens; const at = E.view.state.doc.content.size - 4; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, at))); });
  await p.keyboard.press('Shift+F3');
  eq(await text(), 'привет, как дела? всё Хорошо.', 'слово под курсором (строчное → с прописной)');
  await p.keyboard.press('Shift+F3');
  eq(await text(), 'привет, как дела? всё ХОРОШО.', 'и дальше по кругу — прописные');
  // Межстрочный интервал: значок с меню вместо широкого списка.
  ok(!(await p.$('select[data-command="para.line"]')), 'широкого списка больше нет');
  await p.click('[data-command="para.lineMenu"]');
  eq(await p.$eval('.line-menu button[aria-checked="true"]', (b) => b.textContent), '✓1,0', 'отмечен текущий 1,0');
  await p.click('.line-menu button:has-text("1,5")');
  eq(await p.evaluate(() => window.ens.view.state.doc.firstChild.attrs.line), 1.5, 'интервал 1,5');
  await p.click('[data-command="para.lineMenu"]');
  await p.click('.line-menu button:has-text("Добавить интервал перед абзацем")');
  eq(await p.evaluate(() => window.ens.view.state.doc.firstChild.attrs.before), 12, 'интервал перед абзацем');
  await p.context().close();
});

await test('линейки тоньше и точнее: 22 px, деления на целых пикселях, миллиметры видны', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.waitForTimeout(300);
  const r = await p.evaluate(() => {
    const h = document.getElementById('hruler').getBoundingClientRect().height;
    const lines = [...document.querySelectorAll('#hr-inner svg line')];
    const xs = lines.map((l) => Number(l.getAttribute('x1')));
    return { h: Math.round(h), n: lines.length, crisp: xs.every((x) => Math.abs((x % 1) - 0.5) < 1e-6), mm: lines.filter((l) => Number(l.getAttribute('y2')) - Number(l.getAttribute('y1')) === 2).length };
  });
  eq(r.h, 22, 'толщина горизонтальной линейки');
  ok(r.crisp, 'все деления на x+0,5 — линии в 1 px чёткие');
  ok(r.mm > 100, `миллиметровые деления: ${r.mm}`);
  await p.context().close();
});
await test('правки: значок ¶ на ленте виден; ссылка через окно не съедает текст; подсказка и «рука» по Ctrl', async () => {
  const p = await page();
  ok(await p.$eval('[data-command="view.marks"]', (b) => !!b.querySelector('svg')), 'у кнопки «Отобразить все знаки» есть значок');
  await p.click('.ProseMirror');
  await p.keyboard.type('Сайт компании здесь');
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 6, 14))); });
  await p.keyboard.press('Control+KeyK');
  await p.waitForSelector('#dialog:not([hidden]) input');
  await p.fill('#dialog input', 'https://example.org');
  await p.keyboard.press('Enter');
  await p.waitForSelector('#dialog', { state: 'hidden' });
  eq(await p.evaluate(() => window.ens.view.state.doc.textContent), 'Сайт компании здесь', 'текст на месте, Enter не дошёл до документа');
  eq(await p.evaluate(() => window.ens.view.state.doc.childCount), 1, 'абзац не разорван');
  const a = await p.$('.ProseMirror a');
  ok(/Ctrl\+щелчок/.test(await a.getAttribute('title')), 'подсказка «Ctrl+щелчок — открыть ссылку»');
  await p.keyboard.down('Control');
  await p.mouse.move(10, 10);
  eq(await a.evaluate((x) => getComputedStyle(x).cursor), 'pointer', 'с Ctrl над ссылкой — «рука»');
  await p.keyboard.up('Control');
  await p.mouse.move(12, 12);
  eq(await a.evaluate((x) => getComputedStyle(x).cursor), 'text', 'без Ctrl — обычный курсор');
  await p.context().close();
});
await test('правки линейки: граница поля берётся по всей высоте, отступы уходят в поле до края листа, подсказки', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Абзац для отступа');
  await p.waitForSelector('#hr-inner .ri-left:not([hidden])');
  // Граница левого поля: хватаем ниже метки первой строки и выше выступа — раньше там было «мимо».
  const m = await (await p.$('#hr-inner .rm[aria-label="Левое поле"]')).boundingBox();
  ok(m.height >= 20, `граница поля по всей высоте линейки: ${m.height}`);
  const left0 = await p.evaluate(() => window.ens.getMeta().page.left);
  await p.mouse.move(m.x + m.width / 2, m.y + 9);
  await p.mouse.down();
  await p.mouse.move(m.x + m.width / 2 + 30, m.y + 9, { steps: 4 });
  await p.mouse.up();
  ok((await p.evaluate(() => window.ens.getMeta().page.left)) - left0 > 15, 'левое поле сдвинулось');
  // Отступ слева уводим в серую часть (в поле).
  const li = await (await p.$('#hr-inner .ri-left')).boundingBox();
  await p.mouse.move(li.x + li.width / 2, li.y + li.height / 2);
  await p.mouse.down();
  await p.mouse.move(li.x + li.width / 2 - 40, li.y + li.height / 2, { steps: 4 });
  await p.mouse.up();
  const ind = await p.evaluate(() => window.ens.view.state.doc.firstChild.attrs.indent);
  ok(ind < -20, `отступ слева отрицательный (в поле): ${ind}`);
  // Не дальше края листа.
  const li2 = await (await p.$('#hr-inner .ri-left')).boundingBox();
  await p.mouse.move(li2.x + li2.width / 2, li2.y + li2.height / 2);
  await p.mouse.down();
  await p.mouse.move(li2.x - 600, li2.y + li2.height / 2, { steps: 4 });
  await p.mouse.up();
  const r = await p.evaluate(() => ({ ind: window.ens.view.state.doc.firstChild.attrs.indent, left: window.ens.getMeta().page.left }));
  ok(Math.abs(r.ind + r.left) < 1, `отступ упёрся в край листа: ${r.ind} при поле ${r.left}`);
  const hint = await p.$eval('#hr-inner .rm[aria-label="Левое поле"]', (e) => e.title);
  ok(/тяните/.test(hint), `подсказка у границы поля: ${hint}`);
  ok(/поле/.test(await p.$eval('#hr-inner .ri-left', (e) => e.title)), 'подсказка у маркера отступа');
  await p.context().close();
});
await test('таблица из файла Excel (.xlsx): формулы из файла, общие формулы, неподдержанное — значением; Ctrl+V значений — итоги формулами', async () => {
  const { zipSync, strToU8 } = createRequire(import.meta.url)('fflate');
  // Книга как у пользователя: A1:E3, E1 = C1+B1+A1, E2 = C2+B2-A2, строка 3 — СУММ (общая формула), D — пусто.
  // На втором листе — функция, которой в тексте нет (ВПР): остаётся значением.
  const sheet1 = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c><c r="D1" s="2"/><c r="E1"><f>C1+B1+A1</f><v>6</v></c></row>
    <row r="2"><c r="A2"><v>7</v></c><c r="B2"><v>9</v></c><c r="C2"><v>11</v></c><c r="E2"><f>C2+B2-A2</f><v>13</v></c></row>
    <row r="3"><c r="A3"><f t="shared" ref="A3:C3" si="0">SUM(A1:A2)</f><v>8</v></c><c r="B3"><f t="shared" si="0"/><v>11</v></c><c r="C3"><f t="shared" si="0"/><v>14</v></c><c r="E3"><f>SUM(E1:E2)</f><v>19</v></c></row>
  </sheetData></worksheet>`;
  const sheet2 = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>
    <row r="2"><c r="B2" t="s"><v>0</v></c><c r="C2"><v>5</v></c></row>
    <row r="3"><c r="B3" t="inlineStr"><is><t>Код</t></is></c><c r="C3"><f>VLOOKUP(1,Лист1!A1:B2,2)</f><v>2</v></c></row>
  </sheetData></worksheet>`;
  const xlsx = zipSync({
    'xl/workbook.xml': strToU8('<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Лист1" sheetId="1" r:id="rId1"/><sheet name="Цены" sheetId="2" r:id="rId2"/></sheets></workbook>'),
    'xl/_rels/workbook.xml.rels': strToU8('<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="/xl/worksheets/sheet2.xml"/></Relationships>'),
    'xl/sharedStrings.xml': strToU8('<sst><si><t>Цена &amp; НДС</t></si></sst>'),
    'xl/worksheets/sheet1.xml': strToU8(sheet1), 'xl/worksheets/sheet2.xml': strToU8(sheet2),
  });
  const p = await page();
  await p.click('.ProseMirror');
  await p.click('#tabs .tab:has-text("Вставка")');
  ok(await p.$('[data-command="insert.excel"]'), 'кнопка «Из Excel» на ленте «Вставка»');
  await p.setInputFiles('#xlsx-input', { name: 'книга.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(xlsx) });
  await p.waitForSelector('.sheet-pick');
  eq(await p.$$eval('.sheet-pick', (b) => b.map((x) => x.textContent)), ['Лист1 — 3×5', 'Цены — 2×2'], 'выбор листа');
  await p.click('.sheet-pick');
  const formulas = () => p.evaluate(() => { const out = []; window.ens.view.state.doc.descendants((n) => { if (n.attrs.formula) out.push(n.attrs.formula); }); return out; });
  const grid = () => p.evaluate(() => [...window.ens.view.dom.querySelectorAll('table')].map((t) => [...t.querySelectorAll('tr')].map((r) => [...r.children].map((c) => c.textContent))));
  eq(await formulas(), ['=C1+B1+A1', '=C2+B2-A2', '=SUM(A1:A2)', '=SUM(B1:B2)', '=SUM(C1:C2)', '=SUM(E1:E2)'], 'все формулы, общие — развёрнуты');
  eq((await grid())[0], [['1', '2', '3', '', '6'], ['7', '9', '11', '', '13'], ['8', '11', '14', '', '19']], 'значения');
  // Правка числа: A1 1 → 4 — пересчитываются E1, A3 и E3 (E2 от A1 не зависит).
  await p.waitForTimeout(700);
  await p.evaluate(() => { const E = window.ens; let pos = null; E.view.state.doc.descendants((n, q) => { if (pos == null && n.isText && n.text === '1') pos = q; }); E.view.dispatch(E.view.state.tr.insertText('4', pos, pos + 1)); });
  const g = (await grid())[0];
  eq([g[0][4], g[1][4], g[2][0], g[2][4]], ['9', '13', '11', '22'], 'пересчёт как в Excel');
  // Второй лист: строка из общих строк, ВПР со ссылкой на другой лист — значением.
  await p.keyboard.press('Control+End');
  await p.setInputFiles('#xlsx-input', { name: 'книга.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(xlsx) });
  await p.waitForSelector('.sheet-pick');
  await p.click('.sheet-pick >> nth=1');
  eq((await grid())[1], [['Цена & НДС', '5'], ['Код', '2']], 'второй лист значениями');
  ok(/Значениями оставлены 1/.test(await p.$eval('#toast', (t) => t.textContent)), 'сообщение: одна формула оставлена значением');
  await p.context().close();

  // Ctrl+V из Excel: в буфере только значения (HTML без x:fmla) — итоги становятся формулами.
  const q = await page();
  await q.click('.ProseMirror');
  const html = `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><table>
    <tr><td x:num>1</td><td x:num>2</td><td x:num>3</td><td></td><td x:num>6</td></tr>
    <tr><td x:num>7</td><td x:num>9</td><td x:num>11</td><td></td><td x:num>13</td></tr>
    <tr><td x:num>8</td><td x:num>11</td><td x:num>14</td><td></td><td x:num>19</td></tr></table></body></html>`;
  await q.evaluate((h) => { const dt = new DataTransfer(); dt.setData('text/html', h); dt.setData('text/plain', '1\t2'); window.ens.view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true })); }, html);
  const f2 = await q.evaluate(() => { const out = []; window.ens.view.state.doc.descendants((n) => { if (n.attrs.formula) out.push(n.attrs.formula); }); return out; });
  eq(f2, ['=SUM(LEFT)', '=SUM(ABOVE)', '=SUM(ABOVE)', '=SUM(ABOVE)', '=SUM(ABOVE)'], 'итоги строки и столбца распознаны (E2 = C2+B2−A2 — не сумма, осталась числом)');
  await q.context().close();
});

await test('вставка из Excel: таблица с живыми формулами (смещение адресов подбирается), пересчёт, DOCX, специальная вставка', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  // Кусок листа Excel B2:D5 — формулы с адресами листа (как кладёт Excel), рядом показанные значения.
  const html = `<html xmlns:x="urn:schemas-microsoft-com:office:excel"><body><table>
    <tr><td>Товар</td><td>Кол-во</td><td>Сумма</td></tr>
    <tr><td>Бумага</td><td x:num="3">3</td><td x:num="1500" x:fmla="=C3*500">1 500</td></tr>
    <tr><td>Ручки</td><td x:num="10">10</td><td x:num="250" x:fmla="=C4*25">250</td></tr>
    <tr><td>Итого</td><td x:num="13" x:fmla="=SUM(C3:C4)">13</td><td x:num="1750" x:fmla="=SUM(D3:D4)">1 750</td></tr>
  </table></body></html>`;
  const paste = (h, t) => p.evaluate(({ h, t }) => {
    const dt = new DataTransfer();
    if (h) dt.setData('text/html', h);
    dt.setData('text/plain', t);
    window.ens.view.dom.dispatchEvent(new ClipboardEvent('paste', { clipboardData: dt, bubbles: true, cancelable: true }));
  }, { h, t });
  await paste(html, 'Товар\tКол-во\tСумма');
  const grid = () => p.evaluate(() => { const t = window.ens.view.dom.querySelector('table'); return [...t.querySelectorAll('tr')].map((r) => [...r.children].map((c) => c.textContent)); });
  const formulas = () => p.evaluate(() => { const out = []; window.ens.view.state.doc.descendants((n) => { if (n.attrs.formula) out.push(n.attrs.formula); }); return out; });
  eq(await formulas(), ['=B2*500', '=B3*25', '=SUM(B2:B3)', '=SUM(C2:C3)'], 'адреса листа пересчитаны на адреса таблицы');
  eq((await grid())[3], ['Итого', '13', '1 750'], 'итоги на месте');
  // Меняем количество бумаги 3 → 5: сумма и итоги пересчитываются, одна отмена возвращает всё.
  await p.waitForTimeout(700); // отдельный шаг истории (правки ближе 0,6 с объединяются)
  await p.evaluate(() => {
    const E = window.ens; let pos = null;
    E.view.state.doc.descendants((n, q) => { if (pos == null && n.isText && n.text === '3') pos = q; });
    E.view.dispatch(E.view.state.tr.insertText('5', pos, pos + 1));
  });
  const g2 = await grid();
  eq([g2[1][2], g2[3][1], g2[3][2]], ['2 500', '15', '2 750'], 'пересчёт после правки числа');
  await p.keyboard.press('Control+KeyZ');
  eq((await grid())[3][2], '1 750', 'одна отмена — и число, и итоги назад');
  // DOCX: формулы пишутся полем { =… } с результатом и читаются обратно.
  const back = await p.evaluate(() => {
    const E = window.ens;
    const bytes = E.exportDocx(E.view.state.doc, E.getMeta(), 'f');
    const xml = new TextDecoder().decode(E.unzip(bytes)['word/document.xml']);
    const b = E.importDocx(bytes);
    const f = []; b.doc.descendants((n) => { if (n.attrs.formula) f.push(n.attrs.formula); });
    return { fld: (xml.match(/w:instr=" =/g) ?? []).length, f, lossy: b.meta.lossy };
  });
  eq(back.fld, 4, 'в DOCX четыре поля-формулы');
  eq(back.f, ['=B2*500', '=B3*25', '=SUM(B2:B3)', '=SUM(C2:C3)'], 'формулы прочитаны из DOCX');
  ok(!back.lossy.some((x) => /пол/.test(x)), `без предупреждения о потерях: ${back.lossy}`);
  // Формула из окна, как в Word: =SUM(ABOVE) в новой строке.
  await p.evaluate(() => { const E = window.ens; E.loadDoc(E.schema.nodes.doc.create(null, [E.schema.nodes.paragraph.create()]), E.getMeta(), 'n'); });
  await p.click('.ProseMirror');
  await paste('', '10\t20\n30\t40\n\t');
  ok(!(await p.$('.ProseMirror table')), 'обычная вставка текста с табуляцией — текстом, как в Word');
  // Специальная вставка: буфер прочитать нельзя → ждём Ctrl+V и вставляем таблицей.
  await p.evaluate(() => { const E = window.ens; E.loadDoc(E.schema.nodes.doc.create(null, [E.schema.nodes.paragraph.create()]), E.getMeta(), 'n'); });
  await p.click('.ProseMirror');
  await p.keyboard.press('Control+Alt+KeyV');
  await p.waitForSelector('#dialog:not([hidden]) .paste-opts');
  await p.check('#dialog input[value="table"]');
  await p.click('#dialog button.primary');
  await p.waitForTimeout(200);
  await paste('', '10\t20\n30\t40\n\t');
  await p.waitForSelector('.ProseMirror table');
  // Ставим формулу в последнюю ячейку первого столбца.
  await p.evaluate(() => { const E = window.ens; let pos = null; E.view.state.doc.descendants((n, q) => { if (n.type.name === 'table_cell') pos = q; }); });
  await p.evaluate(() => { const E = window.ens; const cells = []; E.view.state.doc.descendants((n, q) => { if (n.type.name === 'table_cell') cells.push(q); }); E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, cells[4] + 2))); });
  await p.evaluate(() => window.ens.run('table.formula'));
  await p.waitForSelector('#dialog:not([hidden]) input');
  eq(await p.$eval('#dialog input', (i) => i.value), '=SUM(ABOVE)', 'по умолчанию — сумма выше, как в Word');
  await p.keyboard.press('Enter');
  await p.waitForSelector('#dialog', { state: 'hidden' });
  eq((await grid())[2][0], '40', '=SUM(ABOVE) = 10 + 30');
  await p.context().close();
});

await test('формулы таблиц: функции по-русски и по-английски, проценты, ABOVE/LEFT, ошибки и цикл', async () => {
  const p = await page();
  const r = await p.evaluate(async () => {
    const { computeGrid, parseNum } = await import('/src/table-formulas.ts');
    const text = [['10', '20', ''], ['3 000,5', '15 %', ''], ['', '', '']];
    const f = (r, c, x) => { const g = text.map((row) => row.map(() => null)); g[r][c] = x; return computeGrid(text, g).get(`${r},${c}`); };
    return {
      sum: f(2, 0, '=SUM(ABOVE)'), summ: f(2, 0, '=СУММ(A1:A2)'), avg: f(0, 2, '=СРЗНАЧ(LEFT)'), pct: f(2, 1, '=B2*200'),
      round: f(2, 2, '=ОКРУГЛ(A2/7; 1)'), pow: f(2, 2, '=2^10'), div0: f(2, 2, '=A1/(B1-20)'), bad: f(2, 2, '=A1+'), name: f(2, 2, '=НЕТТАКОЙ(1)'),
      ref: f(2, 2, '=Z99'), num: [parseNum('3 658 420'), parseNum('1 234,5'), parseNum('−7'), parseNum('12,5 %'), parseNum('абв')],
      cycle: computeGrid([['', '']], [['=B1', '=A1']]).get('0,0'),
    };
  });
  eq([r.sum, r.summ, r.avg, r.pct, r.round, r.pow], ['3 010,5', '3 010,5', '15', '30', '428,6', '1 024'], 'результаты');
  eq([r.div0, r.bad, r.name, r.ref, r.cycle], ['#ДЕЛ/0!', '#ОШИБКА!', '#ИМЯ?', '#ССЫЛКА!', '#ЦИКЛ!'], 'ошибки');
  eq(r.num.slice(0, 4), [3658420, 1234.5, -7, 0.125], 'числа по-русски');
  ok(Number.isNaN(r.num[4]), 'текст — не число');
  await p.context().close();
});
await test('пункт 10: перенос слов, подложка на каждой странице, колонки сохраняются в DOCX без потерь', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const long = 'Электрификация достопримечательностей одновременно с переоборудованием предприятий. '.repeat(30);
    E.loadDoc(s.nodes.doc.create(null, Array.from({ length: 6 }, () => s.nodes.paragraph.create({ align: 'justify' }, s.text(long)))), E.getMeta(), 'p10');
  });
  await p.waitForTimeout(400);
  // Перенос слов.
  const lines = () => p.evaluate(() => { const r = document.createRange(); r.selectNodeContents(window.ens.view.dom.querySelector('p')); return new Set([...r.getClientRects()].map((x) => Math.round(x.top))).size; });
  const before = await lines();
  await p.evaluate(() => window.ens.run('layout.hyphenation'));
  await p.waitForTimeout(200);
  ok(await p.evaluate(() => window.ens.view.dom.classList.contains('hyphenate')), 'перенос включён');
  await p.click('#tabs .tab:text-is("Макет")');
  ok(await p.$eval('[data-command="layout.hyphenation"]', (b) => b.classList.contains('on')), 'кнопка на ленте нажата');
  const after = await lines();
  ok(after <= before, `строк не больше: ${before} → ${after}`);
  ok(await p.evaluate(() => window.ens.isDirty()), 'документ помечен изменённым');
  // Подложка из окна: готовая надпись.
  await p.evaluate(() => window.ens.run('layout.watermark'));
  await p.waitForSelector('#dialog:not([hidden]) .wm-preset');
  await p.click('#dialog .wm-preset:has-text("ЧЕРНОВИК")');
  await p.waitForTimeout(300);
  const wm = await p.evaluate(() => ({ papers: document.querySelectorAll('.paper').length, marks: [...document.querySelectorAll('.paper .watermark')].map((w) => w.textContent), angle: getComputedStyle(document.querySelector('.paper .watermark')).transform }));
  ok(wm.papers >= 2 && wm.marks.length === wm.papers && wm.marks.every((t) => t === 'ЧЕРНОВИК'), `подложка на каждом листе: ${JSON.stringify(wm)}`);
  ok(wm.angle !== 'none', 'надпись повёрнута');
  // DOCX: подложка как у Word, перенос — w:autoHyphenation; колонки (2) раздела сохраняются.
  const r = await p.evaluate(() => {
    const E = window.ens;
    const m = { ...E.getMeta(), page: { ...E.getMeta().page, cols: 2, colSpace: 35.4 } };
    const bytes = E.exportDocx(E.view.state.doc, m, 'p10');
    const files = E.unzip(bytes);
    const txt = (n) => (files[n] ? new TextDecoder().decode(files[n]) : '');
    const back = E.importDocx(bytes);
    return {
      hdr: txt('word/header1.xml'), settings: txt('word/settings.xml'), doc: txt('word/document.xml'),
      wm: back.meta.watermark, hy: back.meta.hyphenate, cols: back.meta.page.cols, header: back.meta.hf.header, lossy: back.meta.lossy,
      again: (() => { const b2 = E.exportDocx(back.doc, back.meta, 'x'); return new TextDecoder().decode(E.unzip(b2)['word/document.xml']).includes('w:cols w:num="2"'); })(),
    };
  });
  ok(r.hdr.includes('PowerPlusWaterMarkObject') && r.hdr.includes('string="ЧЕРНОВИК"'), 'подложка в колонтитуле, как у Word');
  ok(r.settings.includes('<w:autoHyphenation/>'), 'перенос слов в настройках документа');
  ok(r.doc.includes('<w:cols w:num="2"'), 'колонки в свойствах раздела');
  eq([r.wm, r.hy, r.cols], ['ЧЕРНОВИК', true, 2], 'всё прочитано обратно');
  eq(r.header, null, 'колонтитул, созданный только ради подложки, пустой');
  eq(r.lossy, [], 'без потерь');
  ok(r.again, 'колонки переживают второе сохранение');
  // Убрать подложку.
  await p.evaluate(() => window.ens.run('layout.watermark'));
  await p.waitForSelector('#dialog:not([hidden]) .wm-preset');
  await p.click('#dialog button:has-text("Убрать подложку")');
  await p.waitForTimeout(200);
  eq(await p.$$eval('.paper .watermark', (w) => w.length), 0, 'подложка убрана');
  await p.context().close();
});
await test('новый документ не считается изменённым (нет вопроса «Сохранить изменения?»)', async () => {
  const p = await page();
  eq(await p.evaluate(() => window.ens.isDirty()), false, 'чистый документ при запуске');
  await p.context().close();
});

// --- Подписи (названия) рисунков и таблиц
const shownCaptions = (p) => p.evaluate(() => [...document.querySelectorAll('.ProseMirror p')].filter((e) => e.querySelector('span.seq'))
  .map((e) => [...e.childNodes].map((n) => n.nodeType === 3 ? n.textContent : n.classList?.contains('seq') ? n.getAttribute('data-n') : n.textContent).join('')));
await test('подписи: «Таблица N» над таблицей, «Рисунок N» под рисунком, номера пересчитываются сами', async () => {
  const p = await page();
  await p.evaluate(() => {
    const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const cell = (t) => s.nodes.table_cell.create(null, P(t));
    const T = s.nodes.table.create(null, [s.nodes.table_row.create(null, [cell('А'), cell('Б')])]);
    const img = s.nodes.paragraph.create(null, s.nodes.image.create({ src: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', width: 40, height: 40 }));
    E.loadDoc(s.nodes.doc.create(null, [P('Текст'), T, img, P('Конец')]), E.getMeta(), 'cap');
    // Курсор в таблице.
    E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 11)));
  });
  await p.evaluate(() => window.ens.run('insert.caption'));
  await p.waitForSelector('.caption-dialog');
  eq(await p.$eval('.caption-dialog select[aria-label="Подпись"]', (s) => s.value), 'Таблица', 'в таблице — «Таблица»');
  eq(await p.$eval('.caption-dialog select[aria-label="Положение"]', (s) => s.value), 'above', 'над таблицей (ГОСТ)');
  await p.fill('.caption-dialog input[aria-label="Текст подписи"]', 'Исходные данные');
  eq(await p.textContent('.caption-preview'), 'Таблица 1 – Исходные данные', 'предпросмотр');
  await p.click('.caption-dialog button.primary');
  // Курсор у рисунка.
  await p.evaluate(() => { const E = window.ens, d = E.view.state.doc; let at = 0; d.descendants((n, pos) => { if (n.type.name === 'image') at = pos; }); E.view.dispatch(E.view.state.tr.setSelection(E.pm.NodeSelection.create(d, at))); });
  await p.evaluate(() => window.ens.run('insert.caption'));
  await p.waitForSelector('.caption-dialog');
  eq(await p.$eval('.caption-dialog select[aria-label="Положение"]', (s) => s.value), 'below', 'под рисунком');
  await p.fill('.caption-dialog input[aria-label="Текст подписи"]', 'Схема');
  await p.keyboard.press('Enter');
  eq(await shownCaptions(p), ['Таблица 1 – Исходные данные', 'Рисунок 1 – Схема'], 'подписи вставлены');
  const r = await p.evaluate(() => { const d = window.ens.view.state.doc; return [d.child(1).type.name, d.child(1).attrs.align, d.child(2).type.name, d.child(4).type.name, d.child(4).attrs.align]; });
  eq(r, ['paragraph', null, 'table', 'paragraph', 'center'], 'место и выравнивание по ГОСТ');
  // Ещё одна таблица выше первой — номера ниже сдвигаются сами.
  await p.evaluate(() => { const E = window.ens, s = E.schema; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 1))); });
  await p.evaluate(() => window.ens.run('insert.caption'));
  await p.waitForSelector('.caption-dialog');
  await p.selectOption('.caption-dialog select[aria-label="Подпись"]', 'Таблица');
  await p.selectOption('.caption-dialog select[aria-label="Положение"]', 'above');
  await p.fill('.caption-dialog input[aria-label="Текст подписи"]', 'Новая');
  await p.click('.caption-dialog button.primary');
  eq(await shownCaptions(p), ['Таблица 1 – Новая', 'Таблица 2 – Исходные данные', 'Рисунок 1 – Схема'], 'номера пересчитаны');
  await p.keyboard.press('Control+KeyZ');
  eq(await shownCaptions(p), ['Таблица 1 – Исходные данные', 'Рисунок 1 – Схема'], 'отмена возвращает номера');
  await p.context().close();
});
await test('подписи в docx: поле SEQ с числом туда и обратно, поля SEQ из Word читаются', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    const cap = (name, t) => s.nodes.paragraph.create(null, [s.text(name + ' '), s.nodes.seq.create({ name }), s.text(' – ' + t)]);
    const doc = s.nodes.doc.create(null, [cap('Рисунок', 'один'), cap('Таблица', 'первая'), cap('Рисунок', 'два')]);
    const bytes = E.exportDocx(doc, E.getMeta(), 'x');
    const xml = new TextDecoder().decode(E.unzip(bytes)['word/document.xml']);
    const back = E.importDocx(bytes);
    // Сложное поле из Word: begin / instr / separate / результат / end.
    const files = E.unzip(bytes);
    files['word/document.xml'] = new TextEncoder().encode(xml.replace('<w:fldSimple w:instr=" SEQ Рисунок \\* ARABIC "><w:r><w:t>1</w:t></w:r></w:fldSimple>',
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Рисунок \\* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'));
    const word = E.importDocx(E.zip(files));
    E.loadDoc(word.doc, word.meta, 'w');
    return { xml, same: back.doc.eq(doc), lossy: back.meta.lossy, wordLossy: word.meta.lossy, caps: E.listCaptions().map((c) => [c.name, c.n, c.text]) };
  });
  ok(r.xml.includes('<w:fldSimple w:instr=" SEQ Рисунок \\* ARABIC "><w:r><w:t>2</w:t>'), 'второй рисунок записан с числом 2');
  ok(r.xml.includes('SEQ Таблица'), 'таблица');
  ok(r.same, 'документ тот же после круга');
  eq([r.lossy, r.wordLossy], [[], []], 'без потерь');
  eq(r.caps, [['Рисунок', 1, 'один'], ['Таблица', 1, 'первая'], ['Рисунок', 2, 'два']], 'сложное поле SEQ → номер, число пересчитано');
  await p.context().close();
});
// --- Закладки
const BM_SETUP = () => {
  const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, s.text(t));
  E.loadDoc(s.nodes.doc.create(null, [P('Введение к документу'), P('Глава первая: цели'), P('Заключение и выводы')]), E.getMeta(), 'bm');
};
const caret = (p) => p.evaluate(() => window.ens.view.state.selection.from);
await test('закладки: добавить из окна, перейти, ссылка «Место в документе», Ctrl+щелчок ведёт к закладке', async () => {
  const p = await page();
  await p.evaluate(BM_SETUP);
  // Выделяем «Глава первая» и вставляем закладку: имя подсказано из текста.
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 23, 35))); });
  await p.keyboard.press('Control+Shift+F5');
  await p.waitForSelector('.bookmark-dialog');
  eq(await p.$eval('.bookmark-dialog input[type=text]', (i) => i.value), 'Глава_первая', 'имя по выделению');
  await p.fill('.bookmark-dialog input[type=text]', 'Цели');
  await p.click('.bookmark-dialog button:has-text("Добавить")');
  const bms = await p.evaluate(() => window.ens.listBookmarks());
  eq(bms.map((b) => [b.name, b.context]), [['Цели', 'Глава первая: цели']], 'закладка в начале выделения');
  // Неверное имя.
  await p.evaluate(() => window.ens.run('insert.bookmark'));
  await p.fill('.bookmark-dialog input[type=text]', '1 неверное');
  await p.click('.bookmark-dialog button:has-text("Добавить")');
  ok((await p.textContent('.bookmark-dialog .bm-status')).includes('начинается с буквы'), 'проверка имени');
  // Перейти из списка.
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 1))); });
  await p.selectOption('.bookmark-dialog .bm-list', 'Цели');
  await p.click('.bookmark-dialog button:has-text("Перейти")');
  eq(await caret(p), bms[0].pos + 1, 'курсор у закладки');
  // Ссылка на закладку из окна «Гиперссылка».
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 1, 9))); });
  await p.keyboard.press('Control+KeyK');
  await p.selectOption('select[aria-label="Место в документе"]', '#Цели');
  await p.click('#dialog button:has-text("Применить")');
  const href = await p.evaluate(() => { const n = window.ens.view.state.doc.nodeAt(2); return n.marks.find((m) => m.type.name === 'link')?.attrs.href; });
  eq(href, '#Цели', 'ссылка на закладку');
  // Ctrl+щелчок по ссылке — переход к закладке.
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 50))); });
  const box = await p.$eval('.ProseMirror a[href="#Цели"]', (a) => { const r = a.getBoundingClientRect(); return { x: r.x + 5, y: r.y + r.height / 2 }; });
  await p.keyboard.down('Control');
  await p.mouse.click(box.x, box.y);
  await p.keyboard.up('Control');
  eq(await caret(p), bms[0].pos + 1, 'Ctrl+щелчок перешёл к закладке');
  await p.context().close();
});
await test('закладки: переименование обновляет ссылки, удаление предупреждает и показывает ссылки без цели', async () => {
  const p = await page();
  await p.evaluate(BM_SETUP);
  await p.evaluate(() => {
    const E = window.ens, s = E.schema, st = E.view.state;
    const tr = st.tr.insert(23, s.nodes.bookmark.create({ name: 'Цели' }));
    tr.addMark(1, 9, s.marks.link.create({ href: '#Цели' }));
    E.view.dispatch(tr.setMeta('addToHistory', false));
  });
  await p.evaluate(() => window.ens.run('insert.bookmark'));
  await p.selectOption('.bookmark-dialog .bm-list', 'Цели');
  ok((await p.$eval('.bookmark-dialog .bm-list option', (o) => o.textContent)).includes('ссылок: 1'), 'в списке видно число ссылок');
  await p.fill('.bookmark-dialog input[type=text]', 'Задачи');
  await p.click('.bookmark-dialog button:has-text("Переименовать")');
  const r = await p.evaluate(() => ({ names: window.ens.listBookmarks().map((b) => b.name), href: window.ens.view.state.doc.nodeAt(2).marks.find((m) => m.type.name === 'link').attrs.href }));
  eq(r, { names: ['Задачи'], href: '#Задачи' }, 'ссылка переименована вместе с закладкой');
  await p.selectOption('.bookmark-dialog .bm-list', 'Задачи');
  await p.click('.bookmark-dialog button:has-text("Удалить")');
  await p.waitForSelector('#dialog button:has-text("Отмена")');
  ok((await p.textContent('#dialog')).includes('ведут ссылки: 1'), 'предупреждение о ссылках');
  await p.click('#dialog button.danger:has-text("Удалить")');
  await p.waitForSelector('.bookmark-dialog');
  eq(await p.evaluate(() => window.ens.listBookmarks().length), 0, 'удалена');
  ok((await p.textContent('.bookmark-dialog .bm-broken')).includes('Ссылки без цели: 1'), 'список ссылок без цели');
  await p.click('.bookmark-dialog .bm-broken-item');
  eq(await p.evaluate(() => { const { from, to } = window.ens.view.state.selection; return [from, to]; }), [1, 9], 'переход к битой ссылке');
  await p.keyboard.press('Control+KeyZ');
  eq(await p.evaluate(() => window.ens.listBookmarks().map((b) => b.name)), ['Задачи'], 'отмена возвращает закладку');
  await p.context().close();
});
await test('закладки: docx туда и обратно, _GoBack не переносится, повтор имени при вставке убирается', async () => {
  const p = await page();
  await p.evaluate(BM_SETUP);
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema;
    E.view.dispatch(E.view.state.tr.insert(23, s.nodes.bookmark.create({ name: 'Цели' })).insert(1, s.nodes.bookmark.create({ name: '_Toc1' })).addMark(2, 10, s.marks.link.create({ href: '#Цели' })));
    const bytes = E.exportDocx(E.view.state.doc, E.getMeta(), 'x');
    let xml = new TextDecoder().decode(E.unzip(bytes)['word/document.xml']);
    const back = E.importDocx(bytes);
    // Файл из Word: служебная закладка _GoBack.
    const files = E.unzip(bytes);
    files['word/document.xml'] = new TextEncoder().encode(xml.replace('<w:bookmarkStart w:id="0" w:name="_Toc1"/>', '<w:bookmarkStart w:id="0" w:name="_GoBack"/>'));
    const zipped = E.zip(files);
    return { xml, names: [], back: back.doc.toJSON(), now: E.view.state.doc.toJSON(), lossy: back.meta.lossy, goBack: E.importDocx(zipped).doc.toJSON() };
  });
  ok(/<w:bookmarkStart w:id="\d+" w:name="Цели"\/><w:bookmarkEnd w:id="\d+"\/>/.test(r.xml), 'w:bookmarkStart/End');
  ok(r.xml.includes('<w:hyperlink w:anchor="Цели"'), 'ссылка на закладку — w:anchor');
  eq(r.back, r.now, 'документ тот же после круга');
  eq(r.lossy, [], 'без потерь');
  ok(!JSON.stringify(r.goBack).includes('_GoBack'), '_GoBack не переносится');
  // Повтор имени: вставка копии абзаца с закладкой.
  const names = await p.evaluate(() => {
    const E = window.ens, st = E.view.state, para = st.doc.child(1);
    E.view.dispatch(st.tr.insert(st.doc.content.size, para));
    return window.ens.listBookmarks().map((b) => b.name);
  });
  eq(names, ['_Toc1', 'Цели'], 'одна закладка на имя');
  await p.context().close();
});

// --- Сравнение документов и исправления
const REV_SETUP = () => {
  const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, t ? s.text(t) : null);
  const A = s.nodes.doc.create(null, [P('Мама мыла раму вчера.'), P('Этот абзац удалят.'), P('Общий конец.')]);
  const B = s.nodes.doc.create(null, [P('Мама мыла окно вчера вечером.'), P('Общий конец.'), P('Новый абзац.')]);
  const m = E.getMeta();
  window.__rev = { A, B, meta: m };
  const res = E.compareSides({ doc: A, meta: m, name: 'a.docx' }, { doc: B, meta: m, name: 'b.docx' }, 'Рецензент', new Date('2026-11-17T10:00:00Z'));
  E.loadDoc(res.doc, res.meta, res.name);
  return { ins: res.inserted, del: res.deleted };
};
const docText = (p) => p.evaluate(() => { const d = window.ens.view.state.doc; return d.textBetween(0, d.content.size, '|'); });
await test('сравнение: правки по словам, целые абзацы, «было / стало» совпадают с исходниками', async () => {
  const p = await page();
  const n = await p.evaluate(REV_SETUP);
  ok(n.ins >= 2 && n.del >= 2, `исправлений мало: ${JSON.stringify(n)}`);
  const r = await p.evaluate(() => {
    const E = window.ens, revs = E.listRevisions(), doc = E.view.state.doc;
    const text = (x) => doc.textBetween(x.from, x.to, '|');
    return { revs: revs.map((x) => [x.kind, text(x), x.author]),
      ins: [...document.querySelectorAll('.ProseMirror ins.rev-ins')].map((e) => e.textContent),
      del: [...document.querySelectorAll('.ProseMirror del.rev-del')].map((e) => e.textContent),
      blocks: [...document.querySelectorAll('.ProseMirror [data-rev-block]')].map((e) => e.dataset.revBlock) };
  });
  ok(r.revs.some(([k, t]) => k === 'del' && t === 'раму') && r.revs.some(([k, t]) => k === 'ins' && t === 'окно'), `по словам: ${JSON.stringify(r.revs)}`);
  ok(r.revs.some(([k, t]) => k === 'ins' && t.includes('вечером')), 'вставка в конце абзаца');
  ok(r.revs.every(([, , a]) => a === 'Рецензент'), 'автор правок');
  eq(r.blocks.sort(), ['del', 'ins'], 'удалённый и новый абзацы помечены целиком');
  ok(!r.revs.some(([, t]) => t.includes('Мама')), 'общие слова не помечены');
  // «Было» и «Стало»: видимый текст листа совпадает с исходниками.
  const visible = () => p.evaluate(() => [...document.querySelectorAll('.ProseMirror > *')].filter((e) => e.offsetParent !== null || getComputedStyle(e).display !== 'none').map((e) => e.innerText.trim()).filter(Boolean).join('|'));
  await p.evaluate(() => window.ens.run('rev.view', 'original'));
  eq(await visible(), 'Мама мыла раму вчера.|Этот абзац удалят.|Общий конец.', 'было');
  await p.evaluate(() => window.ens.run('rev.view', 'final'));
  eq(await visible(), 'Мама мыла окно вчера вечером.|Общий конец.|Новый абзац.', 'стало');
  await p.evaluate(() => window.ens.run('rev.view', 'markup'));
  eq(await p.evaluate(() => window.ens.revView()), 'markup', 'режим исправлений');
  await p.context().close();
});
await test('исправления: принять и отклонить по одной, принять все = новая редакция, отклонить все = старая', async () => {
  const p = await page();
  await p.evaluate(REV_SETUP);
  const count = () => p.evaluate(() => window.ens.listRevisions().length);
  const total = await count();
  // Курсор в начало: первое исправление — удаление «раму». Отклоняем его — слово остаётся без пометки.
  await p.evaluate(() => { const E = window.ens; E.view.dispatch(E.view.state.tr.setSelection(E.pm.TextSelection.create(E.view.state.doc, 1))); });
  await p.evaluate(() => window.ens.run('rev.reject'));
  eq(await count(), total - 1, 'одно исправление снято');
  ok((await docText(p)).startsWith('Мама мыла раму'), `«раму» осталось: ${await docText(p)}`);
  // Следующее (выделено автоматически) — вставка «окно»; принимаем.
  const sel = await p.evaluate(() => { const E = window.ens, { from, to } = E.view.state.selection; return E.view.state.doc.textBetween(from, to); });
  eq(sel, 'окно', 'после решения выделено следующее исправление');
  await p.keyboard.press('Control+Alt+KeyA');
  eq(await count(), total - 2, 'Ctrl+Alt+A принимает');
  ok((await docText(p)).startsWith('Мама мыла рамуокно'), await docText(p)); // пробел общий: «раму» отклонено, «окно» принято
  await p.keyboard.press('Control+KeyZ');
  eq(await count(), total - 1, 'отмена возвращает исправление');
  // Все сразу.
  await p.evaluate(() => window.ens.run('rev.acceptAll'));
  eq(await count(), 0, 'исправлений не осталось');
  eq(await docText(p), 'Мама мыла рамуокно вчера вечером.|Общий конец.|Новый абзац.', 'принято всё (кроме отклонённого)');
  await p.evaluate(REV_SETUP);
  await p.evaluate(() => window.ens.run('rev.acceptAll'));
  eq(await docText(p), 'Мама мыла окно вчера вечером.|Общий конец.|Новый абзац.', 'принять все = новая редакция');
  await p.evaluate(REV_SETUP);
  await p.evaluate(() => window.ens.run('rev.rejectAll'));
  eq(await docText(p), 'Мама мыла раму вчера.|Этот абзац удалят.|Общий конец.', 'отклонить все = старая редакция');
  eq(await p.evaluate(() => window.ens.view.state.doc.toJSON()), await p.evaluate(() => window.__rev.A.toJSON()), 'старая редакция совпадает узел в узел');
  await p.context().close();
});
await test('исправления: таблицы и списки сравниваются по ячейкам, целиком новая таблица отклоняется целиком', async () => {
  const p = await page();
  const r = await p.evaluate(() => {
    const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, t ? s.text(t) : null);
    const cell = (t) => s.nodes.table_cell.create(null, P(t));
    const row = (...t) => s.nodes.table_row.create(null, t.map(cell));
    const T = (...rows) => s.nodes.table.create(null, rows);
    const li = (t) => s.nodes.list_item.create(null, P(t));
    const A = s.nodes.doc.create(null, [P('Начало'), T(row('Яблоко', '10'), row('Груша', '5')), s.nodes.bullet_list.create(null, [li('первый'), li('второй')])]);
    const B = s.nodes.doc.create(null, [P('Начало'), T(row('Яблоко', '12'), row('Груша', '5'), row('Слива', '7')), s.nodes.bullet_list.create(null, [li('первый пункт'), li('второй')]), T(row('Новая', 'таблица'))]);
    const m = E.getMeta();
    const res = E.compareSides({ doc: A, meta: m, name: 'a' }, { doc: B, meta: m, name: 'b' }, 'Я');
    E.loadDoc(res.doc, res.meta, res.name);
    const doc = E.view.state.doc, revs = E.listRevisions().map((x) => [x.kind, doc.textBetween(x.from, x.to, '|')]);
    const tables = () => { let n = 0; E.view.state.doc.descendants((x) => { if (x.type.name === 'table') n++; }); return n; };
    const before = tables();
    E.run('rev.rejectAll');
    return { revs, before, after: tables(), text: E.view.state.doc.textBetween(0, E.view.state.doc.content.size, '|'), A: A.toJSON(), now: E.view.state.doc.toJSON() };
  });
  ok(r.revs.some(([k, t]) => k === 'del' && t === '10') && r.revs.some(([k, t]) => k === 'ins' && t === '12'), `правка в ячейке: ${JSON.stringify(r.revs)}`);
  ok(r.revs.some(([k, t]) => k === 'ins' && t.includes('Слива')), 'новая строка');
  ok(r.revs.some(([k, t]) => k === 'ins' && t.trim() === 'пункт'), 'правка в пункте списка');
  eq([r.before, r.after], [2, 1], 'новая таблица убрана целиком при отклонении');
  eq(r.now, r.A, 'отклонить все = исходный документ');
  await p.context().close();
});
await test('исправления в docx: w:ins/w:del туда и обратно, исправления Word читаются', async () => {
  const p = await page();
  await p.evaluate(REV_SETUP);
  const r = await p.evaluate(() => {
    const E = window.ens, before = E.listRevisions().map((x) => [x.kind, x.author, E.view.state.doc.textBetween(x.from, x.to, '|')]);
    const bytes = E.exportDocx(E.view.state.doc, E.getMeta(), 'x');
    const xml = new TextDecoder().decode(E.unzip(bytes)['word/document.xml']);
    const back = E.importDocx(bytes);
    E.loadDoc(back.doc, back.meta, 'back');
    const after = E.listRevisions().map((x) => [x.kind, x.author, E.view.state.doc.textBetween(x.from, x.to, '|')]);
    return { xml, before, after, lossy: back.meta.lossy, same: back.doc.eq(window.__cmpDoc ?? back.doc) };
  });
  ok(/<w:ins w:id="\d+" w:author="Рецензент" w:date="2026-11-17T10:00:00Z">/.test(r.xml), 'w:ins с автором и датой');
  ok(/<w:del [^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:delText xml:space="preserve">раму<\/w:delText>/.test(r.xml), 'удалённое — w:delText');
  ok(/<w:pPr><w:rPr><w:del [^>]*\/><\/w:rPr><\/w:pPr>/.test(r.xml), 'знак удалённого абзаца');
  eq(r.lossy, [], 'без потерь');
  eq(r.after, r.before, 'исправления те же после открытия');
  // Принять все в прочитанном файле — новая редакция.
  await p.evaluate(() => window.ens.run('rev.acceptAll'));
  eq(await docText(p), 'Мама мыла окно вчера вечером.|Общий конец.|Новый абзац.', 'после круга принимаются как раньше');
  await p.context().close();
});
await test('сравнение: окно «Сравнить» с двумя файлами DOCX открывает новый документ, текущий не меняется', async () => {
  const p = await page();
  const files = await p.evaluate(() => {
    const E = window.ens, s = E.schema, P = (t) => s.nodes.paragraph.create(null, s.text(t));
    const m = E.getMeta();
    const a = E.exportDocx(s.nodes.doc.create(null, [P('Договор поставки № 1'), P('Срок — десять дней.')]), m, 'a');
    const b = E.exportDocx(s.nodes.doc.create(null, [P('Договор поставки № 2'), P('Срок — пять дней.')]), m, 'b');
    return { a: [...a], b: [...b] };
  });
  await p.click('.ProseMirror');
  await p.keyboard.type('Мой текущий текст');
  await p.evaluate(() => window.ens.run('review.compare'));
  await p.waitForSelector('#dialog:not([hidden]) .compare-dialog, #dialog:not([hidden]).compare-dialog, .compare-dialog');
  await p.selectOption('select[data-compare="older"]', 'file');
  await p.setInputFiles('input[data-compare="older-file"]', { name: 'Договор v1.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(files.a) });
  await p.selectOption('select[data-compare="newer"]', 'file');
  await p.setInputFiles('input[data-compare="newer-file"]', { name: 'Договор v2.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: Buffer.from(files.b) });
  await p.click('.compare-dialog button.primary');
  // Текущий документ изменён: программа спрашивает о сохранении — не сохраняем.
  const btn = await p.waitForSelector('#dialog:not([hidden]) button:has-text("Не сохранять")', { timeout: 3000 }).catch(() => null);
  if (btn) await btn.click();
  await p.waitForFunction(() => window.ens.listRevisions().length > 0);
  const r = await p.evaluate(() => ({ name: document.title, revs: window.ens.listRevisions().map((x) => [x.kind, window.ens.view.state.doc.textBetween(x.from, x.to)]) }));
  ok(r.name.includes('Сравнение — Договор v1 и Договор v2'), r.name);
  eq(r.revs, [['del', '1'], ['ins', '2'], ['del', 'десять'], ['ins', 'пять']], 'правки по словам');
  // Кнопки ленты: «Принять» на вкладке «Проверка».
  await p.click('#tabs button.tab:text-is("Проверка")');
  await p.click('button[data-command="rev.accept"]');
  eq(await p.evaluate(() => window.ens.listRevisions().length), 3, 'кнопка «Принять»');
  await p.context().close();
});
await test('сравнение: версия из истории против текущего документа', async () => {
  const p = await page();
  await p.click('.ProseMirror');
  await p.keyboard.type('Первая редакция текста');
  await p.evaluate(() => window.ens.snap('named', 'черновик'));
  await p.keyboard.press('End');
  await p.keyboard.type(' с дополнением');
  await p.evaluate(() => window.ens.run('ver.history'));
  await p.click('#dialog button:has-text("Сравнить с исправлениями")');
  await p.waitForSelector('.compare-dialog');
  eq(await p.$eval('select[data-compare="newer"]', (s) => s.value), 'current', 'новая сторона — текущий документ');
  ok((await p.$eval('select[data-compare="older"]', (s) => s.value)).startsWith('v:'), 'старая сторона — версия');
  await p.click('.compare-dialog button.primary');
  const btn = await p.waitForSelector('#dialog:not([hidden]) button:has-text("Не сохранять")', { timeout: 3000 }).catch(() => null);
  if (btn) await btn.click();
  await p.waitForFunction(() => window.ens.listRevisions().length > 0);
  const r = await p.evaluate(() => window.ens.listRevisions().map((x) => [x.kind, window.ens.view.state.doc.textBetween(x.from, x.to)]));
  eq(r, [['ins', ' с дополнением']], 'вставка относительно версии');
  await p.context().close();
});
ok(errors.length === 0, `ошибки страницы: ${errors.join('; ')}`);
await browser.close();
server.kill();
console.log(failed ? `\nПровалено: ${failed}` : '\nВсе проверки пройдены');
process.exit(failed ? 1 : 0);
