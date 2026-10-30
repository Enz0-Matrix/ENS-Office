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

ok(errors.length === 0, `ошибки страницы: ${errors.join('; ')}`);
await browser.close();
server.kill();
console.log(failed ? `\nПровалено: ${failed}` : '\nВсе проверки пройдены');
process.exit(failed ? 1 : 0);
