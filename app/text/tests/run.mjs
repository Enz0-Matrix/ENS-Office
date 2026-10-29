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

ok(errors.length === 0, `ошибки страницы: ${errors.join('; ')}`);
await browser.close();
server.kill();
console.log(failed ? `\nПровалено: ${failed}` : '\nВсе проверки пройдены');
process.exit(failed ? 1 : 0);
