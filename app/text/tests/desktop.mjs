// Проверка настольной версии ENS-Tekst на помощнике для РЕД ОС 7 (app/text/redos7): настоящий помощник,
// настоящий интерфейс в Chromium. Окно браузера и окно выбора файла (zenity) подменяются сценариями.
// Оболочку Tauri здесь не собрать (нужен WebKitGTK) — её собирает CI; общие части (files.rs, recovery.rs,
// интерфейс desktop.ts) проверяются здесь. Запуск: npm run test:desktop
import { spawn, execSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { chmodSync, copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, utimesSync, writeFileSync } from 'node:fs';
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

const work = mkdtempSync(join(tmpdir(), 'ens-desktop-'));
const ui = join(work, 'ui');
const docs = join(work, 'docs');
const fakebin = join(work, 'bin');
mkdirSync(docs);
mkdirSync(fakebin);

console.log('ENS-Tekst, настоящие файлы (помощник РЕД ОС 7)');
execSync(`node node_modules/vite/bin/vite.js build --mode test --outDir ${ui}`, { cwd: root, stdio: 'pipe' });
execSync('cargo build --quiet', { cwd: join(root, 'redos7'), stdio: 'inherit' });
const helper = join(root, 'redos7', 'target', 'debug', 'ens-text-r7');
ok(existsSync(helper), 'помощник не собрался');

// Подставной браузер: просто ждёт (окно рисует Playwright). Подставной zenity: отвечает путём из ENS_TEST_PICK.
writeFileSync(join(fakebin, 'fakebrowser'), '#!/bin/sh\nexec sleep 3600\n');
writeFileSync(join(fakebin, 'zenity'), '#!/bin/sh\ncat "$ENS_TEST_PICK_FILE" 2>/dev/null\n');
chmodSync(join(fakebin, 'fakebrowser'), 0o755);
chmodSync(join(fakebin, 'zenity'), 0o755);
const pickFile = join(work, 'pick.txt');
writeFileSync(pickFile, '');

function ok(c, what) {
  if (!c) throw new Error(what);
}
function eq(a, b, what) {
  if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(`${what}: ожидалось ${JSON.stringify(b)}, получено ${JSON.stringify(a)}`);
}

const dataHome = join(work, 'data');
function startHelper(...args) {
  const p = spawn(helper, args, {
    env: { ...process.env, HOME: work, XDG_DATA_HOME: dataHome, ENS_BROWSER: join(fakebin, 'fakebrowser'), ENS_UI_DIR: ui, PATH: `${fakebin}:${process.env.PATH}`, ENS_TEST_PICK_FILE: pickFile },
    stdio: 'ignore',
  });
  return p;
}
async function instance(previous = '') {
  const f = join(dataHome, 'ens-text', 'instance.json');
  for (let i = 0; i < 100; i++) {
    if (existsSync(f)) {
      const raw = readFileSync(f, 'utf8');
      if (raw && raw !== previous) return JSON.parse(raw);
    }
    await new Promise((r) => setTimeout(r, 100));
  }
  throw new Error('помощник не запустился');
}

/** Ждём, пока помощник начнёт отвечать (после запуска файл instance.json появляется раньше сервера). */
async function waitUp(i) {
  for (let n = 0; n < 100; n++) {
    if (await fetch(`http://127.0.0.1:${i.port}/index.html`).then((r) => r.ok, () => false)) return;
    await new Promise((r) => setTimeout(r, 100));
  }
  throw new Error('помощник не отвечает');
}

const browser = await pw.chromium.launch();
const errors = [];
let failed = 0;
async function test(name, f) {
  try {
    await f();
    console.log(`  ✓ ${name}`);
  } catch (e) {
    failed++;
    console.log(`  ✗ ${name}\n      ${e.message}`);
  }
}
async function openApp(inst, query = '') {
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  const p = await ctx.newPage();
  p.on('pageerror', (e) => errors.push(e.message));
  await p.goto(`http://127.0.0.1:${inst.port}/index.html?${query}${query ? '&' : ''}t=${inst.token}`);
  await p.waitForSelector('.ProseMirror');
  await p.waitForFunction(() => window.ens);
  return p;
}
const text = (p) => p.evaluate(() => window.ens.view.state.doc.textBetween(0, window.ens.view.state.doc.content.size, '|'));
const readDocx = (p, path) => p.evaluate((arr) => window.ens.importDocx(new Uint8Array(arr)).doc.textContent, [...readFileSync(path)]);

let helperProc = startHelper();
let inst = await instance();
await waitUp(inst);

copyFileSync(join(here, 'fixtures', 'sample.docx'), join(docs, 'sample.docx'));
copyFileSync(join(here, 'fixtures', 'header.docx'), join(docs, 'header.docx'));
writeFileSync(join(docs, 'книга.docx'), readFileSync(join(here, 'fixtures', 'sample.docx')));
writeFileSync(join(docs, 'not.docx'), 'просто текст, а не архив');

await test('открытие файла с диска и сохранение на место (атомарно, без временных файлов)', async () => {
  const path = join(docs, 'sample.docx');
  const p = await openApp(inst, `open=${encodeURIComponent(path)}`);
  await p.waitForFunction(() => window.ens.view.state.doc.textContent.includes('Заголовок первого уровня'));
  ok(await p.evaluate(() => document.title.includes('sample')), 'имя файла в заголовке');
  await p.click('.ProseMirror p >> nth=0');
  await p.keyboard.press('Control+Home');
  await p.keyboard.type('ПРАВКА ');
  await p.keyboard.press('Control+KeyS');
  await p.waitForFunction(() => !window.ens.isDirty());
  ok((await readDocx(p, path)).includes('ПРАВКА '), 'правка записана в файл');
  eq(readdirSync(docs).filter((n) => n.includes('.tmp') || n.startsWith('.ens')), [], 'временные файлы остались');
  await p.context().close();
});

await test('файл не docx: понятное объяснение, документ остаётся пустым', async () => {
  const p = await openApp(inst, `open=${encodeURIComponent(join(docs, 'not.docx'))}`);
  await p.waitForSelector('#toast:not([hidden])');
  ok((await p.textContent('#toast')).includes('а не документ Word'), await p.textContent('#toast'));
  eq(await text(p), '', 'документ не тронут');
  await p.context().close();
});

await test('документ с потерями: «Сохранить» поверх исходного не пишет, «Сохранить как» пишет копию', async () => {
  const path = join(docs, 'header.docx');
  const before = readFileSync(path);
  const p = await openApp(inst, `open=${encodeURIComponent(path)}`);
  await p.waitForSelector('#notice:not([hidden])');
  await p.click('.ProseMirror p >> nth=0');
  await p.keyboard.type('Добавка ');
  await p.keyboard.press('Control+KeyS');
  await p.waitForSelector('#dialog:not([hidden])');
  ok((await p.textContent('#dialog')).includes('Поверх этого файла сохранять нельзя'), 'предупреждение');
  eq(readFileSync(path).equals(before), true, 'исходный файл изменён!');
  const copy = join(docs, 'копия.docx');
  writeFileSync(pickFile, copy);
  await p.click('#dialog button.primary');
  await p.waitForFunction(() => !window.ens.isDirty());
  ok(existsSync(copy), 'копия создана');
  eq(readFileSync(path).equals(before), true, 'исходный файл после копии');
  ok((await readDocx(p, copy)).includes('Добавка'), 'правка в копии');
  // Теперь это наш файл: обычное сохранение на место работает.
  await p.keyboard.type('ещё ');
  await p.keyboard.press('Control+KeyS');
  await p.waitForFunction(() => !window.ens.isDirty());
  ok((await readDocx(p, copy)).includes('ещё'), 'сохранение копии на место');
  await p.context().close();
});

await test('файл изменили снаружи: перед перезаписью спрашивает', async () => {
  const path = join(docs, 'книга.docx');
  const p = await openApp(inst, `open=${encodeURIComponent(path)}`);
  await p.waitForFunction(() => window.ens.view.state.doc.textContent.length > 10);
  const t = new Date(Date.now() + 5000);
  utimesSync(path, t, t); // «другая программа» записала файл
  let asked = '';
  p.on('dialog', async (d) => { asked = d.message(); await d.accept(); });
  await p.click('.ProseMirror p >> nth=0');
  await p.keyboard.type('Моя версия ');
  await p.keyboard.press('Control+KeyS');
  await p.waitForFunction(() => !window.ens.isDirty());
  ok(asked.includes('изменён другой программой'), `вопрос не задан: ${asked}`);
  ok((await readDocx(p, path)).includes('Моя версия'), 'после согласия файл записан');
  await p.context().close();
});

await test('восстановление после аварийного завершения (процесс убит)', async () => {
  const p = await openApp(inst);
  await p.click('.ProseMirror');
  await p.keyboard.type('Текст, который нельзя потерять при сбое');
  await p.waitForTimeout(2500);
  // Сбой: убиваем помощника и «браузер» без нормального завершения.
  const old = readFileSync(join(dataHome, 'ens-text', 'instance.json'), 'utf8');
  helperProc.kill('SIGKILL');
  execSync(`pkill -f '[s]leep 3600' || true`);
  await p.context().close();
  helperProc = startHelper();
  inst = await instance(old);
  await waitUp(inst);
  const p2 = await openApp(inst);
  await p2.waitForSelector('#dialog:not([hidden])');
  ok((await p2.textContent('#dialog')).includes('Восстановление документов'), 'окно восстановления');
  await p2.click('#dialog button.primary');
  await p2.waitForFunction(() => window.ens.view.state.doc.textContent.includes('нельзя потерять'));
  eq(await p2.evaluate(() => window.ens.isDirty()), true, 'восстановленный документ помечен несохранённым');
  // Сохранить восстановленное — временные файлы убираются.
  const saved = join(docs, 'восстановленный.docx');
  writeFileSync(pickFile, saved);
  await p2.keyboard.press('Control+KeyS');
  await p2.waitForFunction(() => !window.ens.isDirty());
  ok((await readDocx(p2, saved)).includes('нельзя потерять'), 'сохранено в файл');
  await p2.context().close();
});

ok(errors.length === 0, `ошибки страницы: ${errors.join('; ')}`);
await browser.close();
helperProc.kill();
execSync("pkill -f '[s]leep 3600' || true");
console.log(failed ? `\nПровалено: ${failed}` : '\nВсе проверки пройдены');
process.exit(failed ? 1 : 0);
