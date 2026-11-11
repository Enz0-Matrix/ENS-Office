import { spawn, execSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');
let pw;
try { pw = createRequire(import.meta.url)('playwright'); }
catch { pw = createRequire(join(execSync('npm root -g').toString().trim(), 'x.js'))('playwright'); }

const PORT = 5177;
const server = spawn(process.execPath, [join(root, 'node_modules/vite/bin/vite.js'), '--port', String(PORT), '--strictPort'], { cwd: root, stdio: 'pipe' });
await new Promise((res, rej) => {
  server.stdout.on('data', (d) => { if (String(d).includes('Local')) res(); });
  server.on('exit', () => rej(new Error('vite не запустился')));
  setTimeout(() => rej(new Error('vite: таймаут')), 30000);
});

const browser = await pw.chromium.launch();
let failed = 0;
const errors = [];
async function page() {
  const ctx = await browser.newContext({ viewport: { width: 1280, height: 900 } });
  const p = await ctx.newPage();
  p.on('pageerror', (e) => errors.push(e.message));
  await p.goto(`http://localhost:${PORT}/`);
  await p.waitForSelector('.ProseMirror');
  await p.waitForFunction(() => window.ens);
  return p;
}
function eq(a, b, what) {
  const [x, y] = [JSON.stringify(a), JSON.stringify(b)];
  if (x !== y) throw new Error(`${what}: ожидалось ${y}, получено ${x}`);
}
function ok(c, what) { if (!c) throw new Error(what); }
async function test(name, f) {
  try { await f(); console.log(`  ✓ ${name}`); }
  catch (e) { failed++; console.log(`  ✗ ${name}\n      ${e.message}`); }
}

console.log('ENS-Text: меню «Файл»');

await test('меню «Файл»: кнопка «Закрыть» видна, в фокусе и закрывает меню', async () => {
  const p = await page();
  await p.click('.tab.file');
  await p.waitForSelector('#dialog:not([hidden]) .menu-close', { timeout: 3000 });
  const st = await p.evaluate(() => {
    const b = document.querySelector('#dialog .menu-close');
    const r = b.getBoundingClientRect();
    return { text: b.textContent.trim(), icon: !!b.querySelector('svg'), w: Math.round(r.width), h: Math.round(r.height), title: b.getAttribute('title'), focused: document.activeElement === b };
  });
  eq(st.text, 'Закрыть', 'подпись кнопки');
  ok(st.icon, 'значок как в ENS-Table');
  ok(st.w > 40 && st.h > 16, `кнопка не нулевая: ${JSON.stringify(st)}`);
  ok(/Esc/.test(st.title || ''), 'подсказка про Esc');
  ok(st.focused, 'кнопка в фокусе');
  await p.click('#dialog .menu-close');
  await p.waitForFunction(() => document.getElementById('dialog').hidden, null, { timeout: 3000 });
  ok(true, 'меню закрылось кнопкой');
  await p.context().close();
});

await test('меню открыто: правая кнопка и Ctrl+R/F5 не проходят; в полях ввода меню остаётся', async () => {
  const p = await page();
  const outcome = () => p.evaluate(() => {
    const fire = (el, init) => { const e = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init }); el.dispatchEvent(e); return e.defaultPrevented; };
    const ctx = new MouseEvent('contextmenu', { bubbles: true, cancelable: true });
    document.querySelector('.ProseMirror').dispatchEvent(ctx);
    return {
      ctx: ctx.defaultPrevented,
      ctrlR: fire(document.body, { key: 'r', ctrlKey: true }),
      f5: fire(document.body, { key: 'F5' }),
    };
  });
  eq(await outcome(), { ctx: false, ctrlR: false, f5: false }, 'без меню ничего не подавляется');

  await p.click('.tab.file');
  await p.waitForSelector('#dialog:not([hidden])', { timeout: 3000 });
  eq(await outcome(), { ctx: true, ctrlR: true, f5: true }, 'при открытом меню перезапуск заблокирован');

  await p.keyboard.press('Escape');
  await p.waitForFunction(() => document.getElementById('dialog').hidden, null, { timeout: 3000 });

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

  eq(await outcome(), { ctx: false, ctrlR: false, f5: false }, 'после закрытия снова как обычно');
  await p.context().close();
});


ok(errors.length === 0, `ошибки страницы: ${errors.join('; ')}`);
await browser.close();
server.kill();
console.log(failed ? `\nПровалено: ${failed}` : '\nВсе проверки пройдены');
process.exit(failed ? 1 : 0);
