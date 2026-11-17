// SPDX-License-Identifier: AGPL-3.0-or-later
// Выбор шрифта как в Word: каждый шрифт показан своим начертанием, рядом — есть ли кириллица и чем он нарисуется,
// если его нет в системе (встроенный аналог с теми же размерами букв, см. fonts.ts). Окно открывается поверх обычного
// списка ленты (select[data-command="format.fontName"]) и выбирает в нём значение — команда ленты та же.
// Шрифты системы находятся замером по каталогу распространённых имён (font-catalog.ts) — без окна разрешения.
// Полный список (Local Font Access, Chromium 103+) — только по явной кнопке: браузер спросит разрешение.

import { fontStack, OFFICE_FONTS } from '../../table/src/fonts';
import { FONT_CATALOG } from './font-catalog';

const RECENT = 'ens-text.recentFonts';
const SAMPLE = 'АаБбЖж Aa';

interface Info { installed: boolean; cyr: boolean; via: string | null }

let canvas: CanvasRenderingContext2D | null = null;
const cache = new Map<string, Info>();
let systemFonts: { family: string; styles: Set<string> }[] | null = null;

const width = (text: string, family: string): number => {
  canvas ??= document.createElement('canvas').getContext('2d');
  if (!canvas) return 0;
  canvas.font = `40px ${family}`;
  return canvas.measureText(text).width;
};

/** Рисует ли браузер text шрифтом family (а не запасным): ширина не зависит от запасного шрифта. */
function draws(family: string, text: string): boolean {
  const a = width(text, `"${family}", monospace`), b = width(text, `"${family}", serif`);
  return Math.abs(a - b) < 0.01 && !(Math.abs(a - width(text, 'monospace')) < 0.01 && Math.abs(b - width(text, 'serif')) < 0.01);
}

/** Встроенный аналог шрифта Office (Times New Roman → Tinos) или null. */
function substitute(name: string): string | null {
  const parts = fontStack(name).split(', ').map((x) => x.replace(/"/g, ''));
  const tail = fontStack().split(', ').map((x) => x.replace(/"/g, ''));
  return parts.length > tail.length + 1 ? parts[1] : null;
}

/** Есть ли шрифт и кириллица в нём; если шрифта нет — чем он нарисуется (аналог) и есть ли кириллица в аналоге. */
export function fontInfo(name: string): Info {
  const hit = cache.get(name);
  if (hit) return hit;
  const installed = draws(name, 'abcmwXYZ');
  const sub = substitute(name);
  const via = installed ? null : sub && draws(sub, 'abcmwXYZ') ? sub : null;
  const face = installed ? name : via;
  const info = { installed, via, cyr: face ? draws(face, 'ЖжЩщЯяЫы') : false };
  cache.set(name, info);
  return info;
}

const recent = (): string[] => { try { const v = JSON.parse(localStorage.getItem(RECENT) ?? '[]'); return Array.isArray(v) ? v.slice(0, 6) : []; } catch { return []; } };
const remember = (name: string) => { try { localStorage.setItem(RECENT, JSON.stringify([name, ...recent().filter((x) => x !== name)].slice(0, 6))); } catch { /* без хранилища */ } };

let pop: HTMLElement | null = null;
let owner: HTMLSelectElement | null = null;

export function closeFontPicker(): void {
  pop?.remove();
  pop = null;
  owner = null;
}

function choose(sel: HTMLSelectElement, name: string): void {
  if (![...sel.options].some((o) => o.value === name)) sel.add(new Option(name, name));
  sel.value = name;
  remember(name);
  closeFontPicker();
  sel.dispatchEvent(new Event('change', { bubbles: true }));
}

function row(sel: HTMLSelectElement, name: string, styles?: Set<string>): HTMLElement {
  const info = fontInfo(name);
  const b = document.createElement('button');
  b.type = 'button';
  b.className = 'fp-row';
  b.setAttribute('role', 'option');
  b.dataset.font = name;
  if (sel.value === name) b.classList.add('on');
  const title = document.createElement('span');
  title.className = 'fp-name';
  title.textContent = name;
  title.style.fontFamily = fontStack(name);
  const sample = document.createElement('span');
  sample.className = 'fp-sample';
  sample.textContent = SAMPLE;
  sample.style.fontFamily = fontStack(name);
  const tags = document.createElement('span');
  tags.className = 'fp-tags';
  const tag = (text: string, cls: string, tip: string) => { const t = document.createElement('span'); t.className = `fp-tag ${cls}`; t.textContent = text; t.title = tip; tags.append(t); };
  if (!info.installed && !info.via) tag('нет в системе', 'warn', 'Шрифта нет в системе: текст будет нарисован запасным шрифтом. В файле имя шрифта сохраняется.');
  else {
    if (info.cyr) tag('кириллица', 'ok', 'Русские буквы есть в шрифте');
    else tag('без кириллицы', 'warn', 'Русских букв в шрифте нет: они будут нарисованы другим шрифтом');
    if (info.via) tag(`аналог: ${info.via}`, 'sub', `Шрифта «${name}» нет в системе: рисуется встроенный аналог с теми же размерами букв. В файле остаётся «${name}».`);
  }
  if (styles?.size) tag([...styles].slice(0, 4).join(', '), 'styles', `Начертания в системе: ${[...styles].join(', ')}`);
  b.append(title, sample, tags);
  b.addEventListener('click', () => choose(sel, name));
  return b;
}

let detected: string[] | null = null;
/** Шрифты каталога, которые есть в системе (замер, без разрешения). */
export function detectedFonts(): string[] {
  detected ??= [...new Set(FONT_CATALOG)].filter((n) => draws(n, 'abcmwXYZ')).sort((a, b) => a.localeCompare(b, 'ru'));
  return detected;
}

async function loadSystemFonts(): Promise<void> {
  const q = (window as unknown as { queryLocalFonts?: () => Promise<{ family: string; style: string }[]> }).queryLocalFonts;
  if (!q) throw new Error('нет доступа');
  const list = await q();
  const map = new Map<string, Set<string>>();
  for (const f of list) {
    if (!map.has(f.family)) map.set(f.family, new Set());
    map.get(f.family)!.add(f.style);
  }
  systemFonts = [...map].map(([family, styles]) => ({ family, styles })).sort((a, b) => a.family.localeCompare(b.family, 'ru'));
}

/** Подгрузить встроенные (веб-)шрифты, чтобы замер видел их, а не запасной шрифт; не дольше 1,5 с. */
async function preload(names: string[]): Promise<void> {
  const fams = new Set<string>();
  for (const n of names) { fams.add(n); const sub = substitute(n); if (sub) fams.add(sub); }
  const all = Promise.all([...fams].map((f) => document.fonts.load(`40px "${f}"`, 'abcЖж').catch(() => [])));
  await Promise.race([all, new Promise((r) => setTimeout(r, 1500))]);
  cache.clear();
}

export async function openFontPicker(sel: HTMLSelectElement): Promise<void> {
  if (owner === sel && pop) { closeFontPicker(); return; }
  await preload([...OFFICE_FONTS, ...recent(), ...[...sel.options].map((o) => o.value)]);
  closeFontPicker();
  owner = sel;
  pop = document.createElement('div');
  pop.className = 'font-picker';
  pop.setAttribute('role', 'listbox');
  pop.setAttribute('aria-label', 'Шрифт');
  const search = document.createElement('input');
  search.type = 'text';
  search.placeholder = 'Найти шрифт';
  search.className = 'fp-search';
  const list = document.createElement('div');
  list.className = 'fp-list';
  pop.append(search, list);

  const draw = () => {
    const q = search.value.trim().toLowerCase();
    const fit = (n: string) => !q || n.toLowerCase().includes(q);
    list.replaceChildren();
    const section = (title: string, names: { family: string; styles?: Set<string> }[]) => {
      const shown = names.filter((x) => fit(x.family));
      if (!shown.length) return;
      const h = document.createElement('div');
      h.className = 'fp-head';
      h.textContent = title;
      list.append(h, ...shown.map((x) => row(sel, x.family, x.styles)));
    };
    const office = [...new Set([...OFFICE_FONTS, ...[...sel.options].map((o) => o.value)])].filter(Boolean);
    section('Недавние', recent().map((family) => ({ family })));
    section('Шрифты Office и документа', office.map((family) => ({ family })));
    const shownNames = new Set(office);
    const sys = systemFonts ?? detectedFonts().map((family) => ({ family, styles: undefined as Set<string> | undefined }));
    section('Шрифты системы', sys.filter((x) => !shownNames.has(x.family)));
    if (!systemFonts) {
      const more = document.createElement('button');
      more.type = 'button';
      more.className = 'fp-more';
      more.textContent = 'Найти остальные шрифты…';
      more.title = 'Показать все шрифты компьютера, включая редкие. Окно программы спросит разрешение на доступ к списку шрифтов — его можно дать один раз.';
      more.addEventListener('click', () => {
        more.disabled = true;
        more.textContent = 'Загрузка…';
        loadSystemFonts().then(draw).catch(() => { more.textContent = 'Полный список недоступен: показаны найденные шрифты'; });
      });
      list.append(more);
    }
    if (q && !list.querySelector('.fp-row')) {
      // Свой шрифт по имени (например, установленный позже): выбрать как есть.
      list.append(row(sel, search.value.trim()));
    }
  };
  search.addEventListener('input', draw);
  search.addEventListener('keydown', (e) => {
    const rows = [...list.querySelectorAll<HTMLElement>('.fp-row')];
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeFontPicker(); sel.focus(); }
    else if (e.key === 'Enter') { e.preventDefault(); const first = rows[0]; if (first) choose(sel, first.dataset.font!); }
    else if (e.key === 'ArrowDown') { e.preventDefault(); rows[0]?.focus(); }
  });
  list.addEventListener('keydown', (e) => {
    const rows = [...list.querySelectorAll<HTMLElement>('.fp-row')];
    const i = rows.indexOf(document.activeElement as HTMLElement);
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); rows[Math.max(0, Math.min(rows.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1)))]?.focus(); }
    else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); closeFontPicker(); sel.focus(); }
  });
  draw();
  document.body.append(pop);
  const r = sel.getBoundingClientRect();
  pop.style.left = `${Math.max(4, Math.min(r.left, window.innerWidth - pop.offsetWidth - 4))}px`;
  pop.style.top = `${r.bottom + 2}px`;
  search.focus();
  list.querySelector<HTMLElement>('.fp-row.on')?.scrollIntoView({ block: 'nearest' });
}

/** Подключить окно к списку шрифтов ленты (после каждой перерисовки ленты). */
export function attachFontPicker(): void {
  const sel = document.querySelector<HTMLSelectElement>('select[data-command="format.fontName"]');
  if (!sel || sel.dataset.picker) return;
  sel.dataset.picker = '1';
  sel.addEventListener('mousedown', (e) => { e.preventDefault(); sel.focus(); void openFontPicker(sel); });
  sel.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ' || e.key === 'F4' || (e.altKey && e.key === 'ArrowDown')) { e.preventDefault(); void openFontPicker(sel); }
  });
}

document.addEventListener('mousedown', (e) => {
  if (pop && !pop.contains(e.target as Node) && e.target !== owner) closeFontPicker();
}, true);
window.addEventListener('resize', closeFontPicker);
