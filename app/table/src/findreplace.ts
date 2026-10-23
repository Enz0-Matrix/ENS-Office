// «Найти и заменить» (Ctrl+F / Ctrl+H), как в Excel: окно не мешает работать
// с листом, «Найти все» показывает список, по щелчку — переход к ячейке.

import { columnName, type Model } from './engine';

export interface FindHost {
  model(): Model;
  sheet(): number;
  active(): { row: number; column: number };
  goTo(sheet: number, row: number, column: number): void;
  changed(): void;
  toast(text: string): void;
}

interface Found { sheet: number; row: number; column: number }

const panel = document.createElement('div');
panel.className = 'find-panel';
panel.hidden = true;
panel.innerHTML = `
  <div class="find-head" data-drag>
    <button class="find-tab" data-tab="find">Найти</button>
    <button class="find-tab" data-tab="replace">Заменить</button>
    <button class="find-close" title="Закрыть (Esc)">✕</button>
  </div>
  <label class="find-row">Найти: <input data-f="text" autocomplete="off" spellcheck="false"></label>
  <label class="find-row find-repl">Заменить на: <input data-f="with" autocomplete="off" spellcheck="false"></label>
  <div class="find-opts">
    <label><input type="checkbox" data-f="case"> Учитывать регистр</label>
    <label><input type="checkbox" data-f="whole"> Ячейка целиком</label>
    <label>Искать: <select data-f="where"><option value="sheet">на листе</option><option value="book">в книге</option></select></label>
    <label class="find-only">Область: <select data-f="look"><option value="values">значения</option><option value="formulas">формулы</option></select></label>
  </div>
  <div class="find-buttons">
    <button data-act="replaceAll" class="find-repl">Заменить все</button>
    <button data-act="replace" class="find-repl">Заменить</button>
    <button data-act="findAll">Найти все</button>
    <button data-act="next" class="primary">Найти далее</button>
  </div>
  <div class="find-status"></div>
  <div class="find-list"></div>`;
document.body.append(panel);

const $ = <T extends HTMLElement>(sel: string) => panel.querySelector<T>(sel)!;
const text = $<HTMLInputElement>('[data-f="text"]');
const withInput = $<HTMLInputElement>('[data-f="with"]');
const matchCase = $<HTMLInputElement>('[data-f="case"]');
const whole = $<HTMLInputElement>('[data-f="whole"]');
const where = $<HTMLSelectElement>('[data-f="where"]');
const look = $<HTMLSelectElement>('[data-f="look"]');
const status = $('.find-status');
const list = $('.find-list');

let host: FindHost;
let mode: 'find' | 'replace' = 'find';

export function initFind(h: FindHost): void {
  host = h;
}

export function showFind(m: 'find' | 'replace', initial = ''): void {
  mode = m;
  panel.classList.toggle('replace', m === 'replace');
  panel.querySelectorAll<HTMLButtonElement>('.find-tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === m));
  panel.hidden = false;
  if (initial) text.value = initial;
  text.focus();
  text.select();
}

export function hideFind(): void {
  panel.hidden = true;
}

function search(): Found[] {
  const inFormulas = mode === 'replace' || look.value === 'formulas';
  try {
    return host.model().findAll(text.value, matchCase.checked, whole.checked, inFormulas, where.value === 'sheet' ? host.sheet() : -1);
  } catch (e) {
    host.toast(String(e));
    return [];
  }
}

/** Следующая найденная ячейка после активной (по строкам, затем по листам). */
function findNext(): Found | undefined {
  const all = search();
  if (!all.length) {
    status.textContent = `Не удалось найти «${text.value}».`;
    return undefined;
  }
  const s = host.sheet(), a = host.active();
  const after = (f: Found) => f.sheet > s || (f.sheet === s && (f.row > a.row || (f.row === a.row && f.column > a.column)));
  const next = all.find(after) ?? all[0];
  status.textContent = `Найдено ячеек: ${all.length}`;
  host.goTo(next.sheet, next.row, next.column);
  return next;
}

function sheetName(i: number): string {
  return host.model().getWorksheetsProperties()[i]?.name ?? '';
}

function showAll(): void {
  const all = search();
  status.textContent = all.length ? `Найдено ячеек: ${all.length}` : `Не удалось найти «${text.value}».`;
  list.replaceChildren();
  for (const f of all.slice(0, 2000)) {
    const b = document.createElement('button');
    const value = host.model().getFormattedCellValue(f.sheet, f.row, f.column);
    b.innerHTML = `<span>${esc(sheetName(f.sheet))}</span><span>${columnName(f.column)}${f.row}</span><span>${esc(value)}</span>`;
    b.addEventListener('click', () => host.goTo(f.sheet, f.row, f.column));
    list.append(b);
  }
}

function replaceOne(): void {
  const s = host.sheet(), a = host.active();
  const m = host.model();
  // Если активная ячейка подходит — заменяем в ней, затем ищем следующую.
  const hit = search().find((f) => f.sheet === s && f.row === a.row && f.column === a.column);
  if (hit) {
    m.replaceAll(text.value, withInput.value, matchCase.checked, whole.checked, s, a.row, a.column);
    host.changed();
  }
  findNext();
}

function replaceAll(): void {
  const n = host.model().replaceAll(text.value, withInput.value, matchCase.checked, whole.checked,
    where.value === 'sheet' ? host.sheet() : -1, 0, 0);
  host.changed();
  status.textContent = n ? `Замен: ${n}. Отменить — Ctrl+Z.` : `Не удалось найти «${text.value}».`;
  list.replaceChildren();
}

panel.addEventListener('click', (e) => {
  const t = e.target as HTMLElement;
  if (t.closest('.find-close')) { hideFind(); return; }
  const tab = t.closest<HTMLElement>('.find-tab')?.dataset.tab;
  if (tab) { showFind(tab as 'find' | 'replace'); return; }
  const act = t.closest<HTMLElement>('[data-act]')?.dataset.act;
  if (!act || !text.value) return;
  if (act === 'next') findNext();
  else if (act === 'findAll') showAll();
  else if (act === 'replace') replaceOne();
  else if (act === 'replaceAll') replaceAll();
});

panel.addEventListener('keydown', (e) => {
  e.stopPropagation(); // клавиши окна не уходят в таблицу
  if (e.key === 'Escape') { e.preventDefault(); hideFind(); }
  if (e.key === 'Enter' && text.value) { e.preventDefault(); findNext(); }
});

// Окно можно перетаскивать за заголовок.
const head = $('.find-head');
head.addEventListener('pointerdown', (e) => {
  if ((e.target as HTMLElement).closest('button')) return;
  const r = panel.getBoundingClientRect();
  const dx = e.clientX - r.left, dy = e.clientY - r.top;
  const move = (ev: PointerEvent) => {
    panel.style.left = `${Math.max(0, ev.clientX - dx)}px`;
    panel.style.top = `${Math.max(0, ev.clientY - dy)}px`;
    panel.style.right = 'auto';
  };
  const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); };
  window.addEventListener('pointermove', move);
  window.addEventListener('pointerup', up);
});

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
