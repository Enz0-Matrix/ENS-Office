// Автофильтр, как в Excel: кнопка в заголовке столбца → сортировка, поиск,
// флажки значений. Отбор скрывает строки (их высота становится 0), формулы
// и данные не меняются.

import type { Model } from './engine';

export interface SheetFilter {
  r1: number; c1: number; r2: number; c2: number;
  /** Столбец → показываемые значения (как они видны в ячейке). */
  criteria: Map<number, Set<string>>;
}

export interface FilterHost {
  model(): Model;
  sheet(): number;
  sort(column: number, descending: boolean): void;
  changed(): void;
}

const EMPTY = '(Пустые)';
/** Как в Excel, отбор не различает регистр: «Москва» и «москва» — одно значение. */
const key = (shown: string) => (shown || EMPTY).toLowerCase();
const filters = new Map<number, SheetFilter>();
let host: FilterHost;

export function initFilter(h: FilterHost): void {
  host = h;
}

export function filterOf(sheet: number): SheetFilter | undefined {
  let f = filters.get(sheet);
  if (!f && host) {
    // Фильтр из файла (Excel хранит его область в книге): кнопки возвращаются.
    const r = (() => { try { return host.model().getAutoFilter(sheet); } catch { return []; } })();
    if (r.length === 4) {
      f = { r1: r[0], c1: r[1], r2: r[2], c2: r[3], criteria: new Map() };
      filters.set(sheet, f);
    }
  }
  return f;
}

/** Новая книга в окне: фильтры прежней книги забываем. */
export function resetFilters(): void {
  filters.clear();
}

/** Область фильтра — в книгу: сохранится в xlsx, как у Excel. */
function persist(sheet: number, f: SheetFilter | null): void {
  try { host.model().setAutoFilter(sheet, f ? [f.r1, f.c1, f.r2, f.c2] : []); } catch { /* старый движок */ }
}

/** Строки данных фильтра: до последней непустой строки под заголовком. */
function extend(f: SheetFilter): void {
  const m = host.model(), s = host.sheet();
  let last = f.r1;
  for (let c = f.c1; c <= f.c2; c++) {
    const rows = m.getRowsWithData(s, c);
    for (const r of rows) if (r > last) last = r;
  }
  f.r2 = Math.max(f.r2, last);
}

export function setFilter(sheet: number, f: SheetFilter | null): void {
  const old = filterOf(sheet);
  if (old) {
    // Снять фильтр — показать все строки.
    host.model().setRowsHidden(sheet, old.r1 + 1, old.r2, false);
    filters.delete(sheet);
  }
  if (f) {
    filters.set(sheet, f);
    extend(f);
  }
  persist(sheet, f);
  host.changed();
}

function apply(f: SheetFilter): void {
  const m = host.model(), s = host.sheet();
  extend(f);
  let runStart = 0, runHidden = false;
  const flush = (end: number) => { if (runStart) m.setRowsHidden(s, runStart, end, runHidden); };
  for (let r = f.r1 + 1; r <= f.r2; r++) {
    let show = true;
    for (const [c, allowed] of f.criteria) {
      if (!allowed.has(key(m.getFormattedCellValue(s, r, c)))) { show = false; break; }
    }
    if (!runStart) { runStart = r; runHidden = !show; continue; }
    if (runHidden !== !show) { flush(r - 1); runStart = r; runHidden = !show; }
  }
  flush(f.r2);
  persist(s, f);
  host.changed();
}

const popup = document.createElement('div');
popup.className = 'filter-popup';
popup.hidden = true;
document.body.append(popup);

document.addEventListener('mousedown', (e) => {
  if (!popup.hidden && !popup.contains(e.target as Node)) popup.hidden = true;
});

/** Окно фильтра столбца под кнопкой. */
export function openFilterPopup(column: number, x: number, y: number): void {
  const f = filters.get(host.sheet());
  if (!f) return;
  extend(f);
  const m = host.model(), s = host.sheet();
  const values = new Map<string, string>(); // ключ → как показать
  for (let r = f.r1 + 1; r <= f.r2; r++) {
    // Учитываем отбор по другим столбцам: в списке — только видимые варианты.
    let ok = true;
    for (const [c, allowed] of f.criteria) {
      if (c !== column && !allowed.has(key(m.getFormattedCellValue(s, r, c)))) { ok = false; break; }
    }
    if (!ok) continue;
    const shown = m.getFormattedCellValue(s, r, column) || EMPTY;
    if (!values.has(key(shown))) values.set(key(shown), shown);
  }
  const sorted = [...values.values()].sort((a, b) => {
    if (a === EMPTY) return 1;
    if (b === EMPTY) return -1;
    const na = Number(a.replace(/\s/g, '').replace(',', '.')), nb = Number(b.replace(/\s/g, '').replace(',', '.'));
    return Number.isFinite(na) && Number.isFinite(nb) ? na - nb : a.localeCompare(b, 'ru');
  });
  const current = f.criteria.get(column);
  popup.innerHTML = `
    <button data-act="asc">↑ Сортировка от А до Я</button>
    <button data-act="desc">↓ Сортировка от Я до А</button>
    <button data-act="clear" ${current ? '' : 'disabled'}>Снять фильтр со столбца</button>
    <hr>
    <input class="filter-search" placeholder="Поиск" autocomplete="off">
    <label class="filter-all"><input type="checkbox" data-all checked> (Выделить все)</label>
    <div class="filter-values"></div>
    <div class="filter-buttons"><button data-act="ok" class="primary">ОК</button><button data-act="cancel">Отмена</button></div>`;
  const box = popup.querySelector<HTMLElement>('.filter-values')!;
  for (const v of sorted) {
    const l = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = key(v);
    cb.checked = !current || current.has(key(v));
    l.append(cb, ` ${v}`);
    box.append(l);
  }
  const boxes = () => [...box.querySelectorAll<HTMLInputElement>('input')];
  const all = popup.querySelector<HTMLInputElement>('[data-all]')!;
  const sync = () => { const b = boxes().filter((x) => !x.parentElement!.hidden); all.checked = b.every((x) => x.checked); };
  sync();
  all.addEventListener('change', () => boxes().forEach((b) => { if (!b.parentElement!.hidden) b.checked = all.checked; }));
  box.addEventListener('change', sync);
  const search = popup.querySelector<HTMLInputElement>('.filter-search')!;
  search.addEventListener('input', () => {
    const q = search.value.toLowerCase();
    // Как в Excel: поиск сразу отмечает найденное.
    for (const b of boxes()) { const hit = b.value.toLowerCase().includes(q); b.parentElement!.hidden = !hit; b.checked = hit; }
    sync();
  });
  popup.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') popup.hidden = true;
    if (e.key === 'Enter') (popup.querySelector('[data-act="ok"]') as HTMLButtonElement).click();
  });
  popup.onclick = (e) => {
    const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
    if (!act) return;
    if (act === 'asc' || act === 'desc') { popup.hidden = true; host.sort(column, act === 'desc'); return; }
    if (act === 'clear') { f.criteria.delete(column); popup.hidden = true; apply(f); return; }
    if (act === 'cancel') { popup.hidden = true; return; }
    const chosen = new Set(boxes().filter((b) => b.checked && !b.parentElement!.hidden).map((b) => b.value));
    if (chosen.size === sorted.length) f.criteria.delete(column);
    else f.criteria.set(column, chosen);
    popup.hidden = true;
    apply(f);
  };
  popup.hidden = false;
  const r = popup.getBoundingClientRect();
  popup.style.left = `${Math.max(8, Math.min(x, window.innerWidth - r.width - 8))}px`;
  popup.style.top = `${Math.max(8, Math.min(y, window.innerHeight - r.height - 8))}px`;
  search.focus();
}
