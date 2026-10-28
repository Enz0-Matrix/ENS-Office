// Автофильтр, как в Excel: кнопка в заголовке столбца → сортировка, поиск,
// флажки значений. Отбор скрывает строки (их высота становится 0), формулы
// и данные не меняются.
//
// ENS-Office: на листе может быть несколько таблиц со своими фильтрами.
// Таблица фильтра — от строки заголовков до первой полностью пустой строки:
// всё, что ниже, фильтр не трогает. Первый фильтр листа хранится как в Excel
// (автофильтр), остальные — именами «_ENS_Filter2», «_ENS_Filter3»… со ссылкой
// на диапазон: при вставке и удалении строк он сдвигается вместе с таблицей.

import type { Model } from './engine';

export interface SheetFilter {
  r1: number; c1: number; r2: number; c2: number;
  /** Столбец → показываемые значения (как они видны в ячейке). */
  criteria: Map<number, Set<string>>;
}

export interface FilterHost {
  model(): Model;
  sheet(): number;
  /** Сортировка таблицы фильтра f по столбцу. */
  sort(column: number, descending: boolean, f: SheetFilter): void;
  changed(): void;
}

const EMPTY = '(Пустые)';
const EXTRA = '_ENS_Filter';
/** Как в Excel, отбор не различает регистр: «Москва» и «москва» — одно значение. */
const key = (shown: string) => (shown || EMPTY).toLowerCase();
/** Отбор по цвету хранится в том же наборе: «fill:#ffff00» — заливка, «font:#ff0000» — шрифт («none» — без цвета). */
const COLOR_TOKEN = /^(fill|font):/;
const filters = new Map<number, SheetFilter[]>();
let host: FilterHost;

export function initFilter(h: FilterHost): void {
  host = h;
}

/** «Лист1!$A$30:$D$40» → прямоугольник. */
function parseRef(formula: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const m = /([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\s*$/i.exec(formula.replace(/\$/g, ''));
  if (!m) return null;
  const col = (s: string) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  return { r1: Number(m[2]), c1: col(m[1]), r2: Number(m[4]), c2: col(m[3]) };
}

const colName = (c: number) => { let s = ''; for (let n = c; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

/** Все фильтры листа (из книги — при первом обращении). */
export function filtersOf(sheet: number): SheetFilter[] {
  let list = filters.get(sheet);
  if (!list && host) {
    list = [];
    const m = host.model();
    const r = (() => { try { return m.getAutoFilter(sheet); } catch { return []; } })();
    if (r.length === 4) list.push({ r1: r[0], c1: r[1], r2: r[2], c2: r[3], criteria: new Map() });
    try {
      const extra = m.getDefinedNameList()
        .filter((d) => d.scope === sheet && d.name.startsWith(EXTRA))
        .sort((a, b) => a.name.localeCompare(b.name, 'en', { numeric: true }));
      for (const d of extra) {
        const rect = parseRef(d.formula);
        if (rect) list.push({ ...rect, criteria: new Map() });
      }
    } catch { /* старый движок */ }
    // Отбор (что показано) — из книги: [[[смещение столбца, [значения…]], …], …].
    try {
      const saved = JSON.parse(m.getFilterCriteria(sheet) || '[]') as [number, string[]][][];
      saved.forEach((cols, i) => {
        const f = list![i];
        if (!f || !Array.isArray(cols)) return;
        for (const [off, vals] of cols) if (Array.isArray(vals)) f.criteria.set(f.c1 + off, new Set(vals.map((v) => String(v).toLowerCase())));
      });
    } catch { /* нет отбора или старый движок */ }
    filters.set(sheet, list);
  }
  return list ?? [];
}

/** Фильтр, в таблицу которого попадает ячейка (заголовок или данные). */
export function filterAt(sheet: number, row: number, column: number): SheetFilter | undefined {
  return filtersOf(sheet).find((f) => row >= f.r1 && row <= f.r2 && column >= f.c1 && column <= f.c2);
}

/** Есть ли на листе хоть один фильтр. */
export function filterOf(sheet: number): SheetFilter | undefined {
  return filtersOf(sheet)[0];
}

/** Новая книга в окне: фильтры прежней книги забываем. */
export function resetFilters(): void {
  filters.clear();
}

/** Области фильтров — в книгу: первый — автофильтр Excel, остальные — имена. */
function persist(sheet: number): void {
  const m = host.model();
  const list = filtersOf(sheet);
  try { m.setAutoFilter(sheet, list[0] ? [list[0].r1, list[0].c1, list[0].r2, list[0].c2] : []); } catch { /* старый движок */ }
  try {
    const crit = list.map((f) => [...f.criteria].map(([c, vals]) => [c - f.c1, [...vals]]));
    m.setFilterCriteria(sheet, crit.some((c) => c.length) ? JSON.stringify(crit) : '');
  } catch { /* старый движок */ }
  try {
    const name = m.getWorksheetsProperties()[sheet].name;
    const q = /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
    const want = new Map(list.slice(1).map((f, i) => [`${EXTRA}${i + 2}`, `${q}!$${colName(f.c1)}$${f.r1}:$${colName(f.c2)}$${f.r2}`]));
    const have = m.getDefinedNameList().filter((d) => d.scope === sheet && d.name.startsWith(EXTRA));
    for (const d of have) {
      const w = want.get(d.name);
      if (w === undefined) m.deleteDefinedName(d.name, sheet);
      else if (d.formula.replace(/^=/, '') !== w) m.updateDefinedName(d.name, sheet, d.name, sheet, w);
      want.delete(d.name);
    }
    for (const [n, f] of want) m.newDefinedName(n, sheet, f);
  } catch { /* имена не записались — фильтры живут до закрытия книги */ }
}

/** Цвет заливки или шрифта ячейки как «fill:#rrggbb» / «font:#rrggbb»; без цвета — «…:none». */
function colorKey(kind: 'fill' | 'font', r: number, c: number): string {
  const m = host.model(), s = host.sheet();
  const st = m.getCellStyle(s, r, c).style;
  const raw = kind === 'fill' ? st.fill?.color : st.font.color;
  let hex = '';
  try { hex = raw ? m.resolveColor(raw) : ''; } catch { /* нет цвета */ }
  hex = hex.toLowerCase();
  // Белая заливка и чёрный шрифт — то же, что «без цвета».
  if (!hex || (kind === 'fill' && hex === '#ffffff') || (kind === 'font' && hex === '#000000')) return `${kind}:none`;
  return `${kind}:${hex}`;
}

/** Проходит ли строка r по отбору столбца c: по значению или по цвету. */
function passes(allowed: Set<string>, r: number, c: number): boolean {
  const m = host.model(), s = host.sheet();
  let byColor = false;
  for (const v of allowed) {
    const kind = COLOR_TOKEN.exec(v)?.[1] as 'fill' | 'font' | undefined;
    if (!kind) continue;
    byColor = true;
    if (colorKey(kind, r, c) === v) return true;
  }
  return !byColor && allowed.has(key(m.getFormattedCellValue(s, r, c)));
}

/** Строки таблицы: от заголовка до первой полностью пустой строки (ниже — другая таблица). */
function extend(f: SheetFilter): void {
  const m = host.model(), s = host.sheet();
  const filled = new Set<number>();
  for (let c = f.c1; c <= f.c2; c++) for (const r of m.getRowsWithData(s, c)) if (r > f.r1) filled.add(r);
  let r = f.r1 + 1;
  while (filled.has(r)) r++;
  f.r2 = Math.max(f.r1, r - 1);
}

/** Включить фильтр f; таблицы, которые он задевает по строкам и столбцам, теряют свой. */
export function addFilter(sheet: number, f: SheetFilter): void {
  extend(f);
  const list = filtersOf(sheet);
  for (const o of [...list]) {
    if (o.r1 <= f.r2 && o.r2 >= f.r1 && o.c1 <= f.c2 && o.c2 >= f.c1) removeFilter(sheet, o, false);
  }
  list.push(f);
  list.sort((a, b) => a.r1 - b.r1 || a.c1 - b.c1);
  persist(sheet);
  host.changed();
}

/** Снять фильтр: показать его строки. */
export function removeFilter(sheet: number, f: SheetFilter, notify = true): void {
  const list = filtersOf(sheet);
  const i = list.indexOf(f);
  if (i < 0) return;
  if (f.criteria.size) host.model().setRowsHidden(sheet, f.r1 + 1, f.r2, false);
  list.splice(i, 1);
  if (notify) { persist(sheet); host.changed(); }
}

/** Снять все фильтры листа. */
export function clearFilters(sheet: number): void {
  for (const f of [...filtersOf(sheet)]) removeFilter(sheet, f, false);
  persist(sheet);
  host.changed();
}

function apply(f: SheetFilter): void {
  const m = host.model(), s = host.sheet();
  const before = f.r2;
  extend(f);
  // Строки, что выпали из таблицы (стали отделены пустой строкой), — показать.
  if (before > f.r2) m.setRowsHidden(s, f.r2 + 1, before, false);
  let runStart = 0, runHidden = false;
  const flush = (end: number) => { if (runStart) m.setRowsHidden(s, runStart, end, runHidden); };
  for (let r = f.r1 + 1; r <= f.r2; r++) {
    let show = true;
    for (const [c, allowed] of f.criteria) {
      if (!passes(allowed, r, c)) { show = false; break; }
    }
    if (!runStart) { runStart = r; runHidden = !show; continue; }
    if (runHidden !== !show) { flush(r - 1); runStart = r; runHidden = !show; }
  }
  flush(f.r2);
  persist(s);
  host.changed();
}

const popup = document.createElement('div');
popup.className = 'filter-popup';
popup.hidden = true;
document.body.append(popup);

document.addEventListener('mousedown', (e) => {
  if (!popup.hidden && !popup.contains(e.target as Node)) popup.hidden = true;
});

/** Окно фильтра столбца под кнопкой (row — строка заголовков этой таблицы). */
export function openFilterPopup(row: number, column: number, x: number, y: number): void {
  const f = filtersOf(host.sheet()).find((t) => t.r1 === row && column >= t.c1 && column <= t.c2);
  if (!f) return;
  extend(f);
  const m = host.model(), s = host.sheet();
  const values = new Map<string, string>(); // ключ → как показать
  const colors = new Map<string, string>(); // токен цвета → подпись
  for (let r = f.r1 + 1; r <= f.r2; r++) {
    // Учитываем отбор по другим столбцам: в списке — только видимые варианты.
    let ok = true;
    for (const [c, allowed] of f.criteria) {
      if (c !== column && !passes(allowed, r, c)) { ok = false; break; }
    }
    if (!ok) continue;
    for (const kind of ['fill', 'font'] as const) {
      const token = colorKey(kind, r, column);
      if (!colors.has(token)) colors.set(token, token.endsWith(':none') ? (kind === 'fill' ? 'без заливки' : 'авто') : token.slice(kind.length + 1));
    }
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
    ${colors.size > 2 ? '<div class="filter-colors-title">По цвету</div><div class="filter-colors"></div>' : ''}
    <hr>
    <input class="filter-search" placeholder="Поиск" autocomplete="off">
    <label class="filter-all"><input type="checkbox" data-all checked> (Выделить все)</label>
    <div class="filter-values"></div>
    <div class="filter-buttons"><button data-act="ok" class="primary">ОК</button><button data-act="cancel">Отмена</button></div>`;
  const palette = popup.querySelector<HTMLElement>('.filter-colors');
  for (const [token, label] of colors) {
    if (!palette) break;
    const none = token.endsWith(':none');
    const sw = document.createElement('button');
    sw.className = `filter-swatch${current?.has(token) ? ' on' : ''}${none ? ' none' : ''}`;
    sw.dataset.token = token;
    sw.title = `${token.startsWith('fill') ? 'Заливка' : 'Шрифт'}: ${label}`;
    if (!none) {
      if (token.startsWith('fill')) sw.style.background = label;
      else { sw.style.color = label; sw.textContent = 'A'; }
    } else sw.textContent = token.startsWith('fill') ? '∅' : 'A';
    palette.append(sw);
  }
  const box = popup.querySelector<HTMLElement>('.filter-values')!;
  for (const v of sorted) {
    const l = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = key(v);
    cb.checked = !current || [...current].some((x) => COLOR_TOKEN.test(x)) || current.has(key(v));
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
  popup.onkeydown = (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') popup.hidden = true;
    if (e.key === 'Enter') (popup.querySelector('[data-act="ok"]') as HTMLButtonElement).click();
  };
  popup.onclick = (e) => {
    const token = (e.target as HTMLElement).closest<HTMLElement>('.filter-swatch')?.dataset.token;
    if (token) {
      // Выбранный цвет — снова щелчок снимает; по цвету отбирается одно значение за раз.
      if (current?.has(token)) f.criteria.delete(column); else f.criteria.set(column, new Set([token]));
      popup.hidden = true;
      apply(f);
      return;
    }
    const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
    if (!act) return;
    if (act === 'asc' || act === 'desc') { popup.hidden = true; host.sort(column, act === 'desc', f); return; }
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
