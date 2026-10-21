import './styles.css';
import { allCommands, handleKey, keyLabel, matchScore, register, run } from './commands';
import { type BorderType, columnName, columnNumber, loadEngine, newWorkbook, nextSheetName, openXlsx, type Area, type Model } from './engine';
import { fontStack, loadFonts } from './fonts';
import { initFind, showFind } from './findreplace';
import { printSheet, renderPreview } from './print';
import { exampleFor, type FunctionExample } from './examples';
import { askExpand, columnTitle, customSort, type SortRequest } from './sortdialog';
import { filterOf, initFilter, openFilterPopup, resetFilters, setFilter } from './filter';
import { hideBackstage, initBackstage, showBackstage } from './backstage';
import { copyText, readText } from './platform';
import { Grid } from './grid';
import { icon, renderRibbon } from './ribbon';
import { FUNCTIONS } from './functions';
import { localeId, saveSettings, settings } from './settings';
import * as desk from './desktop';
import { isDesktop, isTauri } from './platform';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const LAST_ROW = 1_048_576;
const LAST_COLUMN = 16_384;

await Promise.all([loadEngine(), loadFonts()]);
let model: Model = newWorkbook(localeId());
model.setRecognizeDates(settings.recognizeDates);
let doc: desk.Doc = desk.newDoc();
window.__ensDoc = () => doc;

const grid = new Grid($('grid'), () => model);
// Для проверок в режиме разработки (в сборку не попадает).
if (import.meta.env.DEV) Object.assign(window, { __ens: { model: () => model, grid } });
const nameBox = $<HTMLInputElement>('name-box');
const formula = $<HTMLInputElement>('formula');
const editor = $<HTMLTextAreaElement>('cell-editor');
const status = $('status-stats');

// --- Редактирование ячейки ---

let editing = false;

function selectionArea(): Area {
  const s = grid.sel;
  return { sheet: grid.sheet, row: s.r1, column: s.c1, width: s.c2 - s.c1 + 1, height: s.r2 - s.r1 + 1 };
}

/** Поле ввода всегда стоит на активной ячейке. Пока не редактируем, оно
 *  невидимо и только принимает ввод: так работают любые способы набора текста
 *  (клавиатура в любой раскладке, IME, экранная клавиатура, вставка). */
function placeEditor(): void {
  const r = grid.activeRect();
  const host = grid.scroller.getBoundingClientRect();
  editor.style.left = `${host.left + r.x}px`;
  editor.style.top = `${host.top + r.y}px`;
  editor.style.width = `${r.w}px`;
  editor.style.height = `${r.h}px`;
  if (editing && settings.expandEditor) growEditor();
}

/** Настройка «Расширять поле ввода»: поле растёт вправо по мере набора.
 *  По умолчанию выключено — текст прокручивается внутри ячейки. */
const measure = document.createElement('canvas').getContext('2d')!;
function growEditor(): void {
  const r = grid.cellRect(grid.sel.row, grid.sel.column);
  const cs = getComputedStyle(editor);
  measure.font = `${cs.fontSize} ${cs.fontFamily}`;
  const need = measure.measureText(editor.value).width + 14;
  const max = grid.scroller.clientWidth - r.x - 4;
  editor.style.width = `${Math.min(Math.max(r.w, need), Math.max(r.w, max))}px`;
}

/** Режим ввода, как в Excel: «ввод» (начали печатать — стрелки завершают ввод)
 *  или «правка» (F2, двойной щелчок, строка формул — стрелки двигают курсор). */
let editMode: 'enter' | 'edit' = 'enter';

/** Ячейка, которую редактируют: при вводе формулы можно перейти на другой
 *  лист и выбрать ячейки там — результат запишется сюда. */
let editCell = { sheet: 0, row: 1, column: 1 };

function focusGrid(): void {
  placeEditor();
  editor.focus({ preventScroll: true });
}

function startEdit(initial?: string): void {
  editMode = initial === undefined ? 'edit' : 'enter';
  editCell = { sheet: grid.sheet, row: grid.sel.row, column: grid.sel.column };
  placeEditor();
  editor.value = initial ?? model.getCellContent(grid.sheet, grid.sel.row, grid.sel.column);
  formula.value = editor.value;
  editing = true;
  editor.classList.add('editing');
  if (settings.expandEditor) growEditor();
  editor.focus();
  editor.setSelectionRange(editor.value.length, editor.value.length);
}

function commitEdit(dRow: number, dCol: number): void {
  if (!editing) return;
  endPoint();
  const s = editCell;
  const value = activeValue();
  const before = model.getCellStyle(s.sheet, s.row, s.column).style.num_fmt;
  try {
    model.setUserInput(s.sheet, s.row, s.column, value);
  } catch (err) {
    showToast(`Не удалось записать значение: ${err}`);
  }
  if (value.startsWith('=') && model.getCellType(s.sheet, s.row, s.column) === 16
    && ['#ОШИБКА!', '#ERROR!'].includes(model.getFormattedCellValue(s.sheet, s.row, s.column))) {
    // Как Excel: недописанную формулу («=B5+») не записываем, ввод продолжается.
    model.undo();
    showToast('В формуле ошибка: она не дописана или в ней лишний символ. Исправьте её или нажмите Esc, чтобы отменить ввод.');
    const inp = editor.classList.contains('away') ? formula : editor;
    inp.focus();
    return;
  }
  editing = false;
  editor.classList.remove('editing', 'away');
  showRefs();
  const after = model.getCellStyle(s.sheet, s.row, s.column).style.num_fmt;
  editor.value = '';
  if (grid.sheet !== s.sheet) setSheet(s.sheet, false);
  fitRows(s.row, s.row);
  grid.select(s.row, s.column);
  if (after !== before) offerFormatUndo(s.row, s.column, before, after);
  grid.render();
  grid.move(dRow, dCol, false);
  focusGrid();
}

/** Текст ввода: при работе с другим листом он в строке формул. */
function activeValue(): string {
  return editor.classList.contains('away') ? formula.value : editor.value;
}

function cancelEdit(): void {
  endPoint();
  editing = false;
  editor.classList.remove('editing', 'away');
  showRefs();
  editor.value = '';
  if (grid.sheet !== editCell.sheet) setSheet(editCell.sheet, false);
  grid.select(editCell.row, editCell.column);
  updateFormulaBar();
  focusGrid();
}

editor.addEventListener('input', () => {
  if (!editing) {
    // Первый введённый символ начинает редактирование, как в Excel.
    editing = true;
    editMode = 'enter';
    editCell = { sheet: grid.sheet, row: grid.sel.row, column: grid.sel.column };
    editor.classList.add('editing');
  }
  endPoint();
  if (clip?.cut) dropClip();
  formula.value = editor.value;
  showRefs();
  if (settings.expandEditor) growEditor();
});

// --- Выбор ячеек для формулы стрелками и мышью (как в Excel) ---
// После «=», оператора, «(» или «;» стрелка или щелчок вставляют ссылку на
// ячейку; Shift+стрелка или протяжка мышью — диапазон. Следующий набранный
// символ закрепляет ссылку. Работает и в ячейке, и в строке формул.

interface Point { start: number; end: number; anchorRow: number; anchorCol: number; row: number; col: number }
let point: Point | null = null;

function activeInput(): HTMLInputElement | HTMLTextAreaElement {
  return document.activeElement === formula ? formula : editor;
}

function canPoint(inp: HTMLInputElement | HTMLTextAreaElement): boolean {
  const v = inp.value;
  if (!v.startsWith('=')) return false;
  const caret = inp.selectionStart ?? v.length;
  if (point && caret === point.end) return true;
  if (inp.selectionStart !== inp.selectionEnd) return false;
  return /[=+\-*/^&(;,:<>]\s*$/.test(v.slice(0, caret));
}

function pointText(p: Point): string {
  const r1 = Math.min(p.anchorRow, p.row), r2 = Math.max(p.anchorRow, p.row);
  const c1 = Math.min(p.anchorCol, p.col), c2 = Math.max(p.anchorCol, p.col);
  const a = `${columnName(c1)}${r1}`;
  const ref = r1 === r2 && c1 === c2 ? a : `${a}:${columnName(c2)}${r2}`;
  if (grid.sheet === editCell.sheet) return ref;
  const name = model.getWorksheetsProperties()[grid.sheet].name;
  return `${/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`}!${ref}`;
}

function applyPoint(inp: HTMLInputElement | HTMLTextAreaElement, row: number, col: number, extend: boolean): void {
  row = Math.max(1, row); col = Math.max(1, col);
  if (!point) {
    const caret = inp.selectionStart ?? inp.value.length;
    point = { start: caret, end: caret, anchorRow: row, anchorCol: col, row, col };
  } else if (extend) {
    point.row = row; point.col = col;
  } else {
    point.anchorRow = point.row = row; point.anchorCol = point.col = col;
  }
  const text = pointText(point);
  inp.value = inp.value.slice(0, point.start) + text + inp.value.slice(point.end);
  point.end = point.start + text.length;
  inp.setSelectionRange(point.end, point.end);
  (inp === editor ? formula : editor).value = inp.value;
  showRefs();
  grid.pointRange = {
    r1: Math.min(point.anchorRow, point.row), c1: Math.min(point.anchorCol, point.col),
    r2: Math.max(point.anchorRow, point.row), c2: Math.max(point.anchorCol, point.col),
  };
  grid.ensureVisible(row, col);
  grid.render();
}

function endPoint(): void {
  if (!point) return;
  point = null;
  grid.pointRange = null;
  grid.render();
}

const ARROWS: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };

/** Клавиши во время редактирования (в ячейке или в строке формул). */
function editKey(e: KeyboardEvent): void {
  const inp = activeInput();
  const d = ARROWS[e.key];
  if (d && !e.ctrlKey && !e.altKey && canPoint(inp)) {
    e.preventDefault();
    const from = point ? { row: point.row, col: point.col }
      : grid.sheet === editCell.sheet ? { row: editCell.row, col: editCell.column } : { row: grid.sel.row, col: grid.sel.column };
    applyPoint(inp, from.row + d[0], from.col + d[1], e.shiftKey && point !== null);
    return;
  }
  if (!['Shift', 'Control', 'Alt', 'Meta'].includes(e.key)) endPoint();
  if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); commitEdit(e.shiftKey ? -1 : 1, 0); }
  else if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); }
  else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
  else if (e.key === 'F2') { e.preventDefault(); editMode = editMode === 'enter' ? 'edit' : 'enter'; }
  else if (d && editMode === 'enter' && inp === editor) {
    // Режим «ввод»: стрелка завершает ввод и переходит к соседней ячейке.
    e.preventDefault();
    commitEdit(d[0], d[1]);
  }
}

grid.onPointer = (row, column, phase, shift) => {
  if (!editing) return false;
  const inp = activeInput();
  if (phase === 'down') {
    if (!canPoint(inp)) return false;
    applyPoint(inp, row, column, shift && point !== null);
    return true;
  }
  if (!point) return false;
  applyPoint(inp, row, column, true);
  return true;
};

editor.addEventListener('keydown', (e) => {
  if (!editing) { gridKey(e); return; }
  editKey(e);
});
// Поле ввода потеряло фокус: щелчок по ленте, другому полю — ввод завершается.
// Не завершаем, если фокус ушёл в строку формул (или ввод на другом листе) и
// если переключились в другое окно: в Excel ввод при этом тоже продолжается.
// Решаем после того, как фокус «осел»: во время blur в WebView2 (Windows)
// activeElement — ещё <body>, и раньше формула записывалась недописанной.
editor.addEventListener('blur', () => setTimeout(() => {
  const now = document.activeElement;
  if (!editing || editor.classList.contains('away') || now === formula || now === editor || !document.hasFocus()) return;
  if (now && (now.closest('#recovery') || now.closest('#context-menu'))) return;
  commitEdit(0, 0);
}, 0));

formula.addEventListener('focus', () => { if (!editing) { startEdit(); formula.focus(); } showRefs(); });
formula.addEventListener('input', () => { endPoint(); editor.value = formula.value; showRefs(); });
formula.addEventListener('scroll', () => syncMirror());
formula.addEventListener('blur', () => setTimeout(() => {
  const now = document.activeElement;
  if (!editing || now === formula || now === editor || !document.hasFocus()) return;
  if (now && (now.closest('#recovery') || now.closest('#context-menu') || now.closest('.sheet-tabs'))) return;
  commitEdit(0, 0);
}, 0));
formula.addEventListener('keydown', (e) => editKey(e));

// §3: Ctrl+щелчок по ссылке в строке формул — переход к ячейке.
const REF = /(?:'([^']+)'|([A-Za-zА-Яа-я_][\wА-Яа-я.]*))?!?\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?/g;

interface Ref { sheet?: string; row: number; column: number; row2?: number; column2?: number; text: string; start: number; end: number }

function refsIn(text: string): Ref[] {
  const out: Ref[] = [];
  // Строчные буквы («=b5») — тоже ссылки; длина строки при этом не меняется.
  for (const m of text.toUpperCase().matchAll(REF)) {
    const after = text[m.index! + m[0].length];
    if (after === '(' || (after && /[\wА-Яа-я]/.test(after))) continue; // LOG10( — функция, не ячейка
    const sheet = m[1] ?? (m[0].includes('!') ? m[2] : undefined);
    out.push({
      sheet, column: columnNumber(m[3]), row: Number(m[4]),
      column2: m[5] ? columnNumber(m[5]) : undefined, row2: m[6] ? Number(m[6]) : undefined,
      text: m[0], start: m.index!, end: m.index! + m[0].length,
    });
  }
  return out;
}

const backStack: { sheet: number; row: number; column: number }[] = [];

function goTo(ref: Ref): void {
  backStack.push({ sheet: grid.sheet, row: grid.sel.row, column: grid.sel.column });
  if (ref.sheet) {
    const index = model.getWorksheetsProperties().findIndex((p) => p.name.toLowerCase() === ref.sheet!.toLowerCase());
    if (index >= 0) setSheet(index);
  }
  if (ref.row2 !== undefined && ref.column2 !== undefined) {
    // Диапазон (C9:C12) выделяется целиком.
    grid.selectRange(Math.min(ref.row, ref.row2), Math.min(ref.column, ref.column2),
      Math.max(ref.row, ref.row2), Math.max(ref.column, ref.column2));
  } else {
    grid.select(ref.row, ref.column);
  }
  focusGrid();
}

formula.addEventListener('click', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const at = formula.selectionStart ?? 0;
  const ref = refsIn(formula.value).find((r) => at >= r.start && at <= r.end);
  if (ref) { e.preventDefault(); cancelEdit(); goTo(ref); }
});

// --- Цветные ссылки формулы: ячейки на листе и текст в строке формул ---

const REF_COLORS = ['#2f6fd1', '#c0392b', '#7b3fb5', '#1e8449', '#d35400', '#b8860b', '#c2185b', '#00838f'];
const mirror = $('formula-mirror');

function showRefs(): void {
  const value = editing ? activeValue() : '';
  const refs = value.startsWith('=') ? refsIn(value) : [];
  const names = model.getWorksheetsProperties().map((p) => p.name.toLowerCase());
  const colorOf = new Map<string, string>();
  const colored = refs.map((r) => {
    const key = r.text.toUpperCase().replace(/\$/g, '');
    if (!colorOf.has(key)) colorOf.set(key, REF_COLORS[colorOf.size % REF_COLORS.length]);
    return { ...r, color: colorOf.get(key)! };
  });
  grid.refHighlights = colored
    .filter((r) => (r.sheet ? names.indexOf(r.sheet.toLowerCase()) : editCell.sheet) === grid.sheet)
    .map((r) => ({
      r1: Math.min(r.row, r.row2 ?? r.row), r2: Math.max(r.row, r.row2 ?? r.row),
      c1: Math.min(r.column, r.column2 ?? r.column), c2: Math.max(r.column, r.column2 ?? r.column), color: r.color,
    }));
  grid.render();
  // Строка формул: под прозрачным текстом поля — тот же текст с цветными ссылками.
  const show = colored.length > 0 && document.activeElement === formula;
  formula.classList.toggle('colored', show);
  mirror.hidden = !show;
  if (!show) return;
  let html = '', at = 0;
  for (const r of colored) {
    html += escapeHtml(value.slice(at, r.start)) + `<span style="color:${r.color}">${escapeHtml(value.slice(r.start, r.end))}</span>`;
    at = r.end;
  }
  mirror.innerHTML = html + escapeHtml(value.slice(at));
  syncMirror();
}

function syncMirror(): void {
  mirror.style.transform = `translateX(${-formula.scrollLeft}px)`;
}

const escapeHtml = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

// --- §4: формат, изменённый вводом, можно сразу вернуть ---

const chip = $('format-chip');
let chipTimer = 0;

function formatName(fmt: string): string {
  if (/[dmy]/i.test(fmt) && !/[#0]/.test(fmt)) return 'Дата';
  if (fmt.includes('%')) return 'Процентный';
  if (/[$€₽]/.test(fmt)) return 'Денежный';
  if (fmt === '@') return 'Текстовый';
  return fmt === 'general' ? 'Общий' : 'Числовой';
}

function offerFormatUndo(row: number, column: number, before: string, after: string): void {
  const sheet = grid.sheet;
  const r = grid.cellRect(row, column);
  const host = grid.scroller.getBoundingClientRect();
  chip.innerHTML = `Формат изменён на «${formatName(after)}» <button>Вернуть «${formatName(before)}»</button>`;
  chip.style.left = `${host.left + r.x}px`;
  chip.style.top = `${host.top + r.y + r.h + 4}px`;
  chip.hidden = false;
  chip.querySelector('button')!.onclick = () => {
    model.updateRangeStyle({ sheet, row, column, width: 1, height: 1 }, 'num_fmt', before);
    chip.hidden = true;
    grid.render();
    focusGrid();
  };
  clearTimeout(chipTimer);
  chipTimer = window.setTimeout(() => (chip.hidden = true), 8000);
}

// --- Ctrl+стрелки: к краю области данных, как в Excel ---

function dataEdge(row: number, column: number, dRow: number, dCol: number): { row: number; column: number } {
  const vertical = dRow !== 0;
  const dir = vertical ? dRow : dCol;
  const pos = vertical ? row : column;
  const limit = dir > 0 ? (vertical ? LAST_ROW : LAST_COLUMN) : 1;
  const filled = new Set(vertical ? model.getRowsWithData(grid.sheet, column) : model.getColumnsWithData(grid.sheet, row));
  const has = (p: number) => filled.has(p);
  let target: number;
  if (has(pos) && has(pos + dir)) {
    // Внутри блока данных — к его концу.
    target = pos;
    while (target !== limit && has(target + dir)) target += dir;
  } else {
    // К ближайшей заполненной ячейке, иначе к краю листа.
    const next = [...filled].filter((p) => (dir > 0 ? p > pos : p < pos));
    target = next.length ? (dir > 0 ? Math.min(...next) : Math.max(...next)) : limit;
  }
  return vertical ? { row: target, column } : { row, column: target };
}

// --- Поле имени: ввод адреса (B12, Лист2!A1) ---

nameBox.addEventListener('keydown', (e) => {
  if (e.key !== 'Enter') return;
  e.preventDefault();
  const ref = refsIn(nameBox.value.toUpperCase())[0];
  if (ref) goTo(ref);
});

// --- Строка формул и строка состояния следуют за выделением ---

function updateFormulaBar(): void {
  const s = grid.sel;
  nameBox.value = `${columnName(s.column)}${s.row}`;
  if (!editing) formula.value = model.getCellContent(grid.sheet, s.row, s.column);
}

/** Число из отображаемого значения с учётом разделителей локали. */
function parseShown(text: string): number {
  let t = text.replace(/[\s ]/g, '');
  t = settings.decimal === 'comma' ? t.replace(/\./g, '').replace(',', '.') : t.replace(/,/g, '');
  return Number(t.replace(/[^\d.\-E]/g, ''));
}

function updateStatus(): void {
  const s = grid.sel;
  let count = 0, numbers = 0, sum = 0;
  const cells = (s.r2 - s.r1 + 1) * (s.c2 - s.c1 + 1);
  if (cells > 1 && cells <= 200_000) {
    for (let r = s.r1; r <= s.r2; r++) {
      for (let c = s.c1; c <= s.c2; c++) {
        const text = model.getFormattedCellValue(grid.sheet, r, c);
        if (!text) continue;
        count++;
        if (model.getCellType(grid.sheet, r, c) === 1) {
          const v = parseShown(text);
          if (Number.isFinite(v)) { numbers++; sum += v; }
        }
      }
    }
  }
  const fmt = (v: number) => v.toLocaleString('ru-RU', { maximumFractionDigits: 10 });
  const f = filterOf(grid.sheet);
  if (f && f.criteria.size && count <= 1) {
    let shown = 0;
    for (let r = f.r1 + 1; r <= f.r2; r++) if (model.getRowHeight(grid.sheet, r) > 0) shown++;
    status.innerHTML = `<span class="stat cnt">Найдено записей: ${shown} из ${f.r2 - f.r1}</span>`;
    return;
  }
  status.innerHTML = count > 1
    ? `<span class="stat avg">Среднее: ${numbers ? fmt(sum / numbers) : '—'}</span>` +
      `<span class="stat cnt">Количество: ${count}</span>` +
      `<span class="stat sum">Сумма: ${fmt(sum)}</span>`
    : '';
}

grid.onEditRequest = (initial) => startEdit(initial);

// --- Клавиатура сетки ---

function gridKey(e: KeyboardEvent): void {
  if (handleKey(e)) return;
  const shift = e.shiftKey;
  const ctrl = e.ctrlKey || e.metaKey;
  if (ctrl && e.key.startsWith('Arrow')) {
    e.preventDefault();
    const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key]!;
    const from = shift ? grid.corner : { row: grid.sel.row, column: grid.sel.column };
    const to = dataEdge(from.row, from.column, d[0], d[1]);
    grid.select(to.row, to.column, shift);
    return;
  }
  switch (e.key) {
    case 'ArrowUp': e.preventDefault(); grid.move(-1, 0, shift); return;
    case 'ArrowDown': e.preventDefault(); grid.move(1, 0, shift); return;
    case 'ArrowLeft': e.preventDefault(); grid.move(0, -1, shift); return;
    case 'ArrowRight': e.preventDefault(); grid.move(0, 1, shift); return;
    case 'PageDown': e.preventDefault(); grid.move(grid.pageRows(), 0, shift); return;
    case 'PageUp': e.preventDefault(); grid.move(-grid.pageRows(), 0, shift); return;
    case 'Home': e.preventDefault(); ctrl ? grid.select(1, 1) : grid.select(grid.sel.row, 1); return;
    case 'Enter': e.preventDefault(); grid.move(shift ? -1 : 1, 0, false); return;
    case 'Tab': e.preventDefault(); grid.move(0, shift ? -1 : 1, false); return;
    case 'F2': e.preventDefault(); startEdit(); return;
    case 'Escape': if (clip) { e.preventDefault(); dropClip(); } return;
  }
  // Печатные символы не перехватываем: они попадают в поле ввода
  // и начинают редактирование через событие input.
}

// --- Листы ---

function setSheet(index: number, reset = true): void {
  grid.sheet = index;
  model.setSelectedSheet(index);
  grid.invalidateLayout();
  if (reset) grid.select(1, 1);
  renderSheetTabs();
  showRefs();
  syncFilter();
}

function syncFilter(): void {
  const f = filterOf(grid.sheet);
  grid.filter = f ? { r1: f.r1, c1: f.c1, r2: f.r2, c2: f.c2, active: new Set(f.criteria.keys()) } : null;
  grid.invalidateLayout();
}

/** Щелчок по ярлычку во время ввода формулы — выбор ячеек на другом листе. */
function clickSheetTab(index: number): void {
  if (index === grid.sheet) return;
  if (editing && activeValue().startsWith('=')) {
    endPoint();
    const value = activeValue();
    editor.classList.add('away');
    formula.value = value;
    setSheet(index, false);
    formula.focus();
    formula.setSelectionRange(value.length, value.length);
    return;
  }
  if (editing) commitEdit(0, 0);
  setSheet(index);
}

function renderSheetTabs(): void {
  const host = $('sheet-tabs');
  host.replaceChildren();
  model.getWorksheetsProperties().forEach((p, i) => {
    if (p.state !== 'visible') return;
    const b = document.createElement('button');
    b.className = `sheet-tab${i === grid.sheet ? ' active' : ''}`;
    b.textContent = p.name;
    if (p.color) b.style.borderBottom = `3px solid ${p.color}`;
    b.addEventListener('mousedown', (e) => { if (editing) e.preventDefault(); });
    b.addEventListener('click', () => clickSheetTab(i));
    b.addEventListener('dblclick', () => renameSheet(i, b));
    b.addEventListener('contextmenu', (e) => { e.preventDefault(); sheetMenu(i, e.clientX, e.clientY); });
    host.append(b);
  });
  const add = document.createElement('button');
  add.className = 'sheet-add';
  add.title = 'Новый лист (Shift+F11)';
  add.textContent = '+';
  add.addEventListener('click', () => run('sheet.new'));
  host.append(add);
}

function renameSheet(index: number, tab?: HTMLElement): void {
  const tabEl = tab ?? [...document.querySelectorAll<HTMLElement>('.sheet-tab')].find((t) => t.textContent === model.getWorksheetsProperties()[index].name);
  if (!tabEl) return;
  const old = model.getWorksheetsProperties()[index].name;
  const input = document.createElement('input');
  input.className = 'sheet-rename';
  input.value = old;
  tabEl.replaceWith(input);
  input.focus();
  input.select();
  let done = false;
  const finish = (save: boolean) => {
    if (done) return;
    done = true;
    const name = input.value.trim();
    if (save && name && name !== old) {
      try { model.renameSheet(index, name); } catch (e) { showToast(`Нельзя так назвать лист: ${e}`); }
    }
    renderSheetTabs();
    focusGrid();
  };
  input.addEventListener('keydown', (e) => { if (e.key === 'Enter') finish(true); if (e.key === 'Escape') finish(false); });
  input.addEventListener('blur', () => finish(true));
}

const SHEET_COLORS = ['', '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#7030A0'];

function sheetMenu(index: number, x: number, y: number): void {
  const props = model.getWorksheetsProperties();
  const visible = props.filter((p) => p.state === 'visible').length;
  const hidden = props.map((p, i) => ({ p, i })).filter(({ p }) => p.state !== 'visible');
  const items: { label: string; run: () => void; disabled?: boolean }[] = [
    { label: 'Вставить лист', run: () => run('sheet.new') },
    { label: 'Удалить', disabled: visible <= 1, run: () => {
      if (!confirm(`Удалить лист «${props[index].name}»? Отменить удаление можно сочетанием Ctrl+Z.`)) return;
      model.deleteSheet(index); setSheet(Math.max(0, Math.min(index, model.getWorksheetsProperties().length - 1)));
    } },
    { label: 'Переименовать', run: () => renameSheet(index) },
    { label: 'Копировать лист', run: () => { model.duplicateSheet(index); setSheet(index + 1); } },
    { label: 'Переместить влево', disabled: index === 0, run: () => { model.moveSheet(index, index - 1); setSheet(index - 1, false); } },
    { label: 'Переместить вправо', disabled: index >= props.length - 1, run: () => { model.moveSheet(index, index + 1); setSheet(index + 1, false); } },
    { label: 'Скрыть', disabled: visible <= 1, run: () => {
      model.hideSheet(index);
      const next = model.getWorksheetsProperties().findIndex((p) => p.state === 'visible');
      setSheet(next);
    } },
    ...hidden.map(({ p, i }) => ({ label: `Показать «${p.name}»`, run: () => { model.unhideSheet(i); setSheet(i); } })),
  ];
  const menu = showMenu(items, x, y);
  const colors = document.createElement('div');
  colors.className = 'menu-colors';
  colors.title = 'Цвет ярлычка';
  for (const c of SHEET_COLORS) {
    const s = document.createElement('button');
    s.className = 'swatch';
    s.style.background = c || '#ffffff';
    s.title = c ? c : 'Без цвета';
    s.addEventListener('click', () => { menu.hidden = true; model.setSheetColor(index, c); renderSheetTabs(); });
    colors.append(s);
  }
  const label = document.createElement('div');
  label.className = 'menu-label';
  label.textContent = 'Цвет ярлычка';
  menu.append(label, colors);
  placeMenu(menu, x, y, true);
}

interface MenuItem { label: string; run: () => void; disabled?: boolean; keys?: string; separator?: boolean }

/** Контекстное меню: пункты; separator — черта перед пунктом. */
function showMenu(items: MenuItem[], x: number, y: number): HTMLElement {
  const menu = $('context-menu');
  menu.replaceChildren();
  for (const it of items) {
    if (it.separator) menu.append(Object.assign(document.createElement('hr'), { className: 'menu-sep' }));
    const b = document.createElement('button');
    b.innerHTML = `<span>${escapeHtml(it.label)}</span>${it.keys ? `<small>${escapeHtml(it.keys)}</small>` : ''}`;
    b.disabled = !!it.disabled;
    b.addEventListener('click', () => { menu.hidden = true; it.run(); });
    menu.append(b);
  }
  placeMenu(menu, x, y, false);
  return menu;
}

/** Меню целиком в окне: у нижнего края — вверх от курсора. */
function placeMenu(menu: HTMLElement, x: number, y: number, above: boolean): void {
  menu.hidden = false;
  const { width, height } = menu.getBoundingClientRect();
  const top = above || y + height > window.innerHeight - 8 ? y - height : y;
  menu.style.left = `${Math.max(8, Math.min(x, window.innerWidth - width - 8))}px`;
  menu.style.top = `${Math.max(8, top)}px`;
}

const keysOf = (id: string) => {
  const c = allCommands().find((x) => x.id === id);
  return c?.keys?.length ? keyLabel(c.keys[0]) : undefined;
};

/** Меню ячейки (правый щелчок), как в Excel. */
function cellMenu(x: number, y: number, header: 'row' | 'column' | 'corner' | null): void {
  const s = grid.sel;
  const rows = s.r2 - s.r1 + 1, cols = s.c2 - s.c1 + 1;
  const fullRows = s.c1 === 1 && s.c2 === LAST_COLUMN;
  const fullCols = s.r1 === 1 && s.r2 === LAST_ROW;
  const items: MenuItem[] = [
    { label: 'Вырезать', keys: keysOf('clipboard.cut'), run: () => run('clipboard.cut') },
    { label: 'Копировать', keys: keysOf('clipboard.copy'), run: () => run('clipboard.copy') },
    { label: 'Вставить', keys: keysOf('clipboard.paste'), run: () => run('clipboard.paste') },
    { label: 'Специальная вставка…', keys: 'Ctrl+Alt+V', run: () => pasteSpecialMenu(x, y) },
  ];
  if (!header) {
    items.push({ label: 'Сортировка от А до Я', separator: true, run: () => sortBy(s.column, false) });
    items.push({ label: 'Сортировка от Я до А', run: () => sortBy(s.column, true) });
    items.push({ label: 'Настраиваемая сортировка…', run: () => sortBy(s.column, false, true) });
    items.push({ label: filterOf(grid.sheet) ? 'Снять фильтр' : 'Фильтр', run: toggleFilter });
  }
  if (header !== 'column' && !fullCols) {
    items.push({ label: rows > 1 ? `Вставить строки (${rows})` : 'Вставить строку', separator: true, run: () => structure(() => model.insertRows(grid.sheet, s.r1, rows)) });
  }
  if (header !== 'row' && !fullRows) {
    items.push({ label: cols > 1 ? `Вставить столбцы (${cols})` : 'Вставить столбец', separator: header !== 'column' && !fullCols ? false : true,
      run: () => structure(() => model.insertColumns(grid.sheet, s.c1, cols)) });
  }
  if (header !== 'column' && !fullCols) {
    items.push({ label: rows > 1 ? `Удалить строки (${rows})` : 'Удалить строку', run: () => structure(() => model.deleteRows(grid.sheet, s.r1, rows)) });
  }
  if (header !== 'row' && !fullRows) {
    items.push({ label: cols > 1 ? `Удалить столбцы (${cols})` : 'Удалить столбец', run: () => structure(() => model.deleteColumns(grid.sheet, s.c1, cols)) });
  }
  items.push({ label: 'Очистить содержимое', keys: keysOf('edit.clear'), run: () => run('edit.clear') });
  if (!header && (rows > 1 || cols > 1)) {
    items.push({ label: 'Объединить и поместить в центр', separator: true, run: () => merge('center') });
    items.push({ label: 'Объединить ячейки', run: () => merge('plain') });
    if (rows > 1) items.push({ label: 'Объединить по строкам', run: () => merge('across') });
  }
  if (!header && model.getMergedCells(grid.sheet).some((m) => m.row <= s.r2 && m.row + m.height - 1 >= s.r1 && m.column <= s.c2 && m.column + m.width - 1 >= s.c1)) {
    items.push({ label: 'Отменить объединение ячеек', run: () => merge('undo') });
  }
  if (header !== 'row') items.push({ label: 'Ширина столбца…', separator: true, run: () => askSize('column') });
  if (header !== 'column') items.push({ label: 'Высота строки…', separator: header === 'row', run: () => askSize('row') });
  if (!header) items.push({ label: 'Перейти к влияющим ячейкам', separator: true, keys: keysOf('nav.precedents'), run: () => run('nav.precedents') });
  showMenu(items, x, y);
}

/** Вставка/удаление строк и столбцов: формулы пересчитываются, сетка перерисовывается. */
function structure(change: () => void): void {
  try { change(); } catch (e) { showToast(`Не получилось: ${e}`); }
  grid.invalidateLayout();
  updateFormulaBar();
  updateStatus();
}

/** «Ширина столбца…» / «Высота строки…» — в пикселях, для всех выделенных. */
function askSize(kind: 'column' | 'row'): void {
  const s = grid.sel;
  const current = kind === 'column' ? Math.round(model.getColumnWidth(grid.sheet, s.column))
    : Math.round(model.getRowHeight(grid.sheet, s.row) * 22 / 25);
  const answer = prompt(kind === 'column' ? 'Ширина столбца (в пикселях):' : 'Высота строки (в пикселях):', String(current));
  const px = Number(answer?.replace(',', '.'));
  if (!answer || !Number.isFinite(px) || px <= 0 || px > 2000) return;
  setSize(kind, kind === 'column' ? s.c1 : s.r1, kind === 'column' ? s.c2 : s.r2, px);
  focusGrid();
}

function setSize(kind: 'column' | 'row', from: number, to: number, px: number): void {
  if (kind === 'column') model.setColumnsWidth(grid.sheet, from, Math.min(to, LAST_COLUMN), px);
  else model.setRowsHeight(grid.sheet, from, Math.min(to, LAST_ROW), Grid.rowPxToModel(px));
  grid.invalidateLayout();
  placeEditor();
}

/** Ширина по содержимому (двойной щелчок по границе), как в Excel. */
function fitColumn(column: number): number {
  let w = 0;
  for (const r of model.getRowsWithData(grid.sheet, column)) {
    const text = model.getFormattedCellValue(grid.sheet, r, column);
    if (!text) continue;
    const f = model.getCellStyle(grid.sheet, r, column).style.font;
    measure.font = `${f.i ? 'italic ' : ''}${f.b ? 'bold ' : ''}${Math.round((f.sz ?? 11) * 1.33)}px ${fontStack(f.name)}`;
    w = Math.max(w, measure.measureText(text).width);
  }
  return w ? Math.ceil(w + 10) : 90;
}

function fitRow(row: number): number {
  let need = 22;
  for (const c of model.getColumnsWithData(grid.sheet, row)) {
    const st = model.getCellStyle(grid.sheet, row, c).style;
    const text = model.getFormattedCellValue(grid.sheet, row, c);
    need = Math.max(need, st.alignment?.wrap_text && text
      ? grid.measureWrapped(text, st, grid.columnPx(c))
      : Math.round((st.font.sz ?? 11) * 1.33 + 8));
  }
  return need;
}

/** Высота строк с переносом текста — по содержимому, как в Excel (только увеличиваем). */
function fitRows(r1: number, r2: number): void {
  if (r2 - r1 > 500) return;
  for (let r = r1; r <= r2; r++) {
    const wrapped = [...model.getColumnsWithData(grid.sheet, r)].some((c) => model.getCellStyle(grid.sheet, r, c).style.alignment?.wrap_text);
    if (!wrapped) continue;
    const need = fitRow(r);
    const now = Math.round(model.getRowHeight(grid.sheet, r) * 22 / 25);
    if (need !== now) setSize('row', r, r, need);
  }
}

/** Объединение ячеек, как в Excel: если данных несколько — останется левое верхнее значение. */
function merge(kind: 'center' | 'plain' | 'across' | 'undo'): void {
  const area = selectionArea();
  const s = grid.sel;
  const inside = model.getMergedCells(grid.sheet).filter((m) => m.row <= s.r2 && m.row + m.height - 1 >= s.r1 && m.column <= s.c2 && m.column + m.width - 1 >= s.c1);
  try {
    if (kind === 'undo' || (kind === 'center' && inside.length)) {
      model.unmergeCells(area);
    } else {
      const filled: [number, number][] = [];
      for (let r = s.r1; r <= Math.min(s.r2, s.r1 + 2000); r++) {
        for (const c of model.getColumnsWithData(grid.sheet, r)) if (c >= s.c1 && c <= s.c2) filled.push([r, c]);
      }
      const perRow = kind === 'across';
      const extra = perRow
        ? filled.filter(([r, c]) => filled.some(([r2, c2]) => r2 === r && c2 < c))
        : filled.slice(1);
      if (extra.length) {
        if (!confirm('В объединённой ячейке сохранится только значение из левой верхней заполненной ячейки, остальные значения будут удалены. Продолжить?')) return;
        for (const [r, c] of extra) model.rangeClearContents(grid.sheet, r, c, r, c);
      }
      if (kind === 'center') model.mergeCellsCenter(area);
      else if (kind === 'across') model.mergeCellsAcross(area);
      else model.mergeCells(area);
    }
  } catch (e) {
    showToast(`Не получилось: ${e}`);
  }
  grid.invalidateLayout();
  updateFormulaBar();
  focusGrid();
}

/** Кнопка «Границы»: меню, как в Excel. */
const BORDER_TYPES: [BorderType, string][] = [
  ['All', 'Все границы'], ['Outer', 'Внешние границы'], ['Inner', 'Внутренние границы'],
  ['Bottom', 'Нижняя граница'], ['Top', 'Верхняя граница'], ['Left', 'Левая граница'], ['Right', 'Правая граница'],
  ['None', 'Нет границы'],
];
const BORDER_STYLES: [string, string][] = [
  ['thin', 'тонкая'], ['medium', 'средняя'], ['thick', 'толстая'], ['double', 'двойная'], ['dashed', 'пунктир'], ['dotted', 'точки'],
];

/** Поставить границы: по умолчанию — последний выбор (сначала «Все границы», тонкая). */
function applyBorders(type = settings.borderType as BorderType, st = settings.borderStyle): void {
  const style = st === 'dashed' ? 'mediumdashed' : st;
  try { model.setAreaWithBorder(selectionArea(), { item: { style, color: '#000000' }, type }); } catch (e) { showToast(String(e)); }
  grid.render();
  focusGrid();
}

/** Стрелка у кнопки «Границы»: какие границы и какой линией. Выбор запоминается. */
function bordersMenu(x: number, y: number): void {
  showMenu([
    ...BORDER_TYPES.map(([type, label]) => ({
      label: `${type === settings.borderType ? '● ' : ''}${label}`, separator: type === 'Bottom' || type === 'None',
      run: () => { settings.borderType = type; saveSettings(); applyBorders(); },
    })),
    ...BORDER_STYLES.map(([st, label], i) => ({
      label: `${st === settings.borderStyle ? '● ' : ''}Линия: ${label}`, separator: i === 0,
      run: () => { settings.borderStyle = st; saveSettings(); showToast(`Линия границ: ${label}. Щелчок по кнопке «Границы» ставит «${BORDER_TYPES.find((b) => b[0] === settings.borderType)?.[1]}» этой линией.`); },
    })),
  ], x, y);
}

grid.onResize = (kind, index, px) => {
  const s = grid.sel;
  // Тянут границу одного из выделенных целиком столбцов (строк) — меняются все.
  const many = kind === 'column' ? s.r1 === 1 && s.r2 === LAST_ROW && index >= s.c1 && index <= s.c2
    : s.c1 === 1 && s.c2 === LAST_COLUMN && index >= s.r1 && index <= s.r2;
  const [from, to] = many ? (kind === 'column' ? [s.c1, s.c2] : [s.r1, s.r2]) : [index, index];
  if (px !== null) { setSize(kind, from, to, px); return; }
  for (let i = from; i <= Math.min(to, from + 500); i++) setSize(kind, i, i, kind === 'column' ? fitColumn(i) : fitRow(i));
};

// Маркер заполнения: формулы протягиваются со сдвигом ссылок, ряды
// продолжаются (1, 2 → 3, 4…; январь → февраль…), как в Excel.
grid.onFill = (t) => {
  const s = grid.sel;
  const area = selectionArea();
  try {
    if (t.r1 !== s.r1 || t.r2 !== s.r2) model.autoFillRows(area, t.r1 < s.r1 ? t.r1 : t.r2);
    else model.autoFillColumns(area, t.c1 < s.c1 ? t.c1 : t.c2);
  } catch (e) {
    showToast(`Не получилось протянуть: ${e}`);
  }
  grid.selectRange(t.r1, t.c1, t.r2, t.c2);
  updateFormulaBar();
  updateStatus();
};

/** Закрепить области: строки выше и столбцы левее активной ячейки. */
function freezeMenu(x: number, y: number): void {
  const s = grid.sel, f = grid.frozen;
  showMenu([
    { label: f.rows || f.columns ? 'Снять закрепление областей' : `Закрепить области (выше и левее ${columnName(s.column)}${s.row})`,
      run: () => (f.rows || f.columns ? grid.setFrozen(0, 0) : grid.setFrozen(s.row - 1, s.column - 1)) },
    { label: 'Закрепить верхнюю строку', run: () => grid.setFrozen(1, 0) },
    { label: 'Закрепить первый столбец', run: () => grid.setFrozen(0, 1) },
  ], x, y);
}

let lastPointer = { x: 200, y: 160 };
document.addEventListener('pointerdown', (e) => { lastPointer = { x: e.clientX, y: e.clientY }; }, true);

// --- Отменить / Повторить: видно, можно ли, и что сработало ---

const qaUndo = $<HTMLButtonElement>('qa-undo');
const qaRedo = $<HTMLButtonElement>('qa-redo');
qaUndo.innerHTML = icon('arrow_undo', 20);
qaRedo.innerHTML = icon('arrow_redo', 20);
for (const [b, kind] of [[qaUndo, 'undo'], [qaRedo, 'redo']] as const) {
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', () => undoRedo(kind));
}

/** Значок «подмигивает», когда действие отменено или повторено. */
function wink(kind: 'undo' | 'redo'): void {
  const els = [kind === 'undo' ? qaUndo : qaRedo, ...document.querySelectorAll<HTMLElement>(`.btn[data-command="edit.${kind}"]`)];
  for (const el of els) {
    el.classList.remove('wink');
    void el.offsetWidth; // перезапуск анимации
    el.classList.add('wink');
  }
}

function undoRedo(kind: 'undo' | 'redo'): void {
  const can = kind === 'undo' ? model.canUndo() : model.canRedo();
  if (!can) { showToast(kind === 'undo' ? 'Отменять нечего.' : 'Повторять нечего.'); return; }
  if (kind === 'undo') model.undo(); else model.redo();
  wink(kind);
  grid.invalidateLayout();
  updateFormulaBar();
  updateStatus();
  syncUndoState();
}

/** Нефритовые — когда можно, серые — когда нельзя. */
function syncUndoState(): void {
  const u = model.canUndo(), r = model.canRedo();
  qaUndo.disabled = !u;
  qaRedo.disabled = !r;
  document.querySelectorAll<HTMLElement>('.btn[data-command="edit.undo"]').forEach((b) => b.classList.toggle('disabled', !u));
  document.querySelectorAll<HTMLElement>('.btn[data-command="edit.redo"]').forEach((b) => b.classList.toggle('disabled', !r));
}
setInterval(syncUndoState, 400);

// --- Данные: сортировка, фильтр, поиск ---

/** «Текущая область», как Ctrl+A в Excel: сплошной блок данных вокруг ячейки. */
type Rect = { r1: number; c1: number; r2: number; c2: number };

function currentRegion(row: number, column: number, start?: Rect): Rect {
  let { r1, c1, r2, c2 } = start ?? { r1: row, c1: column, r2: row, c2: column };
  // Заполненность берём списками строк по столбцам (один вызов движка на столбец),
  // а не вопросом о каждой ячейке: быстро и на больших таблицах.
  const byColumn = new Map<number, Set<number>>();
  const filled = (r: number, c: number) => {
    if (r < 1 || c < 1) return false;
    let rows = byColumn.get(c);
    if (!rows) { rows = new Set(model.getRowsWithData(grid.sheet, c)); byColumn.set(c, rows); }
    return rows.has(r);
  };
  for (let grow = true, guard = 0; grow && guard < 10_000; guard++) {
    grow = false;
    for (let c = c1 - 1; c <= c2 + 1; c++) {
      if (filled(r1 - 1, c)) { r1--; grow = true; break; }
      if (filled(r2 + 1, c)) { r2++; grow = true; break; }
    }
    for (let r = r1 - 1; r <= r2 + 1; r++) {
      if (filled(r, c1 - 1)) { c1--; grow = true; break; }
      if (filled(r, c2 + 1)) { c2++; grow = true; break; }
    }
  }
  return { r1, c1, r2, c2 };
}

/** Похоже ли, что первая строка — заголовки (текст над числами), как угадывает Excel. */
function hasHeader(r: { r1: number; c1: number; r2: number; c2: number }): boolean {
  if (r.r2 <= r.r1) return false;
  let text = true, numberBelow = false;
  for (let c = r.c1; c <= r.c2; c++) {
    const t = model.getCellType(grid.sheet, r.r1, c);
    if (t !== 2 && model.getFormattedCellValue(grid.sheet, r.r1, c) !== '') text = false;
    if (model.getCellType(grid.sheet, r.r1 + 1, c) !== 2) numberBelow = true;
  }
  return text && numberBelow;
}

/** Выделение без пустого «хвоста»: весь столбец — это только его заполненная часть. */
function usedSelection(): Rect | null {
  const s = grid.sel;
  const d = model.getSheetDimension(grid.sheet);
  const r = { r1: s.r1, c1: s.c1, r2: Math.min(s.r2, d[2]), c2: Math.min(s.c2, d[3]) };
  return r.r2 < r.r1 || r.c2 < r.c1 ? null : r;
}

const BLANKS_CODE = { last: 0, first: 1, zero: 2 } as const;

/** Сортировка, как в Excel. По умолчанию строки переезжают целиком (вся таблица);
 *  если выделен только столбец, а рядом есть данные — спрашиваем, как в Excel.
 *  custom — окно «Настраиваемая сортировка» (уровни, заголовки, пустые). */
async function sortBy(column: number, descending: boolean, custom = false): Promise<void> {
  const s = grid.sel;
  const f = filterOf(grid.sheet);
  let range: Rect;
  let header: boolean;
  let expand = true;
  let canExpand = false;
  if (f) {
    range = { r1: f.r1, c1: f.c1, r2: f.r2, c2: f.c2 };
    header = true;
  } else {
    const sel = usedSelection();
    if (!sel) { showToast('В выделении нет данных для сортировки.'); return; }
    const single = s.r1 === s.r2 && s.c1 === s.c2;
    const region = single ? currentRegion(s.row, s.column) : currentRegion(sel.r1, sel.c1, sel);
    if (single) {
      range = region;
    } else {
      canExpand = region.c1 < sel.c1 || region.c2 > sel.c2;
      if (canExpand && !custom) {
        const answer = await askExpand();
        if (!answer) return;
        if (answer === 'custom') custom = true;
        else expand = answer === 'expand';
      }
      range = expand ? { r1: Math.min(sel.r1, region.r1), c1: region.c1, r2: Math.max(sel.r2, region.r2), c2: region.c2 } : sel;
    }
    header = hasHeader(range);
  }
  let request: SortRequest = { levels: [{ column, descending }], header, blanks: settings.sortBlanks, expand };
  if (custom) {
    const cols = [];
    for (let c = range.c1; c <= range.c2; c++) cols.push({ column: c, title: columnTitle(c, header ? model.getFormattedCellValue(grid.sheet, range.r1, c) : '') });
    const answer = await customSort(cols, request, canExpand);
    if (!answer) { focusGrid(); return; }
    request = answer;
    settings.sortBlanks = answer.blanks;
    saveSettings();
    if (canExpand && !answer.expand) {
      const sel = usedSelection()!;
      range = { ...sel };
    }
  }
  const top = request.header ? range.r1 + 1 : range.r1;
  const area = { sheet: grid.sheet, row: top, column: range.c1, width: range.c2 - range.c1 + 1, height: range.r2 - top + 1 };
  if (area.height < 2) { focusGrid(); return; }
  const keys = request.levels
    .filter((l) => l.column >= area.column && l.column < area.column + area.width)
    .flatMap((l) => [l.column, l.descending ? 1 : 0]);
  if (!keys.length) keys.push(Math.min(Math.max(column, area.column), area.column + area.width - 1), descending ? 1 : 0);
  try {
    model.sortRangeBy(area, keys, BLANKS_CODE[request.blanks]);
  } catch (e) {
    showToast(`Не удалось отсортировать: ${e}`);
  }
  grid.render();
  updateFormulaBar();
  focusGrid();
}

function toggleFilter(): void {
  if (filterOf(grid.sheet)) { setFilter(grid.sheet, null); return; }
  const s = grid.sel;
  const sel = usedSelection();
  if (!sel) { showToast('Выделите таблицу с данными, чтобы включить фильтр.'); return; }
  // Одна ячейка (или одна объединённая) — ищем шапку сами; иначе фильтр в выделенной строке.
  const om = grid.mergeOf(s.row, s.column);
  const single = (s.r1 === s.r2 && s.c1 === s.c2)
    || (!!om && s.r1 === om.row && s.c1 === om.column && s.r2 === om.row + om.height - 1 && s.c2 === om.column + om.width - 1);
  const region = single ? currentRegion(s.row, s.column) : currentRegion(sel.r1, sel.c1, sel);
  let header: number, c1 = region.c1, c2 = region.c2;
  if (!single) {
    // Выделена строка шапки (или ячейки в ней) — фильтр ставим именно в неё.
    header = sel.r1;
    if (!(s.c1 === 1 && s.c2 === LAST_COLUMN)) { c1 = sel.c1; c2 = sel.c2; }
  } else {
    // Одна ячейка: пропускаем строки названия (объединены по ширине) и верхние
    // этажи многострочной шапки (объединения, уходящие ниже), как сделал бы человек.
    header = region.r1;
    const merges = model.getMergedCells(grid.sheet);
    for (let guard = 0; guard < 20 && header < region.r2; guard++) {
      const inRow = merges.filter((m) => m.row <= header && m.row + m.height - 1 >= header && m.column <= region.c2 && m.column + m.width - 1 >= region.c1);
      const title = inRow.some((m) => m.width > 1 && m.row === header && m.height === 1);
      const goesDown = inRow.some((m) => m.row + m.height - 1 > header);
      if (!title && !goesDown) break;
      header++;
    }
  }
  setFilter(grid.sheet, { r1: header, c1, r2: Math.max(region.r2, header), c2, criteria: new Map() });
}

initFilter({
  model: () => model,
  sheet: () => grid.sheet,
  sort: (column, descending) => sortBy(column, descending),
  changed: () => { syncFilter(); updateStatus(); },
});
// После текущего щелчка: иначе тот же mousedown сразу закрыл бы окно.
grid.onFilterButton = (column, x, y) => setTimeout(() => openFilterPopup(column, x, y));

initFind({
  model: () => model,
  sheet: () => grid.sheet,
  active: () => ({ row: grid.sel.row, column: grid.sel.column }),
  goTo: (sheet, row, column) => { if (sheet !== grid.sheet) setSheet(sheet, false); grid.select(row, column); },
  changed: () => { grid.render(); updateFormulaBar(); updateStatus(); },
  toast: showToast,
});

/** Специальная вставка (Ctrl+Alt+V): значения, форматы, транспонировать. */
function pasteSpecialMenu(x: number, y: number): void {
  const c = clip;
  if (!c) {
    showToast('Сначала скопируйте ячейки в ENS-Table (Ctrl+C): специальная вставка работает с ними.');
    return;
  }
  const [r1, c1, r2, c2] = c.range;
  const at = () => { model.setSelectedCell(grid.sel.r1, grid.sel.c1); };
  const done = (h: number, w: number) => {
    grid.selectRange(grid.sel.r1, grid.sel.c1, grid.sel.r1 + h - 1, grid.sel.c1 + w - 1);
    grid.render(); updateFormulaBar(); updateStatus(); focusGrid();
  };
  const styles = () => {
    const out = [];
    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let col = c1; col <= c2; col++) row.push(model.getCellStyle(c.sheet, r, col).style);
      out.push(row);
    }
    return out;
  };
  const run2 = (f: () => void, h: number, w: number) => { try { f(); } catch (e) { showToast(`Не удалось вставить: ${e}`); } done(h, w); };
  showMenu([
    { label: 'Вставить всё', keys: 'Ctrl+V', run: () => paste(c.text) },
    { label: 'Значения', run: () => run2(() => { at(); model.pasteValues(c.sheet, c.range, false); }, r2 - r1 + 1, c2 - c1 + 1) },
    { label: 'Значения и форматы', run: () => run2(() => {
      at(); model.pasteValues(c.sheet, c.range, false);
      model.setSelectedRange(grid.sel.r1, grid.sel.c1, grid.sel.r1, grid.sel.c1); model.onPasteStyles(styles());
    }, r2 - r1 + 1, c2 - c1 + 1) },
    { label: 'Форматы', run: () => run2(() => {
      model.setSelectedCell(grid.sel.r1, grid.sel.c1);
      model.setSelectedRange(grid.sel.r1, grid.sel.c1, grid.sel.r2, grid.sel.c2);
      model.onPasteStyles(styles());
    }, Math.max(r2 - r1 + 1, grid.sel.r2 - grid.sel.r1 + 1), Math.max(c2 - c1 + 1, grid.sel.c2 - grid.sel.c1 + 1)) },
    { label: 'Транспонировать (значения)', separator: true, run: () => run2(() => { at(); model.pasteValues(c.sheet, c.range, true); }, c2 - c1 + 1, r2 - r1 + 1) },
  ], x, y);
}

grid.onContextMenu = (x, y, header) => {
  if (editing) commitEdit(0, 0);
  cellMenu(x, y, header);
};

// Меню WebView («Копировать изображение», «Сохранить изображение») не нужно:
// в полях ввода остаётся обычное меню с «Вырезать/Копировать/Вставить».
document.addEventListener('contextmenu', (e) => {
  const t = e.target as HTMLElement;
  if (!t.closest('input, textarea')) e.preventDefault();
});

document.addEventListener('mousedown', (e) => {
  const menu = $('context-menu');
  if (!menu.hidden && !menu.contains(e.target as Node)) menu.hidden = true;
});
document.addEventListener('keydown', (e) => {
  const menu = $('context-menu');
  if (e.key === 'Escape' && !menu.hidden) { menu.hidden = true; e.stopPropagation(); }
}, true);

// --- Файлы ---
// В браузере файл нельзя перезаписать на месте: «Сохранить» скачивает копию.
// Надёжное сохранение (временный файл → проверка → замена) — в оболочке Tauri (ADR-005).

const fileInput = $<HTMLInputElement>('file-input');
fileInput.addEventListener('change', async () => {
  const file = fileInput.files?.[0];
  if (!file) return;
  try {
    const next = openXlsx(new Uint8Array(await file.arrayBuffer()), file.name, localeId());
    setModel(next, { ...desk.newDoc(file.name.replace(/\.xlsx$/i, '')) });
  } catch (err) {
    alert(`Не удалось открыть «${file.name}»:\n${err}`);
  }
  fileInput.value = '';
});

/** Новая модель в окне (открыт или восстановлен файл). */
function setModel(next: Model, nextDoc: desk.Doc): void {
  if (next !== model) model.free();
  model = next;
  doc = nextDoc;
  resetFilters();
  model.setRecognizeDates(settings.recognizeDates);
  syncDocTitle();
  setSheet(0);
  if (nextDoc.path) {
    desk.rememberRecent(nextDoc.path, nextDoc.title.replace(/ \(восстановлено\)$/, ''));
  }
}

function syncDocTitle(): void {
  const mark = doc.dirty ? ' *' : '';
  document.title = `${doc.title}${mark} — ENS-Table`;
  $('doc-name').textContent = `${doc.title}${mark}`;
  if (isTauri) {
    void import('@tauri-apps/api/window').then(({ getCurrentWindow }) =>
      getCurrentWindow().setTitle(`${doc.title}${mark} — ENS-Table`)).catch(() => {});
  }
}

// --- Буфер обмена ---
// Как в Excel: внутри программы копируются формулы (ссылки сдвигаются),
// стили и ссылки; «Вырезать» переносит ячейки при вставке, а формулы,
// которые на них ссылались, начинают ссылаться на новое место. Другим
// программам уходит текст (значения через табуляцию).

interface Clip { data: unknown; sheet: number; range: [number, number, number, number]; text: string; cut: boolean }
let clip: Clip | null = null;

/** Скопировать выделение: во внутренний буфер и текстом — в системный. */
function takeClip(cut: boolean): string {
  const s = grid.sel;
  model.setSelectedCell(s.r1, s.c1);
  model.setSelectedRange(s.r1, s.c1, s.r2, s.c2);
  const c = model.copyToClipboard();
  const text = toTsv();
  clip = { data: c.data, sheet: grid.sheet, range: c.range, text, cut };
  grid.clipRange = { r1: s.r1, c1: s.c1, r2: Math.min(s.r2, c.range[2]), c2: Math.min(s.c2, c.range[3]), sheet: grid.sheet };
  grid.render();
  return text;
}

function dropClip(): void {
  clip = null;
  grid.clipRange = null;
  grid.render();
}

const sameText = (a: string, b: string) => a.replace(/\r\n?/g, '\n').trim() === b.replace(/\r\n?/g, '\n').trim();

/** Вставка: своё — через движок (формулы, форматы), чужое — как текст. */
function paste(text: string): void {
  if (clip && sameText(text, clip.text)) {
    const c = clip;
    const [r1, c1, r2, c2] = c.range;
    const h = r2 - r1 + 1, w = c2 - c1 + 1;
    const area = selectionArea();
    // Скопированное повторяется на всё выделение, если оно кратно (как в Excel).
    const tilesY = !c.cut && area.height > h && area.height % h === 0 ? area.height / h : 1;
    const tilesX = !c.cut && area.width > w && area.width % w === 0 ? area.width / w : 1;
    try {
      for (let ty = 0; ty < tilesY; ty++) {
        for (let tx = 0; tx < tilesX; tx++) {
          model.setSelectedCell(area.row + ty * h, area.column + tx * w);
          model.pasteFromClipboard(c.sheet, c.range, c.data, c.cut);
        }
      }
    } catch (e) {
      showToast(`Не удалось вставить: ${e}`);
    }
    grid.selectRange(area.row, area.column, area.row + Math.max(h * tilesY, 1) - 1, area.column + Math.max(w * tilesX, 1) - 1);
    if (c.cut) dropClip(); // вырезанное вставляется один раз
  } else {
    if (clip) dropClip();
    pasteText(text);
  }
  grid.render();
  updateFormulaBar();
  updateStatus();
}

editor.addEventListener('copy', (e) => {
  if (editing) return;
  e.preventDefault();
  e.clipboardData?.setData('text/plain', takeClip(false));
});
editor.addEventListener('cut', (e) => {
  if (editing) return;
  e.preventDefault();
  e.clipboardData?.setData('text/plain', takeClip(true));
});
editor.addEventListener('paste', (e) => {
  if (editing) return; // в режиме ввода — обычная вставка текста
  e.preventDefault();
  paste(e.clipboardData?.getData('text/plain') ?? '');
});

/** Вставка, как в Excel: если выделено несколько ячеек, а скопировано меньше —
 *  скопированное повторяется на всё выделение (одно значение — во все ячейки). */
function pasteText(text: string): void {
  if (!text) return;
  const rows = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((r) => r.split('\t'));
  const h = rows.length, w = Math.max(...rows.map((r) => r.length));
  const area = selectionArea();
  let out = rows;
  if ((area.height > h || area.width > w) && area.height % h === 0 && area.width % w === 0) {
    out = [];
    for (let r = 0; r < area.height; r++) {
      const src = rows[r % h];
      out.push(Array.from({ length: area.width }, (_, c) => src[c % w] ?? ''));
    }
  }
  model.pasteCsvText(area, out.map((r) => r.join('\t')).join('\n'));
  grid.render();
  updateFormulaBar();
  updateStatus();
}

// --- Команды ---

function style(path: string, value: (current: boolean) => string, flag?: (s: ReturnType<Model['getCellStyle']>['style']) => boolean): void {
  const s = grid.sel;
  const current = flag ? flag(model.getCellStyle(grid.sheet, s.row, s.column).style) : false;
  model.updateRangeStyle(selectionArea(), path, value(current));
  grid.render();
  refreshRibbon();
  syncNumberFormat();
}

function toTsv(): string {
  const s = grid.sel, rows: string[] = [];
  for (let r = s.r1; r <= s.r2; r++) {
    const row: string[] = [];
    for (let c = s.c1; c <= s.c2; c++) row.push(model.getFormattedCellValue(grid.sheet, r, c));
    rows.push(row.join('\t'));
  }
  return rows.join('\n');
}

register(
  { id: 'file.open', label: 'Открыть', keywords: ['open', 'файл', 'xlsx'], keys: ['Ctrl+KeyO'], run: () => void desk.openDialog() },
  { id: 'file.save', label: isDesktop ? 'Сохранить' : 'Сохранить копию (xlsx)', keywords: ['save', 'скачать'], keys: ['Ctrl+KeyS'],
    run: () => void desk.save() },
  { id: 'file.saveAs', label: 'Сохранить как…', keywords: ['save as', 'сохранить под именем'], keys: ['F12', 'Ctrl+Shift+KeyS'],
    run: () => void desk.saveAs() },
  { id: 'edit.undo', label: 'Отменить', keywords: ['undo', 'назад', 'отмена'], keys: ['Ctrl+KeyZ'], run: () => undoRedo('undo') },
  { id: 'edit.redo', label: 'Повторить', keywords: ['redo', 'вернуть'], keys: ['Ctrl+KeyY', 'Ctrl+Shift+KeyZ'], run: () => undoRedo('redo') },
  { id: 'edit.clear', label: 'Очистить содержимое', keywords: ['delete', 'удалить'], keys: ['Delete'], run: () => {
    const s = grid.sel; model.rangeClearContents(grid.sheet, s.r1, s.c1, s.r2, s.c2); grid.render(); updateFormulaBar();
  } },
  { id: 'edit.autosum', label: 'Автосумма', keywords: ['sum', 'сумма', 'сумм', 'итог'], keys: ['Alt+Equal'], run: autosum },
  { id: 'edit.insertFunction', label: 'Вставить функцию', keywords: ['insert function', 'функция', 'fx', 'мастер функций'], keys: ['Shift+F3'],
    run: () => { search.value = 'ƒ '; search.focus(); showPicker(functionList(''), search); } },
  { id: 'clipboard.copy', label: 'Копировать', keywords: ['copy'], keys: ['Ctrl+KeyC'], nativeKeys: true,
    run: async () => { await copyText(takeClip(false)).catch((e) => showToast(String(e))); focusGrid(); } },
  { id: 'clipboard.cut', label: 'Вырезать', keywords: ['cut'], keys: ['Ctrl+KeyX'], nativeKeys: true,
    run: async () => { await copyText(takeClip(true)).catch((e) => showToast(String(e))); focusGrid(); } },
  { id: 'clipboard.paste', label: 'Вставить', keywords: ['paste'], keys: ['Ctrl+KeyV'], nativeKeys: true, run: async () => {
    try {
      paste(await readText());
    } catch {
      showToast('Не удалось прочитать буфер обмена с кнопки. Нажмите Ctrl+V — так вставка работает всегда.');
    }
    focusGrid();
  } },
  { id: 'format.bold', label: 'Полужирный', keywords: ['bold', 'жирный'], keys: ['Ctrl+KeyB'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.font.b,
    run: () => style('font.b', (cur) => String(!cur), (st) => st.font.b) },
  { id: 'format.italic', label: 'Курсив', keywords: ['italic'], keys: ['Ctrl+KeyI'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.font.i,
    run: () => style('font.i', (cur) => String(!cur), (st) => st.font.i) },
  { id: 'format.underline', label: 'Подчёркнутый', keywords: ['underline'], keys: ['Ctrl+KeyU'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.font.u,
    run: () => style('font.u', (cur) => String(!cur), (st) => st.font.u) },
  { id: 'format.fill', label: 'Цвет заливки', keywords: ['fill', 'заливка', 'цвет', 'фон'], run: (c) => style('fill.fg_color', () => c ?? '') },
  { id: 'format.fontColor', label: 'Цвет шрифта', keywords: ['font color', 'цвет текста'], run: (c) => style('font.color', () => c ?? '') },
  { id: 'format.fontName', label: 'Шрифт', keywords: ['font', 'гарнитура', 'начертание'], run: (f) => { if (f) style('font.name', () => f); } },
  { id: 'format.fontSize', label: 'Размер шрифта', keywords: ['font size', 'кегль'], run: (n) => { if (n) style('font.size', () => n); } },
  { id: 'align.left', label: 'По левому краю', keywords: ['align left'], run: () => style('alignment.horizontal', () => 'left') },
  { id: 'align.center', label: 'По центру', keywords: ['align center', 'центр'], run: () => style('alignment.horizontal', () => 'center') },
  { id: 'align.right', label: 'По правому краю', keywords: ['align right'], run: () => style('alignment.horizontal', () => 'right') },
  { id: 'align.top', label: 'По верхнему краю', keywords: ['align top', 'вверху'], run: () => style('alignment.vertical', () => 'top') },
  { id: 'align.middle', label: 'По середине (по вертикали)', keywords: ['middle', 'по вертикали'], run: () => style('alignment.vertical', () => 'center') },
  { id: 'align.bottom', label: 'По нижнему краю', keywords: ['align bottom', 'внизу'], run: () => style('alignment.vertical', () => 'bottom') },
  { id: 'format.wrap', label: 'Перенос текста', keywords: ['wrap text', 'перенос', 'переносить по словам'],
    checked: () => !!model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.wrap_text,
    run: () => { style('alignment.wrap_text', (cur) => String(!cur), (st) => !!st.alignment?.wrap_text); fitRows(grid.sel.r1, grid.sel.r2); } },
  { id: 'format.merge', label: 'Объединить и поместить в центр', keywords: ['merge', 'объединить ячейки', 'объединение'],
    run: () => merge('center') },
  { id: 'format.unmerge', label: 'Отменить объединение ячеек', keywords: ['unmerge', 'разъединить'], run: () => merge('undo') },
  { id: 'format.borders', label: 'Границы (последний выбор)', keywords: ['borders', 'рамка', 'границы', 'обрамление'],
    run: () => applyBorders() },
  { id: 'format.borders:menu', label: 'Границы: выбрать вид и линию', keywords: ['тип линии', 'толщина границы'],
    run: (at) => {
      const [x, y] = (at ?? '').split(',').map(Number);
      bordersMenu(Number.isFinite(x) ? x : lastPointer.x, Number.isFinite(y) ? y : lastPointer.y + 8);
    } },
  { id: 'format.number', label: 'Числовой формат', keywords: ['number format', 'формат ячеек'], run: (fmt) => { if (fmt) style('num_fmt', () => fmt); } },
  { id: 'sheet.new', label: 'Новый лист', keywords: ['new sheet', 'лист'], keys: ['Shift+F11'], run: () => {
    const name = nextSheetName(model);
    model.newSheet();
    const index = model.getWorksheetsProperties().length - 1;
    model.renameSheet(index, name);
    setSheet(index);
  } },
  { id: 'sheet.rename', label: 'Переименовать лист', keywords: ['rename sheet', 'имя листа'], run: () => renameSheet(grid.sheet) },
  { id: 'view.focusCell', label: 'Фокус ячейки (перекрестие)', keywords: ['focus cell', 'подсветка строки', 'перекрестие', 'crosshair'],
    keys: ['Ctrl+Shift+KeyK'], checked: () => grid.focusCell.enabled,
    run: () => { grid.focusCell.enabled = !grid.focusCell.enabled; grid.render(); refreshRibbon(); } },
  { id: 'file.print', label: 'Печать и PDF', keywords: ['print', 'печать', 'pdf', 'распечатать'], keys: ['Ctrl+KeyP'],
    run: () => showBackstage('print') },
  { id: 'edit.find', label: 'Найти', keywords: ['find', 'поиск', 'найти'], keys: ['Ctrl+KeyF'],
    run: () => showFind('find') },
  { id: 'edit.replace', label: 'Заменить', keywords: ['replace', 'замена', 'заменить'], keys: ['Ctrl+KeyH'], run: () => showFind('replace') },
  { id: 'data.sortAsc', label: 'Сортировка от А до Я (по возрастанию)', keywords: ['sort', 'сортировка', 'по возрастанию', 'упорядочить'],
    run: () => sortBy(grid.sel.column, false) },
  { id: 'data.sortDesc', label: 'Сортировка от Я до А (по убыванию)', keywords: ['sort descending', 'по убыванию'],
    run: () => sortBy(grid.sel.column, true) },
  { id: 'data.sortCustom', label: 'Настраиваемая сортировка…', keywords: ['custom sort', 'сортировка по нескольким столбцам', 'уровни сортировки'],
    run: () => sortBy(grid.sel.column, false, true) },
  { id: 'data.filter', label: 'Фильтр', keywords: ['filter', 'автофильтр', 'отбор'], keys: ['Ctrl+Shift+KeyL'],
    checked: () => !!filterOf(grid.sheet), run: toggleFilter },
  { id: 'clipboard.pasteSpecial', label: 'Специальная вставка', keywords: ['paste special', 'вставить значения', 'транспонировать'],
    keys: ['Ctrl+Alt+KeyV'], run: () => pasteSpecialMenu(lastPointer.x, lastPointer.y) },
  { id: 'view.freeze', label: 'Закрепить области', keywords: ['freeze panes', 'закрепить', 'заморозить', 'шапка', 'фиксировать'],
    run: () => freezeMenu(lastPointer.x, lastPointer.y + 8) },
  { id: 'view.gridlines', label: 'Сетка', keywords: ['gridlines', 'линии сетки'], checked: () => model.getShowGridLines(grid.sheet),
    run: () => { model.setShowGridLines(grid.sheet, !model.getShowGridLines(grid.sheet)); grid.render(); refreshRibbon(); } },
  { id: 'nav.precedents', label: 'Перейти к влияющим ячейкам', keywords: ['trace precedents', 'ссылки формулы'], keys: ['Ctrl+BracketLeft'],
    run: () => {
      const refs = refsIn(model.getCellContent(grid.sheet, grid.sel.row, grid.sel.column).startsWith('=')
        ? model.getCellContent(grid.sheet, grid.sel.row, grid.sel.column) : '');
      if (refs.length === 1) goTo(refs[0]);
      else if (refs.length > 1) showPicker(refs.map((r) => ({ label: r.text, run: () => goTo(r) })));
    } },
  { id: 'nav.back', label: 'Вернуться к ячейке, из которой перешли по ссылке', keywords: ['back', 'назад', 'вернуться'], keys: ['Alt+ArrowLeft'], run: () => {
    const prev = backStack.pop();
    if (!prev) {
      showToast('Возвращаться некуда: кнопка работает после перехода по ссылке (Ctrl+[ или Ctrl+щелчок по ссылке в строке формул). Отмена действия — Ctrl+Z.');
      return;
    }
    if (prev.sheet !== grid.sheet) setSheet(prev.sheet);
    grid.select(prev.row, prev.column);
  } },
  { id: 'search.focus', label: 'Поиск команд и функций', keys: ['Alt+KeyQ'], run: () => search.focus() },
  { id: 'options.expandEditor', label: 'Расширять поле ввода вправо', keywords: ['поле ввода', 'редактор', 'expand'],
    checked: () => settings.expandEditor,
    run: () => { settings.expandEditor = !settings.expandEditor; saveSettings(); refreshRibbon(); } },
  { id: 'options.decimal', label: 'Разделитель дробной части', keywords: ['запятая', 'точка', 'разделитель', 'decimal'],
    run: (value) => {
      if (value !== 'comma' && value !== 'point') return;
      settings.decimal = value; saveSettings();
      model.setLocale(localeId());
      grid.render(); updateFormulaBar(); updateStatus();
    } },
  { id: 'options.recognizeDates', label: 'Распознавать даты при вводе', keywords: ['даты', 'автоформат', 'dates', 'autoformat', 'параметры'],
    checked: () => settings.recognizeDates,
    run: () => { settings.recognizeDates = !settings.recognizeDates; saveSettings(); model.setRecognizeDates(settings.recognizeDates); refreshRibbon(); } },
  { id: 'file.backstage', label: 'Файл: создать, открыть, сохранить, печать, параметры', keywords: ['файл', 'file', 'параметры', 'options'], keys: ['Alt+KeyF'],
    run: () => showBackstage() },
  { id: 'file.options', label: 'Параметры', keywords: ['options', 'настройки', 'settings'], run: () => showBackstage('options') },
  { id: 'file.new', label: 'Создать книгу', keywords: ['new', 'новая книга'], keys: ['Ctrl+KeyN'], run: () => newBook() },
);

// --- Выпадающий список (выбор ссылки, результаты поиска) ---

const picker = $('picker');

function showPicker(items: { label: string; hint?: string; run: () => void }[], anchor?: HTMLElement): void {
  picker.replaceChildren();
  items.slice(0, 12).forEach((it, i) => {
    const b = document.createElement('button');
    b.className = `pick${i === 0 ? ' active' : ''}`;
    b.innerHTML = `<span>${it.label}</span>${it.hint ? `<kbd>${it.hint}</kbd>` : ''}`;
    b.addEventListener('mousedown', (e) => { e.preventDefault(); hidePicker(); it.run(); });
    picker.append(b);
  });
  const r = (anchor ?? grid.scroller).getBoundingClientRect();
  const cell = anchor ? null : grid.cellRect(grid.sel.row, grid.sel.column);
  picker.style.left = `${anchor ? r.left : r.left + cell!.x}px`;
  picker.style.top = `${anchor ? r.bottom + 2 : r.top + cell!.y + cell!.h + 2}px`;
  picker.style.minWidth = `${anchor ? r.width : 160}px`;
  picker.hidden = items.length === 0;
}

let toastTimer = 0;
function showToast(text: string): void {
  const t = $('toast');
  t.textContent = text;
  t.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => (t.hidden = true), 6000);
}

function hidePicker(): void {
  picker.hidden = true;
}

// --- §6: поиск в заголовке окна ---

const search = $<HTMLInputElement>('search');
search.insertAdjacentHTML('beforebegin', icon('search', 20));

/** Лист «Пример ВПР»: объяснение, данные и живые формулы (идея пользователя). */
function createExampleSheet(ex: FunctionExample): void {
  const names = model.getWorksheetsProperties().map((p) => p.name.toLowerCase());
  let name = `Пример ${ex.fn}`;
  for (let i = 2; names.includes(name.toLowerCase()); i++) name = `Пример ${ex.fn} (${i})`;
  model.newSheet();
  const sh = model.getWorksheetsProperties().length - 1;
  model.renameSheet(sh, name);
  const set = (r: number, c: number, v: string) => model.setUserInput(sh, r, c, v);
  const text = (v: string) => (/^[=+\-]/.test(v) ? `'${v}` : v); // показать формулу как текст
  const area = (row: number, column: number, width = 1, height = 1) => ({ sheet: sh, row, column, width, height });
  const bold = (r: number, c: number, w = 1) => model.updateRangeStyle(area(r, c, w), 'font.b', 'true');
  const line = { item: { style: 'thin', color: '#9fc9bd' }, type: 'All' as const };
  const fill = '#e6f3ef';

  set(1, 1, ex.title);
  model.updateRangeStyle(area(1, 1), 'font.size', '15');
  bold(1, 1);
  set(2, 1, text(ex.syntax));
  model.updateRangeStyle(area(2, 1), 'font.color', '#1d5a4f');
  ex.args.slice(0, 5).forEach(([arg, what], i) => { set(3 + i, 1, arg); bold(3 + i, 1); set(3 + i, 2, what); });

  set(9, 1, 'Данные для примеров');
  bold(9, 1);
  const width = Math.max(...ex.data.map((r) => r.length));
  ex.data.slice(0, 9).forEach((row, i) => row.forEach((v, j) => set(10 + i, 1 + j, v)));
  model.updateRangeStyle(area(10, 1, width), 'fill.color', fill);
  bold(10, 1, width);
  model.setAreaWithBorder(area(10, 1, width, Math.min(ex.data.length, 9)), line);

  const r0 = 21;
  set(r0 - 1, 1, 'Примеры: меняйте данные выше — результаты пересчитаются');
  bold(r0 - 1, 1);
  ['Формула', 'Результат', 'Что делает'].forEach((h, j) => set(r0, 1 + j, h));
  model.updateRangeStyle(area(r0, 1, 3), 'fill.color', fill);
  bold(r0, 1, 3);
  ex.examples.forEach(([f, note], i) => {
    set(r0 + 1 + i, 1, text(f));
    set(r0 + 1 + i, 2, f);
    set(r0 + 1 + i, 3, note);
  });
  model.setAreaWithBorder(area(r0, 1, 2, ex.examples.length + 1), line);
  model.updateRangeStyle(area(r0 + 1, 1, 1, ex.examples.length), 'font.color', '#1d5a4f');
  model.updateRangeStyle(area(r0 + 1, 2, 1, ex.examples.length), 'font.b', 'true');

  const e0 = r0 + ex.examples.length + 3;
  set(e0, 1, 'Частые ошибки');
  bold(e0, 1);
  ex.errors.forEach((t, i) => set(e0 + 1 + i, 1, `• ${t}`));

  model.setColumnsWidth(sh, 1, 1, 360);
  model.setColumnsWidth(sh, 2, 4, 120);
  setSheet(sh);
  grid.select(r0 + 1, 2);
  showToast(`Лист «${name}» добавлен. Это обычный лист: его можно менять или удалить.`);
}

function functionList(q: string) {
  return FUNCTIONS
    .map((f) => ({ f, score: q ? Math.max(matchScore(q, f.ru), matchScore(q, f.en), matchScore(q, f.description) * 0.7) : 1 }))
    .filter((x) => x.score > 0)
    .sort((a, b) => b.score - a.score)
    .flatMap(({ f }) => {
      const items = [{ label: `ƒ ${f.ru} — ${f.description}`, hint: f.en, run: () => startEdit(`=${f.ru}(`) }];
      const ex = exampleFor(f.ru);
      if (ex && q) items.push({ label: `📘 ${f.ru}: пример на отдельном листе`, hint: 'справка', run: () => createExampleSheet(ex) });
      return items;
    });
}

function searchResults(q: string) {
  if (q.startsWith('ƒ')) return functionList(q.slice(1).trim()).map((x) => ({ ...x, score: 1 }));
  const found: { score: number; label: string; hint?: string; run: () => void }[] = [];
  for (const c of allCommands()) {
    if (c.id === 'search.focus') continue;
    const score = Math.max(matchScore(q, c.label), ...(c.keywords ?? []).map((k) => matchScore(q, k) * 0.9));
    if (score > 0) found.push({ score: score + 5, label: c.label, hint: c.keys?.[0] && keyLabel(c.keys[0]), run: () => c.run() });
  }
  for (const f of FUNCTIONS) {
    const score = Math.max(matchScore(q, f.ru), matchScore(q, f.en), matchScore(q, f.description) * 0.7);
    if (score > 0) {
      found.push({ score, label: `ƒ ${f.ru} (${f.en}) — ${f.description}`, run: () => startEdit(`=${f.ru}(`) });
      const ex = exampleFor(f.ru);
      if (ex) found.push({ score: score - 0.01, label: `📘 ${f.ru}: пример на отдельном листе`, run: () => createExampleSheet(ex) });
    }
  }
  return found.sort((a, b) => b.score - a.score);
}

search.addEventListener('input', () => showPicker(searchResults(search.value), search));
search.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') { search.value = ''; hidePicker(); focusGrid(); }
  if (e.key === 'Enter') {
    const first = searchResults(search.value)[0];
    hidePicker(); search.value = '';
    focusGrid();
    first?.run();
  }
});
search.addEventListener('blur', () => setTimeout(hidePicker, 100));

// Сочетания, которые работают из любого места окна.
window.addEventListener('keydown', (e) => {
  if (e.target === editor || e.target === formula) return;
  if (e.altKey && e.code === 'KeyQ') { e.preventDefault(); search.focus(); }
});

// --- Запуск ---

const refreshRibbon = renderRibbon($('tabs'), $('ribbon'), () => showBackstage());
/** Списки ленты (формат, шрифт, размер) показывают значения активной ячейки. */
function syncNumberFormat(): void {
  const st = model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style;
  const setSelect = (cmd: string, value: string) => {
    const el = document.querySelector<HTMLSelectElement>(`select[data-command="${cmd}"]`);
    if (!el) return;
    if (![...el.options].some((o) => o.value === value)) el.add(new Option(value, value));
    el.value = value;
  };
  setSelect('format.fontName', st.font.name);
  setSelect('format.fontSize', String(st.font.sz));
  setSelect('options.decimal', settings.decimal);
  const select = document.querySelector<HTMLSelectElement>('select[data-command="format.number"]');
  if (!select) return;
  const fmt = st.num_fmt;
  const known = [...select.options].some((o) => o.value === fmt);
  if (!known) {
    select.querySelector('option[data-custom]')?.remove();
    const custom = new Option(formatName(fmt), fmt);
    custom.dataset.custom = '1';
    select.add(custom);
  }
  select.value = fmt;
}

grid.onSelect = () => { updateFormulaBar(); updateStatus(); refreshRibbon(); syncNumberFormat(); if (!editing) placeEditor(); };
grid.scroller.addEventListener('scroll', () => { if (!editing) placeEditor(); });
grid.scroller.addEventListener('mouseup', () => { if (!editing) focusGrid(); });
grid.scroller.addEventListener('focus', () => { if (!editing) focusGrid(); });
syncDocTitle();
renderSheetTabs();
grid.select(1, 1);
focusGrid();

function autosum(): void {
  if (editing) commitEdit(0, 0);
  const s = grid.sel;
  const isNum = (r: number, c: number) => model.getCellType(grid.sheet, r, c) === 1;
  const sheet = grid.sheet;
  const ref = (r1: number, c1: number, r2: number, c2: number) =>
    `${columnName(c1)}${r1}${r1 === r2 && c1 === c2 ? '' : `:${columnName(c2)}${r2}`}`;
  if (s.r1 !== s.r2 || s.c1 !== s.c2) {
    // Выделен диапазон: суммы под каждым столбцом (как в Excel). Если нижняя
    // строка выделения пустая — суммы пишутся в неё.
    const lastEmpty = Array.from({ length: s.c2 - s.c1 + 1 }, (_, i) => s.c1 + i)
      .every((c) => !model.getFormattedCellValue(sheet, s.r2, c));
    const target = lastEmpty && s.r2 > s.r1 ? s.r2 : s.r2 + 1;
    const bottom = lastEmpty && s.r2 > s.r1 ? s.r2 - 1 : s.r2;
    for (let c = s.c1; c <= s.c2; c++) model.setUserInput(sheet, target, c, `=СУММ(${ref(s.r1, c, bottom, c)})`);
    grid.select(target, s.c1);
    grid.render();
    return;
  }
  // Одна ячейка: сумма чисел над ней, а если над ней пусто — слева.
  let top = s.row - 1;
  while (top >= 1 && isNum(top, s.column)) top--;
  if (top < s.row - 1) { startEdit(`=СУММ(${ref(top + 1, s.column, s.row - 1, s.column)})`); return; }
  let left = s.column - 1;
  while (left >= 1 && isNum(s.row, left)) left--;
  startEdit(left < s.column - 1 ? `=СУММ(${ref(s.row, left + 1, s.row, s.column - 1)})` : '=СУММ()');
}

async function newBook(): Promise<void> {
  if (isDesktop) { await desk.newWindow(); return; }
  if (doc.dirty && !confirm('Текущая книга не сохранена. Создать новую? Несохранённые изменения пропадут.')) return;
  setModel(newWorkbook(localeId()), desk.newDoc());
}

initBackstage({
  docInfo: () => ({ title: doc.title, path: doc.path, dirty: doc.dirty }),
  recent: () => settings.recent,
  newBook: () => void newBook(),
  open: () => void desk.openDialog(),
  openPath: (p) => void desk.openPath(p),
  save: () => void desk.save(),
  saveAs: () => void desk.saveAs(),
  close: () => void desk.closeBook(),
  applySettings: () => {
    model.setLocale(localeId());
    model.setRecognizeDates(settings.recognizeDates);
    grid.render(); updateFormulaBar(); updateStatus(); refreshRibbon(); syncNumberFormat();
  },
  recoveryFolder: () => desk.recoveryFolder(),
  chooseRecoveryFolder: () => desk.chooseRecoveryFolder(),
  printPreview: (el, opts) => renderPreview(el, model, grid.sheet, opts, grid.sel),
  print: (opts) => {
    try { printSheet(model, grid.sheet, opts, grid.sel, doc.title); } catch (e) { showToast(String((e as Error).message ?? e)); }
  },
});
void hideBackstage;

if (import.meta.env.DEV) Object.assign((window as unknown as { __ens: object }).__ens, { example: (fn: string) => createExampleSheet(exampleFor(fn)!) });

if (location.protocol === 'file:') {
  // Версия одним файлом для браузера: сохранять на место браузер не умеет.
  showToast('ENS-Table в браузере: «Сохранить» скачивает книгу в папку «Загрузки». Восстановление после сбоя — только в настольной версии.');
}

await desk.start({
  model: () => model,
  setModel,
  doc: () => doc,
  onDocChanged: syncDocTitle,
  toast: showToast,
});
