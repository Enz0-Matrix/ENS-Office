import './styles.css';
import './themes.css';
import './themes2.css';
import { allCommands, handleKey, keyLabel, matchScore, register, run } from './commands';
import { CHART_TEMPLATES, parseRange, type ChartLook } from './chart';
import { type BorderType, type CellStyle, type RichRun, columnName, columnNumber, loadEngine, newBookFont, newWorkbook, nextSheetName, openXlsx, type Area, type Model } from './engine';
import { fontStack, loadFonts } from './fonts';
import { initFind, showFind } from './findreplace';
import { showLog, userName } from './changelog';
import { DEFAULT_PRINT, fromSaved, printAreaOf, printSheet, printTitles, renderPreview, toSaved } from './print';
import { exampleFor, type FunctionExample } from './examples';
import { askExpand, columnTitle, customSort, type SortRequest } from './sortdialog';
import { conditionalMenu, editRule, initConditional, manageRules, rectRef } from './cfdialog';
import { changeDecimals, formatCellsDialog, toggleGrouping } from './numfmt';
import { THEMES, allThemes, applyTheme, setHiddenThemes, parseThemeFile, setCustomThemes, systemReducesMotion, themeTemplate } from './themes';
import { addFilter, filterAt, filtersOf, initFilter, openFilterPopup, removeFilter, resetFilters, type SheetFilter } from './filter';
import { hideBackstage, initBackstage, showBackstage } from './backstage';
import { copyText, readText } from './platform';
import { Grid, type OutlineGroup } from './grid';
import { FUNCTION_CATEGORIES } from './funclib';
import { openSymbols } from './symbols';
import { defineFromSelection, initNames, nameFormula, openNameManager } from './names';
import { icon, onRibbonDrawn, pickColor, renderRibbon } from './ribbon';
import { FUNCTIONS } from './functions';
import { localeId, saveSettings, settings } from './settings';
import * as desk from './desktop';
import { isDesktop, isTauri } from './platform';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
const LAST_ROW = 1_048_576;
const LAST_COLUMN = 16_384;

await Promise.all([loadEngine(), loadFonts()]);
newBookFont.name = settings.defaultFont;
newBookFont.size = settings.defaultFontSize;
let model: Model = newWorkbook(localeId());
model.setRecognizeDates(settings.recognizeDates);
let doc: desk.Doc = desk.newDoc();
window.__ensDoc = () => doc;

const grid = new Grid($('grid'), () => model);
// Для проверок в режиме разработки (в сборку не попадает).
if (import.meta.env.DEV) Object.assign(window, { __ens: { model: () => model, grid } });
const nameBox = $<HTMLInputElement>('name-box');
const formula = $<HTMLTextAreaElement>('formula');
const editor = $<HTMLTextAreaElement>('cell-editor');
const status = $('status-stats');

// --- Редактирование ячейки ---

let editing = false;

/** Выделено несколько диапазонов (Ctrl+щелчок): действие — для каждого, как в Excel,
 *  а Ctrl+Z отменяет его целиком. Без них — просто fn(). */
function forEachRange(fn: () => void): void {
  const extra = grid.extraRanges;
  if (!extra.length) { fn(); return; }
  const keep = grid.sel;
  const mark = model.undoMark();
  try {
    for (const r of [...extra, keep]) {
      grid.sel = { row: r.r1, column: r.c1, r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2 };
      fn();
    }
  } finally {
    grid.sel = keep;
    grid.extraRanges = extra;
    model.mergeUndoSince(mark);
  }
  grid.invalidateLayout();
}

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
  // Шрифт — как у ячейки, высота — не меньше строки текста: в низких строках
  // (файлы с мелким шрифтом) набираемый текст иначе обрезался и был не виден.
  let font = { sz: 11, b: false, i: false, name: '' };
  try { font = model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.font; } catch { /* по умолчанию */ }
  const px = Math.round((font.sz || 11) * 1.33 * grid.zoom);
  editor.style.font = `${font.i ? 'italic ' : ''}${font.b ? 'bold ' : ''}${px}px ${fontStack(font.name)}`;
  editor.style.lineHeight = '1.2';
  editor.style.height = `${Math.max(r.h, Math.ceil(px * 1.2) + 2)}px`;
  if (editing && settings.expandEditor) growEditor();
}

/** Поле ввода остаётся на экране при прокрутке и смене листа: выделяем его тенью и меткой с адресом ячейки. */
function syncEditorTag(): void {
  const tag = $('editor-tag');
  if (!editing || editor.classList.contains('away')) { tag.hidden = true; editor.classList.remove('floating'); return; }
  const er = editor.getBoundingClientRect();
  let away = grid.sheet !== editCell.sheet;
  if (!away) {
    const host = grid.scroller.getBoundingClientRect(), r = grid.cellRect(editCell.row, editCell.column);
    away = Math.abs(host.left + r.x - er.left) > 3 || Math.abs(host.top + r.y - er.top) > 3;
  }
  editor.classList.toggle('floating', away);
  tag.hidden = !away;
  if (!away) return;
  const sheetName = model.getWorksheetsProperties()[editCell.sheet]?.name ?? '';
  tag.textContent = `Ввод: ${sheetName}!${columnName(editCell.column)}${editCell.row}`;
  tag.style.left = `${er.left}px`;
  tag.style.top = `${er.top >= 22 ? er.top - 20 : er.bottom}px`;
}
grid.scroller.addEventListener('scroll', syncEditorTag);
setInterval(syncEditorTag, 250);

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

// --- Rich text: оформление части текста при вводе ---

type Fmt = Omit<RichRun, 'text'>;
const fmtKey = (f: Fmt) => JSON.stringify([f.b, f.i, f.u, f.strike, f.sz, f.color, f.name, f.vert]);
/** Оформление каждого знака редактируемого текста (по индексу знака в строке JS). */
let rich: { fmt: Fmt[]; text: string } | null = null;
let lastEditInput: HTMLTextAreaElement | HTMLInputElement = editor;

function richFromModel(sheet: number, row: number, column: number): void {
  rich = null;
  try {
    const runs = model.getRichText(sheet, row, column);
    if (!runs) return;
    const fmt: Fmt[] = [];
    for (const r of runs) { const { text, ...f } = r; for (let i = 0; i < text.length; i++) fmt.push(f); }
    rich = { fmt, text: runs.map((r) => r.text).join('') };
  } catch { rich = null; }
}

/** Ввод изменил текст: вставленные знаки берут оформление соседнего слева. */
function richReconcile(value: string): void {
  if (!rich) rich = { fmt: [], text: '' };
  const old = rich.text;
  if (old === value) return;
  let p = 0;
  while (p < old.length && p < value.length && old[p] === value[p]) p++;
  let s = 0;
  while (s < old.length - p && s < value.length - p && old[old.length - 1 - s] === value[value.length - 1 - s]) s++;
  const inherit: Fmt = rich.fmt[p - 1] ?? rich.fmt[p] ?? {};
  const inserted = Array.from({ length: value.length - p - s }, () => inherit);
  rich.fmt = [...rich.fmt.slice(0, p), ...inserted, ...rich.fmt.slice(old.length - s)];
  rich.text = value;
}

/** Части текста с одинаковым оформлением; null, если оформления нет или это формула. */
function richRuns(value: string): RichRun[] | null {
  if (!rich || value.startsWith('=')) return null;
  richReconcile(value);
  const f = rich!.fmt;
  if (!f.some((x) => fmtKey(x) !== fmtKey({}))) return null;
  const runs: RichRun[] = [];
  for (let i = 0; i < value.length; i++) {
    const last = runs[runs.length - 1];
    if (last && fmtKey(f[i]) === fmtKey(f[i - 1])) last.text += value[i];
    else runs.push({ ...f[i], text: value[i] });
  }
  return runs;
}

/** Команда оформления при выделенной части текста в поле ввода — применяется к части, а не к ячейке. */
function richStyle(path: string, value: string): boolean {
  if (!editing) return false;
  const inp = lastEditInput;
  const a = inp.selectionStart ?? 0, b = inp.selectionEnd ?? 0;
  if (a === b || inp.value.startsWith('=')) return false;
  richReconcile(inp.value);
  const f = rich!.fmt;
  const part = f.slice(a, b);
  const all = (get: (x: Fmt) => unknown) => part.every((x) => !!get(x));
  const apply = (set: (x: Fmt) => Fmt) => { for (let i = a; i < b; i++) f[i] = set(f[i]); };
  switch (path) {
    case 'font.b': { const on = !all((x) => x.b); apply((x) => ({ ...x, b: on })); break; }
    case 'font.i': { const on = !all((x) => x.i); apply((x) => ({ ...x, i: on })); break; }
    case 'font.u': { const on = !all((x) => x.u); apply((x) => ({ ...x, u: on })); break; }
    case 'font.strike': { const on = !all((x) => x.strike); apply((x) => ({ ...x, strike: on })); break; }
    case 'font.color': if (/^#[0-9a-f]{6}$/i.test(value)) apply((x) => ({ ...x, color: value })); else return false; break;
    case 'font.size': { const n = Number(value); if (!(n > 0)) return false; apply((x) => ({ ...x, sz: n })); break; }
    case 'font.name': apply((x) => ({ ...x, name: value })); break;
    case 'font.vert_align': apply((x) => ({ ...x, vert: value === 'superscript' ? 1 : value === 'subscript' ? 2 : 0 })); break;
    default: return false;
  }
  showToast('Оформление части текста применится при вводе (Enter).');
  return true;
}

editor.addEventListener('input', () => { if (editing) richReconcile(editor.value); lastEditInput = editor; });
formula.addEventListener('input', () => { if (editing) richReconcile(formula.value); lastEditInput = formula; });
editor.addEventListener('focus', () => { lastEditInput = editor; });
formula.addEventListener('focus', () => { lastEditInput = formula; });

function startEdit(initial?: string): void {
  if (model.isSheetProtected(grid.sheet) && (model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.protection?.locked ?? true)) {
    showToast('Ячейка защищена от изменений. Снимите защиту листа (Рецензирование) или защиту ячейки.');
    return;
  }
  editMode = initial === undefined ? 'edit' : 'enter';
  editCell = { sheet: grid.sheet, row: grid.sel.row, column: grid.sel.column };
  if (initial === undefined) richFromModel(grid.sheet, grid.sel.row, grid.sel.column); else rich = null;
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
    // «1 1/2» — число 1,5 с форматом «Дробный» (# ?/?), как в Excel; в ячейке «Общий» и только такой ввод.
    const frac = /^(-?)(\d+) +(\d+)\/(\d+)$/.exec(value.trim());
    if (frac && Number(frac[4]) > 0 && Number(frac[3]) < Number(frac[4]) * 1000 && before.toLowerCase() === 'general') {
      const num = (Number(frac[2]) + Number(frac[3]) / Number(frac[4])) * (frac[1] ? -1 : 1);
      const text = String(Math.round(num * 1e12) / 1e12);
      model.setUserInput(s.sheet, s.row, s.column, settings.decimal === 'comma' ? text.replace('.', ',') : text);
      const den = frac[4].length;
      model.updateRangeStyle({ sheet: s.sheet, row: s.row, column: s.column, width: 1, height: 1 }, 'num_fmt', den > 2 ? '# ???/???' : den > 1 ? '# ??/??' : '# ?/?');
    } else if (/^-?\d+\/\d+$/.test(value.trim()) && /\?\/[?\d]/.test(before) && Number(value.trim().split('/')[1]) > 0) {
      // В ячейке с дробным форматом «1/2» — число 0,5 (в «Общем» это была бы дата или текст, как в Excel).
      const [n, d] = value.trim().split('/').map(Number);
      const text = String(Math.round((n / d) * 1e12) / 1e12);
      model.setUserInput(s.sheet, s.row, s.column, settings.decimal === 'comma' ? text.replace('.', ',') : text);
    } else {
      const runs = richRuns(value);
      if (runs) model.setRichText(s.sheet, s.row, s.column, runs);
      else model.setUserInput(s.sheet, s.row, s.column, value);
    }
    // Текст в несколько строк (Alt+Enter) — сразу с переносом, как в Excel.
    if (!value.startsWith('=') && value.includes('\n') && !model.getCellStyle(s.sheet, s.row, s.column).style.alignment?.wrap_text) {
      model.updateRangeStyle({ sheet: s.sheet, row: s.row, column: s.column, width: 1, height: 1 }, 'alignment.wrap_text', 'true');
    }
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
  rich = null;
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
  rich = null;
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
  // Ctrl+B / I / U по выделенной части текста — оформление части (rich text).
  const rk = (e.ctrlKey || e.metaKey) && !e.altKey && !e.shiftKey ? ({ KeyB: 'font.b', KeyI: 'font.i', KeyU: 'font.u' } as Record<string, string>)[e.code] : undefined;
  if (rk) { lastEditInput = inp; if (richStyle(rk, '')) { e.preventDefault(); return; } }
  if (!['Shift', 'Control', 'Alt', 'Meta'].includes(e.key)) endPoint();
  if (e.key === 'Enter' && e.altKey) {
    // Alt+Enter — новая строка внутри ячейки (и в ячейке, и в строке формул), как в Excel.
    e.preventDefault();
    const at = inp.selectionStart ?? inp.value.length, end = inp.selectionEnd ?? at;
    inp.value = inp.value.slice(0, at) + '\n' + inp.value.slice(end);
    inp.setSelectionRange(at + 1, at + 1);
    (inp === editor ? formula : editor).value = inp.value;
    inp.dispatchEvent(new Event('input', { bubbles: true }));
    return;
  }
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

grid.beforeSelect = () => {
  if (editing) commitEdit(0, 0);
  return !editing; // формула с ошибкой — ввод продолжается
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
  if (now && (now.closest('#recovery') || now.closest('#context-menu') || now.closest('#context-submenu'))) return;
  commitEdit(0, 0);
}, 0));

formula.addEventListener('focus', () => { if (!editing) { startEdit(); formula.focus(); } showRefs(); });
formula.addEventListener('input', () => { endPoint(); editor.value = formula.value; showRefs(); });
formula.addEventListener('scroll', () => syncMirror());
formula.addEventListener('blur', () => setTimeout(() => {
  const now = document.activeElement;
  if (!editing || now === formula || now === editor || !document.hasFocus()) return;
  if (now && (now.closest('#recovery') || now.closest('#context-menu') || now.closest('#context-submenu') || now.closest('.sheet-tabs'))) return;
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

/** Перетащили цветную рамку ссылки: правим её в тексте формулы, знаки $ сохраняются. */
grid.onRefDrag = (id, rect) => {
  if (!editing) return;
  const inp = activeInput();
  const ref = refsIn(inp.value)[id];
  if (!ref) return;
  const corner = (part: string | undefined, row: number, column: number) => {
    const m = /(\$?)([A-Z]+)(\$?)(\d+)/i.exec(part ?? '');
    return `${m?.[1] ?? ''}${columnName(column)}${m?.[3] ?? ''}${row}`;
  };
  const parts = ref.text.slice(ref.text.lastIndexOf('!') + 1).split(':');
  const prefix = ref.text.slice(0, ref.text.lastIndexOf('!') + 1);
  const single = rect.r1 === rect.r2 && rect.c1 === rect.c2 && parts.length === 1;
  const next = prefix + corner(parts[0], rect.r1, rect.c1) + (single ? '' : `:${corner(parts[1] ?? parts[0], rect.r2, rect.c2)}`);
  inp.value = inp.value.slice(0, ref.start) + next + inp.value.slice(ref.end);
  inp.setSelectionRange(ref.start + next.length, ref.start + next.length);
  (inp === editor ? formula : editor).value = inp.value;
  inp.dispatchEvent(new Event('input', { bubbles: true }));
};

function jumpMark(dir: 1 | -1): void {
  const next = grid.nextMark(dir);
  if (!next) { showToast('Меток нет: сделайте двойной щелчок по номеру строки или букве столбца (Shift — красная, Ctrl — зелёная).'); return; }
  if (next.kind === 'row') grid.select(next.index, grid.sel.column);
  else grid.select(grid.sel.row, next.index);
}

/** Положение текста активной ячейки: 'superscript', 'subscript' или 'baseline'. */
const vertAlignOfStyle = (st: CellStyle) => (st.font as { vert_align?: string }).vert_align ?? 'baseline';
const vertAlignOf = () => vertAlignOfStyle(model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style);

/** Вставить знак в редактируемую ячейку (если ввода нет — начать его с текущего содержимого). */
function insertSymbol(text: string): void {
  if (!editing) startEdit(model.getCellContent(grid.sheet, grid.sel.row, grid.sel.column));
  const inp = activeInput();
  const at = inp.selectionStart ?? inp.value.length, end = inp.selectionEnd ?? at;
  inp.value = inp.value.slice(0, at) + text + inp.value.slice(end);
  inp.setSelectionRange(at + text.length, at + text.length);
  (inp === editor ? formula : editor).value = inp.value;
  inp.dispatchEvent(new Event('input', { bubbles: true }));
}

/** Вставить функцию из каталога: в редактируемую ячейку, а если ввода нет — в начало ввода формулы. */
function insertFunctionText(name: string): void {
  if (!editing) { startEdit(`=${name}(`); return; }
  const inp = activeInput();
  const at = inp.selectionStart ?? inp.value.length, end = inp.selectionEnd ?? at;
  const head = inp.value.startsWith('=') ? '' : '=';
  const text = `${head}${name}(`;
  inp.value = inp.value.slice(0, at) + text + inp.value.slice(end);
  inp.setSelectionRange(at + text.length, at + text.length);
  (inp === editor ? formula : editor).value = inp.value;
  inp.dispatchEvent(new Event('input', { bubbles: true }));
  inp.focus();
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

/** Переход к листу и области из журнала изменений: «A1:C5», «5:7» (строки), «B:D» (столбцы). */
function goToArea(sheet: string, area: string): void {
  const index = model.getWorksheetsProperties().findIndex((p) => p.name === sheet);
  if (index < 0) { showToast(`Листа «${sheet}» в книге уже нет.`); return; }
  if (index !== grid.sheet) setSheet(index);
  const rows = /^(\d+):(\d+)$/.exec(area);
  const cols = /^([A-Z]+):([A-Z]+)$/.exec(area);
  if (rows) grid.selectRange(Number(rows[1]), 1, Number(rows[2]), 16_384);
  else if (cols) grid.selectRange(1, columnNumber(cols[1]), 1_048_576, columnNumber(cols[2]));
  else {
    const r = refsIn(area)[0];
    if (r) goTo({ ...r, sheet: undefined });
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
  const colored = refs.map((r, index) => {
    const key = r.text.toUpperCase().replace(/\$/g, '');
    if (!colorOf.has(key)) colorOf.set(key, REF_COLORS[colorOf.size % REF_COLORS.length]);
    return { ...r, color: colorOf.get(key)!, index };
  });
  grid.refHighlights = colored
    .filter((r) => (r.sheet ? names.indexOf(r.sheet.toLowerCase()) : editCell.sheet) === grid.sheet)
    .map((r) => ({
      r1: Math.min(r.row, r.row2 ?? r.row), r2: Math.max(r.row, r.row2 ?? r.row),
      c1: Math.min(r.column, r.column2 ?? r.column), c2: Math.max(r.column, r.column2 ?? r.column), color: r.color, id: r.index,
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
  mirror.style.transform = `translateY(${-formula.scrollTop}px)`;
}

// --- Высота строки формул: потянуть за нижний край или ⌄ (как в Excel) ---
const fxBar = document.querySelector<HTMLElement>('.formula-bar')!;
const fxExpand = $<HTMLButtonElement>('formula-expand');
const FX_MIN = 26;
function setFormulaHeight(px: number, save = true): void {
  const h = Math.round(Math.max(FX_MIN, Math.min(px, window.innerHeight * 0.5)));
  fxBar.style.setProperty('--formula-h', `${h}px`);
  fxExpand.classList.toggle('open', h > FX_MIN);
  fxExpand.title = h > FX_MIN ? 'Свернуть строку формул (Ctrl+Shift+U)' : 'Развернуть строку формул (Ctrl+Shift+U)';
  if (save) { settings.formulaHeight = h; saveSettings(); }
  grid.invalidateLayout();
  if (!editing) placeEditor();
}
function toggleFormulaHeight(): void {
  const now = settings.formulaHeight > FX_MIN ? settings.formulaHeight : FX_MIN;
  setFormulaHeight(now > FX_MIN ? FX_MIN : Math.max(settings.formulaExpanded || 0, 18 * 4 + 8));
}
setFormulaHeight(settings.formulaHeight || FX_MIN, false);
fxExpand.addEventListener('mousedown', (e) => e.preventDefault());
fxExpand.addEventListener('click', toggleFormulaHeight);
$('formula-grip').addEventListener('mousedown', (e) => {
  e.preventDefault();
  const grip = e.currentTarget as HTMLElement;
  const startY = e.clientY, startH = settings.formulaHeight || FX_MIN;
  grip.classList.add('dragging');
  const move = (ev: MouseEvent) => setFormulaHeight(startH + ev.clientY - startY, false);
  const up = (ev: MouseEvent) => {
    window.removeEventListener('mousemove', move);
    window.removeEventListener('mouseup', up);
    grip.classList.remove('dragging');
    setFormulaHeight(startH + ev.clientY - startY);
    if (settings.formulaHeight > FX_MIN) { settings.formulaExpanded = settings.formulaHeight; saveSettings(); }
  };
  window.addEventListener('mousemove', move);
  window.addEventListener('mouseup', up);
});

const escapeHtml = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

// --- §4: формат, изменённый вводом, можно сразу вернуть ---

const chip = $('format-chip');
let chipTimer = 0;

function formatName(fmt: string): string {
  if (/[dmy]/i.test(fmt) && !/[#0]/.test(fmt)) return 'Дата';
  if (fmt.includes('%')) return 'Процентный';
  if (/[$€₽]/.test(fmt)) return 'Денежный';
  if (fmt === '@') return 'Текстовый';
  if (/\?\/[?\d]/.test(fmt)) return 'Дробный';
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
  const typed = nameBox.value.trim();
  const ref = refsIn(typed.toUpperCase())[0];
  // Адрес («B12», «Лист2!A1:B3») — переход; иначе имя: есть — переход, нет — выделение получает это имя.
  if (ref && ref.text === typed.toUpperCase().replace(/\$/g, '')) { goTo(ref); return; }
  const known = typed ? nameFormula(typed) : null;
  if (known) { const target = refsIn(known.toUpperCase())[0]; if (target) goTo(target); }
  else if (typed && !ref) defineFromSelection(typed);
  else if (ref) goTo(ref);
  updateFormulaBar();
  focusGrid();
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
  let count = 0, numbers = 0, sum = 0;
  // Целые строки и столбцы, как в Excel: считаем только по занятой области листа.
  const used = model.getSheetDimension(grid.sheet);
  const ranges = [grid.sel, ...grid.extraRanges];
  let cells = 0;
  const clipped = ranges.map((r) => {
    const c = { r1: r.r1, c1: r.c1, r2: Math.min(r.r2, Math.max(used[2], r.r1)), c2: Math.min(r.c2, Math.max(used[3], r.c1)) };
    cells += (c.r2 - c.r1 + 1) * (c.c2 - c.c1 + 1);
    return c;
  });
  const total = ranges.reduce((n, r) => n + (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1), 0);
  if (total > 1 && cells <= 2_000_000) {
    for (const s of clipped) {
      for (let c = s.c1; c <= s.c2; c++) {
        // Заполненные строки столбца — одним вызовом: пустые ячейки не опрашиваем.
        for (const r of model.getRowsWithData(grid.sheet, c)) {
          if (r < s.r1 || r > s.r2 || model.getRowHeight(grid.sheet, r) === 0) continue; // скрытые и отфильтрованные не считаем
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
  }
  const fmt = (v: number) => v.toLocaleString('ru-RU', { maximumFractionDigits: 10 });
  const f = filterAt(grid.sheet, grid.sel.row, grid.sel.column);
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

// --- Масштаб листа: Ctrl + колесо и кнопки в строке состояния ---

const zoomVal = $<HTMLButtonElement>('zoom-val');
grid.onZoom = (z) => { zoomVal.textContent = `${Math.round(z * 100)}%`; placeEditor(); };
$('zoom-in').addEventListener('click', () => grid.setZoom(grid.zoom * 1.1));
$('zoom-out').addEventListener('click', () => grid.setZoom(grid.zoom / 1.1));
zoomVal.addEventListener('click', () => grid.setZoom(1));
// Браузер (Chrome, «Муром») масштабирует всю страницу по Ctrl + колесо и Ctrl +/−: запрещаем,
// колесо над листом меняет только масштаб листа (grid.ts).
document.addEventListener('wheel', (e) => { if (e.ctrlKey) e.preventDefault(); }, { passive: false });
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && ['Equal', 'Minus', 'NumpadAdd', 'NumpadSubtract', 'Digit0', 'Numpad0'].includes(e.code) && !e.altKey) {
    // Ctrl+0 и прочие браузерные сочетания масштаба страницы не нужны; остальное обрабатывает приложение.
    if (e.code === 'Digit0' || e.code === 'Numpad0') return;
    e.preventDefault();
  }
}, true);

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
  grid.filters = filtersOf(grid.sheet).map((f) => ({ r1: f.r1, c1: f.c1, r2: f.r2, c2: f.c2, active: new Set(f.criteria.keys()) }));
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
  const more = document.createElement('button');
  more.className = 'swatch';
  more.textContent = '…';
  more.title = 'Другие цвета…';
  more.addEventListener('click', () => { menu.hidden = true; pickColor('#2B7A6B', (c) => { model.setSheetColor(index, c); renderSheetTabs(); }); });
  colors.append(more);
  const label = document.createElement('div');
  label.className = 'menu-label';
  label.textContent = 'Цвет ярлычка';
  menu.append(label, colors);
  placeMenu(menu, x, y, true);
}

interface MenuItem {
  label: string; run: () => void; disabled?: boolean; keys?: string; separator?: boolean;
  /** Значок слева (свой SVG) и отметка выбранного пункта. */
  icon?: string; checked?: boolean;
  /** Подменю: открывается по наведению, без щелчка. */
  sub?: MenuItem[];
}

/** Подменю: второе всплывающее меню рядом с пунктом. */
const subMenu = Object.assign(document.createElement('div'), { id: 'context-submenu', className: 'context-menu', hidden: true });
document.body.append(subMenu);

function fillMenu(menu: HTMLElement, items: MenuItem[], onPick: () => void): void {
  menu.replaceChildren();
  for (const it of items) {
    if (it.separator) menu.append(Object.assign(document.createElement('hr'), { className: 'menu-sep' }));
    const b = document.createElement('button');
    const ic = it.icon !== undefined ? `<i class="menu-icon">${it.icon}</i>` : '';
    b.innerHTML = `${ic}<span>${escapeHtml(it.label)}</span>${it.keys ? `<small>${escapeHtml(it.keys)}</small>` : ''}${it.sub ? '<small>▸</small>' : ''}`;
    if (it.checked) b.classList.add('checked');
    b.disabled = !!it.disabled;
    if (it.sub) {
      const open = () => {
        fillMenu(subMenu, it.sub!, onPick);
        const r = b.getBoundingClientRect();
        placeMenu(subMenu, r.right - 2, r.top, false);
      };
      b.addEventListener('mouseenter', open);
      b.addEventListener('click', open);
    } else {
      b.addEventListener('mouseenter', () => { if (menu !== subMenu) subMenu.hidden = true; });
      b.addEventListener('click', () => { onPick(); it.run(); });
    }
    menu.append(b);
  }
}

/** Контекстное меню: пункты; separator — черта перед пунктом. */
function showMenu(items: MenuItem[], x: number, y: number): HTMLElement {
  const menu = $('context-menu');
  subMenu.hidden = true;
  fillMenu(menu, items, () => { menu.hidden = true; subMenu.hidden = true; });
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
    { label: 'Специальная вставка', keys: 'Ctrl+Alt+V', run: () => {}, sub: pasteSpecialItems() },
  ];
  // Вырезанные целые строки (столбцы) — вставить между другими, как в Excel «Вставить вырезанные ячейки».
  const cut = clip?.cut && clip.sheet === grid.sheet ? clip.lines : undefined;
  if (cut?.kind === 'row' && header !== 'column' && !fullCols) {
    const n = cut.count > 1 ? `вырезанные строки (${cut.count})` : 'вырезанную строку';
    items.push({ label: `Вставить ${n} выше`, separator: true, run: () => insertCutLines(s.r1) });
    items.push({ label: `Вставить ${n} ниже`, run: () => insertCutLines(s.r2 + 1) });
  } else if (cut?.kind === 'column' && header !== 'row' && !fullRows) {
    const n = cut.count > 1 ? `вырезанные столбцы (${cut.count})` : 'вырезанный столбец';
    items.push({ label: `Вставить ${n} слева`, separator: true, run: () => insertCutLines(s.c1) });
    items.push({ label: `Вставить ${n} справа`, run: () => insertCutLines(s.c2 + 1) });
  }
  if (!header) {
    items.push({ label: 'Сортировка от А до Я', separator: true, run: () => sortBy(s.column, false) });
    items.push({ label: 'Сортировка от Я до А', run: () => sortBy(s.column, true) });
    items.push({ label: 'Настраиваемая сортировка…', run: () => sortBy(s.column, false, true) });
    items.push({ label: filterAt(grid.sheet, s.row, s.column) ? 'Снять фильтр' : 'Фильтр', run: toggleFilter });
  }
  if (header !== 'column' && !fullCols) {
    items.push({ label: rows > 1 ? `Вставить строки (${rows})` : 'Вставить строку', separator: true, run: () => structure(() => model.insertRows(grid.sheet, s.r1, rows)) });
  }
  if (header !== 'row' && !fullRows) {
    items.push({ label: cols > 1 ? `Вставить столбцы (${cols})` : 'Вставить столбец', separator: header !== 'column' && !fullCols ? false : true,
      run: () => structure(() => model.insertColumns(grid.sheet, s.c1, cols)) });
  }
  if (header !== 'column' && !fullCols) {
    items.push({ label: rows > 1 ? `Удалить строки (${rows})` : 'Удалить строку', run: () => void deleteLines('row') });
  }
  if (header !== 'row' && !fullRows) {
    items.push({ label: cols > 1 ? `Удалить столбцы (${cols})` : 'Удалить столбец', run: () => void deleteLines('column') });
  }
  items.push({ label: 'Очистить содержимое', keys: keysOf('edit.clear'), run: () => run('edit.clear') });
  if (!header) {
    const has = !!commentOf(s.row, s.column);
    items.push({ label: has ? 'Изменить примечание…' : 'Вставить примечание…', separator: true, run: () => editComment(s.row, s.column) });
    if (has) items.push({ label: 'Удалить примечание', run: () => deleteComment(s.row, s.column) });
  }
  if (!header) items.push({ label: 'Формат ячеек…', keys: keysOf('format.cells'), run: () => run('format.cells') });
  if (!header && (rows > 1 || cols > 1)) {
    items.push({ label: 'Объединить и поместить в центр', separator: true, run: () => merge('center') });
    items.push({ label: 'Объединить ячейки', run: () => merge('plain') });
    if (rows > 1) items.push({ label: 'Объединить по строкам', run: () => merge('across') });
  }
  if (!header && model.getMergedCells(grid.sheet).some((m) => m.row <= s.r2 && m.row + m.height - 1 >= s.r1 && m.column <= s.c2 && m.column + m.width - 1 >= s.c1)) {
    items.push({ label: 'Отменить объединение ячеек', run: () => merge('undo') });
  }
  if (header !== 'row') items.push({ label: 'Ширина столбца…', separator: true, run: () => askSize('column') });
  // Скрыть / показать: «Показать» открывает скрытые внутри выделения (выделите соседей скрытых).
  if (header === 'column') {
    items.push({ label: cols > 1 ? `Скрыть столбцы (${cols})` : 'Скрыть столбец', run: () => structure(() => model.setColumnsHidden(grid.sheet, s.c1, s.c2, true)) });
    items.push({ label: 'Показать скрытые столбцы', run: () => structure(() => model.setColumnsHidden(grid.sheet, s.c1, s.c2, false)) });
  }
  if (header === 'row') {
    items.push({ label: rows > 1 ? `Скрыть строки (${rows})` : 'Скрыть строку', separator: true, run: () => structure(() => model.setRowsHidden(grid.sheet, s.r1, s.r2, true)) });
    items.push({ label: 'Показать скрытые строки', run: () => structure(() => model.setRowsHidden(grid.sheet, s.r1, Math.min(s.r2, LAST_ROW), false)) });
  }
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
// --- Рисунки ---

const imageInput = Object.assign(document.createElement('input'), { type: 'file', accept: 'image/png,image/jpeg,image/gif', hidden: true });
document.body.appendChild(imageInput);
imageInput.addEventListener('change', async () => {
  const file = imageInput.files?.[0];
  imageInput.value = '';
  if (!file) return;
  const ext = file.type === 'image/png' ? 'png' : file.type === 'image/jpeg' ? 'jpeg' : file.type === 'image/gif' ? 'gif' : '';
  if (!ext) { showToast('Поддерживаются рисунки PNG, JPEG и GIF.'); return; }
  try {
    const bytes = new Uint8Array(await file.arrayBuffer());
    const bmp = await createImageBitmap(file);
    const k = Math.min(1, 480 / bmp.width, 360 / bmp.height);
    const index = model.addImage(grid.sheet, grid.sel.row, grid.sel.column, Math.max(8, Math.round(bmp.width * k)), Math.max(8, Math.round(bmp.height * k)), ext, file.name, bytes);
    bmp.close();
    grid.selObj = index;
    grid.render();
  } catch (e) { showToast(`Не удалось вставить рисунок: ${e}`); }
  focusGrid();
});
grid.onImagePlace = (kind, i, row, column, dx, dy, w, h) => {
  try {
    if (kind === 'image') model.placeImage(grid.sheet, i, row, column, dx, dy, w, h);
    else model.setCharts(grid.sheet, model.getCharts(grid.sheet).map((c, k) => (k === i ? { ...c, row, column, dx, dy, width: w, height: h } : c)));
  } catch (e) { showToast(String(e)); }
  grid.render();
};

function deleteSelectedObject(): void {
  const o = grid.objAt(grid.selObj);
  try {
    if (o?.kind === 'image') model.deleteImage(grid.sheet, o.index);
    else if (o) model.setCharts(grid.sheet, model.getCharts(grid.sheet).filter((_, k) => k !== o.index));
  } catch (e) { showToast(String(e)); }
  grid.selObj = -1;
  grid.render();
}

// --- Диаграммы ---

/** Окно диаграммы: тип, заголовок, диапазон, шаблон оформления (цвета, легенда, подписи); для существующей — ещё «Удалить». */
function chartDialog(existing: number | null, initial: { kind: number; title: string; range: string; look?: Partial<ChartLook> }): void {
  const overlay = $('recovery');
  const kinds = ['Столбцы', 'Линии', 'Круговая'];
  const templates = [...CHART_TEMPLATES, ...settings.chartTemplates];
  const look: ChartLook = {
    colors: initial.look?.colors?.length ? [...initial.look.colors] : [...CHART_TEMPLATES[0].colors],
    legend: initial.look?.legend ?? 0,
    labels: initial.look?.labels ?? false,
  };
  const isCustom = !!initial.look?.colors?.length;
  overlay.innerHTML = `<div class="modal"><h2>${existing === null ? 'Новая диаграмма' : 'Диаграмма'}</h2>
    <p>Первая строка диапазона — названия рядов, первый столбец — подписи. Круговая берёт первый ряд.</p>
    <p><label>Тип: <select data-f="kind">${kinds.map((k, i) => `<option value="${i}"${i === initial.kind ? ' selected' : ''}>${k}</option>`).join('')}</select></label></p>
    <p><label>Заголовок: <input data-f="title" style="width:60%"></label></p>
    <p><label>Данные: <input data-f="range" style="width:40%" placeholder="A1:C5"></label></p>
    <p><label>Шаблон: <select data-f="tpl"><option value="-1">${isCustom ? 'Свои цвета' : 'Классика'}</option>${templates.map((t, i) => `<option value="${i}">${t.name}</option>`).join('')}</select></label>
      <button data-act="save-tpl" title="Сохранить текущие цвета, легенду и подписи как свой шаблон">Сохранить как шаблон…</button></p>
    <p>Цвета: <span data-f="swatches"></span></p>
    <p><label>Легенда: <select data-f="legend"><option value="0">снизу</option><option value="1">справа</option><option value="2">нет</option></select></label>
      <label style="margin-left:14px"><input type="checkbox" data-f="labels"> Подписи значений</label></p>
    <div class="rec-actions">${existing === null ? '' : '<button data-act="delete" class="danger">Удалить</button>'}
    <button data-act="cancel">Отмена</button><button data-act="ok" class="primary">ОК</button></div></div>`;
  overlay.hidden = false;
  const box = overlay.querySelector<HTMLElement>('.modal')!;
  const f = (n: string) => box.querySelector<HTMLInputElement & HTMLSelectElement>(`[data-f="${n}"]`)!;
  f('title').value = initial.title;
  f('range').value = initial.range;
  const swatches = () => {
    const rg = parseRange(f('range').value.trim().toUpperCase());
    const kind = Number(f('kind').value);
    const count = rg ? Math.min(8, kind === 2 ? rg[2] - rg[0] : rg[3] - rg[1]) : 1;
    box.querySelector('[data-f="swatches"]')!.innerHTML = Array.from({ length: count }, (_, i) =>
      `<input type="color" data-color="${i}" value="${look.colors[i] ?? CHART_TEMPLATES[0].colors[i % 8]}" style="width:32px;height:24px;padding:0;border:1px solid #bbb">`).join(' ');
    f('legend').value = String(look.legend);
    f('labels').checked = look.labels;
  };
  swatches();
  f('range').addEventListener('change', swatches);
  f('kind').addEventListener('change', swatches);
  f('legend').addEventListener('change', () => { look.legend = Number(f('legend').value); });
  f('labels').addEventListener('change', () => { look.labels = f('labels').checked; });
  box.addEventListener('input', (e) => {
    const t = e.target as HTMLInputElement;
    if (t.dataset.color !== undefined) look.colors[Number(t.dataset.color)] = t.value;
  });
  f('tpl').addEventListener('change', () => {
    const t = templates[Number(f('tpl').value)];
    if (t) { look.colors = [...t.colors]; look.legend = t.legend; look.labels = t.labels; swatches(); }
  });
  const close = () => { overlay.hidden = true; focusGrid(); };
  box.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') close(); });
  box.addEventListener('click', (e) => {
    const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
    if (!act) return;
    if (act === 'save-tpl') {
      const name = prompt('Название шаблона:')?.trim();
      if (!name) return;
      const t = { name, colors: [...look.colors], legend: look.legend, labels: look.labels };
      settings.chartTemplates = [...settings.chartTemplates.filter((x) => x.name !== name), t];
      saveSettings();
      showToast(`Шаблон «${name}» сохранён.`);
      return;
    }
    try {
      const list = model.getCharts(grid.sheet);
      if (act === 'delete' && existing !== null) model.setCharts(grid.sheet, list.filter((_, k) => k !== existing));
      if (act === 'ok') {
        const c = { kind: Number(f('kind').value), title: f('title').value.trim(), range: f('range').value.trim().toUpperCase(), colors: look.colors, legend: look.legend, labels: look.labels };
        if (!parseRange(c.range)) { showToast('Укажите диапазон из двух и более строк и столбцов, например A1:C5.'); return; }
        if (existing === null) {
          const s = grid.sel;
          model.setCharts(grid.sheet, [...list, { ...c, row: s.row, column: s.column, dx: 0, dy: 0, width: 480, height: 288 }]);
          grid.selObj = grid.images.length + list.length;
        } else {
          model.setCharts(grid.sheet, list.map((x, k) => (k === existing ? { ...x, ...c } : x)));
        }
      }
    } catch (err) { showToast(String(err)); }
    grid.render();
    close();
  });
}

function insertChart(kind: number): void {
  const s = grid.sel;
  const single = s.r1 === s.r2 && s.c1 === s.c2;
  const [r1, c1, r2, c2] = single ? model.getSheetDimension(grid.sheet) : [s.r1, s.c1, s.r2, s.c2];
  chartDialog(null, { kind, title: '', range: `${columnName(c1)}${r1}:${columnName(c2)}${r2}` });
}

function editSelectedChart(): void {
  const o = grid.objAt(grid.selObj);
  if (o?.kind !== 'chart') { showToast('Выделите диаграмму щелчком мыши.'); return; }
  const c = model.getCharts(grid.sheet)[o.index];
  chartDialog(o.index, { kind: c.kind, title: c.title, range: c.range, look: c });
}
grid.onObjectDouble = () => { if (grid.objAt(grid.selObj)?.kind === 'chart') editSelectedChart(); };

// --- Защита листа, ячеек и структуры книги ---

// Отказы движка на защищённом листе (правка, вставка, форматирование и т. д.) — сообщением, а не ошибкой.
window.addEventListener('unhandledrejection', (e) => { if (/защищ/.test(String(e.reason))) { e.preventDefault(); showToast(String(e.reason).replace(/^Error: /, '')); } });
window.addEventListener('error', (e) => { if (/защищ/.test(e.message)) { e.preventDefault(); showToast(e.message.replace(/^(Uncaught )?Error: /, '')); } });


function toggleSheetProtection(): void {
  if (editing) commitEdit(0, 0);
  const on = !model.isSheetProtected(grid.sheet);
  try {
    model.setSheetProtected(grid.sheet, on);
    showToast(on ? 'Лист защищён: менять можно только незащищённые ячейки.' : 'Защита листа снята.');
  } catch (e) { showToast(String(e)); }
  grid.invalidateLayout(); updateFormulaBar(); refreshRibbon();
}

/** Окно «Защита ячеек»: защищаемая ячейка, скрыть формулы (действует после защиты листа). */
function cellProtectionDialog(): void {
  if (editing) commitEdit(0, 0);
  const s = grid.sel;
  const cur = model.getCellStyle(grid.sheet, s.row, s.column).style.protection ?? { locked: true, hidden: false };
  const overlay = $('recovery');
  overlay.innerHTML = `<div class="modal"><h2>Защита ячеек</h2>
    <p><label><input type="checkbox" data-f="locked"${cur.locked ? ' checked' : ''}> Защищаемая ячейка</label></p>
    <p><label><input type="checkbox" data-f="hidden"${cur.hidden ? ' checked' : ''}> Скрыть формулы</label></p>
    <p>Параметры начинают действовать после защиты листа (Рецензирование → Защитить лист).</p>
    <div class="rec-actions"><button data-act="cancel">Отмена</button><button data-act="ok" class="primary">ОК</button></div></div>`;
  overlay.hidden = false;
  const box = overlay.querySelector<HTMLElement>('.modal')!;
  const close = () => { overlay.hidden = true; focusGrid(); };
  box.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') close(); });
  box.addEventListener('click', (e) => {
    const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
    if (!act) return;
    if (act === 'ok') {
      const flag = (f: string) => String(box.querySelector<HTMLInputElement>(`[data-f="${f}"]`)!.checked);
      const mark = model.undoMark();
      try {
        model.updateRangeStyle(selectionArea(), 'protection.locked', flag('locked'));
        model.updateRangeStyle(selectionArea(), 'protection.hidden', flag('hidden'));
      } catch (err) { showToast(String(err)); }
      model.mergeUndoSince(mark);
      grid.render();
    }
    close();
  });
}

// --- Примечания и разрывы страниц ---

function commentOf(row: number, column: number): { ref: string; text: string; author: string } | undefined {
  const ref = `${columnName(column)}${row}`;
  try { return model.getComments(grid.sheet).find((c) => c.ref === ref); } catch { return undefined; }
}

function deleteComment(row: number, column: number): void {
  if (!commentOf(row, column)) return;
  model.deleteComment(grid.sheet, row, column);
  grid.render();
}

/** Окно примечания: создать, изменить или удалить (пустой текст). */
function editComment(row: number, column: number): void {
  const old = commentOf(row, column);
  const overlay = $('recovery');
  overlay.innerHTML = `<div class="modal"><h2>Примечание к ячейке ${columnName(column)}${row}</h2>
    <textarea class="comment-text" rows="6" style="width:100%;box-sizing:border-box;font:inherit"></textarea>
    <div class="rec-actions">${old ? '<button data-act="delete" class="danger">Удалить</button>' : ''}
    <button data-act="cancel">Отмена</button><button data-act="ok" class="primary">Сохранить</button></div></div>`;
  overlay.hidden = false;
  const box = overlay.querySelector<HTMLElement>('.modal')!;
  const area = box.querySelector<HTMLTextAreaElement>('textarea')!;
  area.value = old?.text ?? '';
  area.focus();
  const close = () => { overlay.hidden = true; focusGrid(); };
  box.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') close(); });
  box.addEventListener('click', (e) => {
    const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
    if (!act) return;
    if (act === 'ok') model.setComment(grid.sheet, row, column, area.value, old?.author || userName(''));
    if (act === 'delete') model.deleteComment(grid.sheet, row, column);
    if (act !== 'cancel') grid.render();
    close();
  });
}

/** Всплывающее примечание при наведении. */
const commentTip = Object.assign(document.createElement('div'), { className: 'comment-tip', hidden: true });
document.body.appendChild(commentTip);
let tipCell = '';
grid.onCellHover = (row, column, x, y) => {
  const c = row ? commentOf(row, column) : undefined;
  if (!c) { commentTip.hidden = true; tipCell = ''; return; }
  const key = `${row}:${column}`;
  if (key !== tipCell) {
    tipCell = key;
    commentTip.textContent = '';
    if (c.author) commentTip.append(Object.assign(document.createElement('b'), { textContent: `${c.author}:` }), document.createElement('br'));
    commentTip.append(c.text);
  }
  commentTip.style.left = `${Math.min(x + 14, window.innerWidth - 280)}px`;
  commentTip.style.top = `${Math.min(y + 14, window.innerHeight - 120)}px`;
  commentTip.hidden = false;
};
document.addEventListener('mouseleave', () => { commentTip.hidden = true; tipCell = ''; });

/** Разрыв страницы выше и левее активной ячейки (как в Excel). */
function pageBreak(on: boolean): void {
  const s = grid.sel;
  const mark = model.undoMark();
  if (s.r1 > 1) model.setPageBreak(grid.sheet, true, s.r1 - 1, on);
  if (s.c1 > 1) model.setPageBreak(grid.sheet, false, s.c1 - 1, on);
  model.mergeUndoSince(mark);
  grid.render();
}

function askSize(kind: 'column' | 'row'): void {
  const s = grid.sel;
  const current = kind === 'column' ? Math.round(model.getColumnWidth(grid.sheet, s.column))
    : Math.round(model.getRowHeight(grid.sheet, s.row) * 22 / 25);
  const answer = prompt(kind === 'column' ? 'Ширина столбца (в пикселях; 0 — скрыть):' : 'Высота строки (в пикселях; 0 — скрыть):', String(current));
  const px = Number(answer?.replace(',', '.'));
  if (!answer || !Number.isFinite(px) || px < 0 || px > 2000) return;
  setSize(kind, kind === 'column' ? s.c1 : s.r1, kind === 'column' ? s.c2 : s.r2, px);
  focusGrid();
}

function setSize(kind: 'column' | 'row', from: number, to: number, px: number): void {
  to = Math.min(to, kind === 'column' ? LAST_COLUMN : LAST_ROW);
  // Ноль — скрыть, как в Excel; другое значение заодно показывает скрытые.
  const hide = px < 1;
  const mark = model.undoMark();
  if (kind === 'column') {
    model.setColumnsHidden(grid.sheet, from, to, hide);
    if (!hide) model.setColumnsWidth(grid.sheet, from, to, px);
  } else {
    model.setRowsHidden(grid.sheet, from, to, hide);
    if (!hide) model.setRowsHeight(grid.sheet, from, to, Grid.rowPxToModel(px));
  }
  model.mergeUndoSince(mark);
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
  // Как Excel: объединённые ячейки в автоподгонке высоты не участвуют.
  const merged = model.getMergedCells(grid.sheet).filter((m) => m.row <= row && row <= m.row + m.height - 1);
  for (const c of model.getColumnsWithData(grid.sheet, row)) {
    if (merged.some((m) => c >= m.column && c <= m.column + m.width - 1)) continue;
    const st = model.getCellStyle(grid.sheet, row, c).style;
    const text = model.getFormattedCellValue(grid.sheet, row, c);
    const rot = st.alignment?.text_rotation ?? 0;
    if (rot && text) { need = Math.max(need, rotatedHeight(text, st, rot)); continue; }
    need = Math.max(need, st.alignment?.wrap_text && text
      ? grid.measureWrapped(text, st, grid.columnPx(c))
      : Math.round((st.font.sz ?? 11) * 1.33 + 8));
  }
  return Math.min(409, need);
}

/** Высоты, которые поставила сама подгонка: лист → строка → px. Остальные — заданы вручную. */
const autoHeights = new Map<number, Map<number, number>>();

/** Высота строки не менялась вручную: стандартная или ровно та, что поставила подгонка. */
function isAutoHeight(row: number, nowPx: number): boolean {
  return nowPx === 22 || autoHeights.get(grid.sheet)?.get(row) === nowPx;
}

/** Высота под повёрнутый текст (как рисует grid.drawText); 255 — буквы столбиком. */
function rotatedHeight(text: string, st: CellStyle, rot: number): number {
  const f = st.font;
  const px = Math.round((f.sz ?? 11) * 1.33);
  if (rot === 255) return Math.min(409, [...text.replace(/\n/g, '')].length * Math.round(px * 1.1) + 8);
  measure.font = `${f.i ? 'italic ' : ''}${f.b ? 'bold ' : ''}${px}px ${fontStack(f.name)}`;
  const tw = measure.measureText(text.replace(/\n/g, ' ')).width;
  const rad = ((rot <= 90 ? rot : rot - 90) * Math.PI) / 180;
  return Math.min(409, Math.ceil(tw * Math.sin(rad) + px * Math.cos(rad) + 8));
}

/** Высота строк с переносом текста — по содержимому, как в Excel; строки с ручной высотой не трогаем. */
function fitRows(r1: number, r2: number): void {
  if (r2 - r1 > 500) return;
  for (let r = r1; r <= r2; r++) {
    const wrapped = [...model.getColumnsWithData(grid.sheet, r)].some((c) => {
      const al = model.getCellStyle(grid.sheet, r, c).style.alignment;
      return al?.wrap_text || al?.text_rotation;
    });
    if (!wrapped) continue;
    const now = Math.round(model.getRowHeight(grid.sheet, r) * 22 / 25);
    if (!isAutoHeight(r, now)) continue;
    const need = fitRow(r);
    if (need === now) continue;
    setSize('row', r, r, need);
    let rows = autoHeights.get(grid.sheet);
    if (!rows) autoHeights.set(grid.sheet, rows = new Map());
    rows.set(r, need);
  }
}

/** Активная ячейка входит в объединение (кнопка «Объединить» нажата). */
function inMerge(): boolean {
  const s = grid.sel;
  return model.getMergedCells(grid.sheet).some((m) => m.row <= s.row && s.row < m.row + m.height && m.column <= s.column && s.column < m.column + m.width);
}

/** Объединение ячеек, как в Excel: если данных несколько — останется левое верхнее значение. */
/** Объединение — для каждого выделенного диапазона (Ctrl): каждый станет своей ячейкой. */
function merge(kind: 'center' | 'plain' | 'across' | 'undo'): void {
  forEachRange(() => mergeOne(kind));
}

function mergeOne(kind: 'center' | 'plain' | 'across' | 'undo'): void {
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
  void onTargets('Границы', () => model.setAreaWithBorder(selectionArea(), { item: { style, color: '#000000' }, type }))
    .catch((e) => showToast(String(e))).then(() => grid.render());
  grid.render();
  focusGrid();
}

/** Линия образца: тонкая, средняя, толстая, двойная, пунктир, точки. */
function borderLine(st: string, x1: number, y1: number, x2: number, y2: number): string {
  const w = st === 'thick' ? 3 : st === 'medium' ? 2 : 1.3;
  const dash = st === 'dashed' ? ' stroke-dasharray="3 2"' : st === 'dotted' ? ' stroke-dasharray="1.2 1.6"' : '';
  if (st === 'double') {
    const dx = y1 === y2 ? 0 : 1.3, dy = y1 === y2 ? 1.3 : 0;
    return `<path d="M${x1 - dx} ${y1 - dy}L${x2 - dx} ${y2 - dy}M${x1 + dx} ${y1 + dy}L${x2 + dx} ${y2 + dy}" stroke="currentColor" stroke-width="0.9"/>`;
  }
  return `<path d="M${x1} ${y1}L${x2} ${y2}" stroke="currentColor" stroke-width="${w}"${dash}/>`;
}

/** Значок «Границ»: какие границы и какой линией (на кнопке и в меню). */
function borderIcon(type: BorderType, st: string): string {
  const a = 3, b = 17, c = 10;
  const sides: Record<string, [number, number, number, number]> = {
    top: [a, a, b, a], bottom: [a, b, b, b], left: [a, a, a, b], right: [b, a, b, b], h: [a, c, b, c], v: [c, a, c, b],
  };
  const on: Record<string, string[]> = {
    All: ['top', 'bottom', 'left', 'right', 'h', 'v'], Outer: ['top', 'bottom', 'left', 'right'], Inner: ['h', 'v'],
    Top: ['top'], Bottom: ['bottom'], Left: ['left'], Right: ['right'], None: [],
  };
  const guide = Object.values(sides).map(([x1, y1, x2, y2]) => `M${x1} ${y1}L${x2} ${y2}`).join('');
  const lines = (on[type] ?? []).map((k) => borderLine(st, ...sides[k])).join('');
  return `<svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="${guide}" stroke="currentColor" stroke-width="0.7" stroke-dasharray="1 1.5" opacity=".45"/>${lines}</svg>`;
}

/** Кнопка «Границы» показывает текущий выбор. */
function syncBordersButton(): void {
  const b = document.querySelector<HTMLElement>('.btn[data-command="format.borders"]');
  if (!b) return;
  b.innerHTML = borderIcon(settings.borderType as BorderType, settings.borderStyle);
  const t = BORDER_TYPES.find((x) => x[0] === settings.borderType)?.[1] ?? '';
  const l = BORDER_STYLES.find((x) => x[0] === settings.borderStyle)?.[1] ?? '';
  b.title = `Границы: ${t.toLowerCase()}, линия ${l}`;
}

/** Стрелка у кнопки «Границы»: какие границы и какой линией. Выбор запоминается. */
function bordersMenu(x: number, y: number): void {
  showMenu([
    ...BORDER_TYPES.map(([type, label]) => ({
      label, icon: borderIcon(type, settings.borderStyle), checked: type === settings.borderType,
      separator: type === 'Bottom' || type === 'None',
      run: () => { settings.borderType = type; saveSettings(); syncBordersButton(); applyBorders(); },
    })),
    ...BORDER_STYLES.map(([st, label], i) => ({
      label: `Линия: ${label}`, icon: `<svg width="20" height="20" viewBox="0 0 20 20" fill="none">${borderLine(st, 2, 10, 18, 10)}</svg>`,
      checked: st === settings.borderStyle, separator: i === 0,
      run: () => { settings.borderStyle = st; saveSettings(); syncBordersButton(); showToast(`Линия границ: ${label}. Щелчок по кнопке «Границы» ставит «${BORDER_TYPES.find((b) => b[0] === settings.borderType)?.[1]}» этой линией.`); },
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
  for (let i = from; i <= Math.min(to, from + 500); i++) {
    const fit = kind === 'column' ? fitColumn(i) : fitRow(i);
    setSize(kind, i, i, fit);
    if (kind === 'row') {
      // Подгонка по двойному щелчку возвращает строку в «авто», как в Excel.
      let rows = autoHeights.get(grid.sheet);
      if (!rows) autoHeights.set(grid.sheet, rows = new Map());
      rows.set(i, fit);
    }
  }
};

// Маркер заполнения: формулы протягиваются со сдвигом ссылок, ряды
// продолжаются (1, 2 → 3, 4…; январь → февраль…), как в Excel.
grid.onFill = async (t) => {
  const s = grid.sel;
  const area = selectionArea();
  const rows = t.r1 !== s.r1 || t.r2 !== s.r2;
  // Куда тянем (без исходных ячеек): там могут быть скрытые строки/столбцы.
  const ext: Rect = rows
    ? (t.r1 < s.r1 ? { r1: t.r1, c1: s.c1, r2: s.r1 - 1, c2: s.c2 } : { r1: s.r2 + 1, c1: s.c1, r2: t.r2, c2: s.c2 })
    : (t.c1 < s.c1 ? { r1: s.r1, c1: t.c1, r2: s.r2, c2: s.c1 - 1 } : { r1: s.r1, c1: s.c2 + 1, r2: s.r2, c2: t.c2 });
  const info = hiddenIn(ext);
  let visibleOnly = false;
  if (info.rows || info.cols) {
    const c = await askHidden('Протягивание', info);
    if (!c) return;
    visibleOnly = c === 'visible';
  }
  // Только видимые: запоминаем скрытые ячейки и возвращаем их после протягивания.
  const keep: { r: number; c: number; content: string; style: CellStyle }[] = [];
  if (visibleOnly) {
    for (let r = ext.r1; r <= ext.r2; r++) {
      for (let c = ext.c1; c <= ext.c2; c++) {
        if (model.getRowHeight(grid.sheet, r) === 0 || model.getColumnWidth(grid.sheet, c) === 0) {
          keep.push({ r, c, content: model.getCellContent(grid.sheet, r, c), style: model.getCellStyle(grid.sheet, r, c).style });
        }
      }
    }
  }
  const mark = model.undoMark();
  try {
    if (rows) model.autoFillRows(area, t.r1 < s.r1 ? t.r1 : t.r2);
    else model.autoFillColumns(area, t.c1 < s.c1 ? t.c1 : t.c2);
    for (const k of keep) {
      if (k.content) model.setUserInput(grid.sheet, k.r, k.c, k.content);
      else model.rangeClearContents(grid.sheet, k.r, k.c, k.r, k.c);
      model.setSelectedCell(k.r, k.c);
      model.onPasteStyles([[k.style]]);
    }
  } catch (e) {
    showToast(`Не получилось протянуть: ${e}`);
  }
  model.mergeUndoSince(mark);
  grid.selectRange(t.r1, t.c1, t.r2, t.c2);
  updateFormulaBar();
  updateStatus();
};

/** Закрепить области: строки выше и столбцы левее активной ячейки. */
function freezeMenu(x: number, y: number): void {
  const s = grid.sel, f = grid.frozen;
  const pins = readPins();
  const d = model.getSheetDimension(grid.sheet);
  // Строки снизу: выделенные; если выделена одна ячейка — её строка (обычно «Итого»).
  const rows = { a: s.r1, b: Math.min(s.r2, Math.max(d[2], s.r1)) };
  const cols = { a: s.c1, b: Math.min(s.c2, Math.max(d[3], s.c1)) };
  const span = (k: 'row' | 'column', p: { a: number; b: number }) =>
    k === 'row' ? (p.a === p.b ? `строку ${p.a}` : `строки ${p.a}–${p.b}`) : (p.a === p.b ? `столбец ${columnName(p.a)}` : `столбцы ${columnName(p.a)}–${columnName(p.b)}`);
  showMenu([
    { label: f.rows || f.columns ? 'Снять закрепление сверху и слева' : `Закрепить области (выше и левее ${columnName(s.column)}${s.row})`,
      run: () => (f.rows || f.columns ? grid.setFrozen(0, 0) : grid.setFrozen(s.row - 1, s.column - 1)) },
    { label: 'Закрепить верхнюю строку', run: () => grid.setFrozen(1, 0) },
    { label: 'Закрепить первый столбец', run: () => grid.setFrozen(0, 1) },
    { label: `Закрепить снизу: ${span('row', rows)}`, separator: true, run: () => setPins(rows, pins.cols) },
    { label: `Закрепить справа: ${span('column', cols)}`, run: () => setPins(pins.rows, cols) },
    ...(pins.rows ? [{ label: `Снять закрепление снизу (${span('row', pins.rows)})`, run: () => setPins(null, pins.cols) }] : []),
    ...(pins.cols ? [{ label: `Снять закрепление справа (${span('column', pins.cols)})`, run: () => setPins(pins.rows, null) }] : []),
  ], x, y);
}

// --- Закрепление снизу и справа: хранится в книге именами _ENS_PinRows /
// _ENS_PinCols (ссылки сдвигаются при вставке и удалении строк). В Excel такого
// закрепления нет: там эти имена просто видны в «Диспетчере имён».

type Span = { a: number; b: number } | null;
const PIN_ROWS = '_ENS_PinRows', PIN_COLS = '_ENS_PinCols';

function readPins(): { rows: Span; cols: Span } {
  let rows: Span = null, cols: Span = null;
  try {
    for (const d of model.getDefinedNameList()) {
      if (d.scope !== grid.sheet) continue;
      const m = /\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\s*$/i.exec(d.formula);
      if (!m) continue;
      if (d.name === PIN_ROWS) rows = { a: Number(m[2]), b: Number(m[4]) };
      if (d.name === PIN_COLS) cols = { a: columnNumber(m[1]), b: columnNumber(m[3]) };
    }
  } catch { /* старый движок */ }
  return { rows, cols };
}

function setPins(rows: Span, cols: Span): void {
  const name = model.getWorksheetsProperties()[grid.sheet].name;
  const q = /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
  const put = (key: string, ref: string | null) => {
    const has = model.getDefinedNameList().some((d) => d.name === key && d.scope === grid.sheet);
    if (!ref) { if (has) model.deleteDefinedName(key, grid.sheet); return; }
    if (has) model.updateDefinedName(key, grid.sheet, key, grid.sheet, ref);
    else model.newDefinedName(key, grid.sheet, ref);
  };
  const mark = model.undoMark();
  try {
    put(PIN_ROWS, rows ? `${q}!$A$${rows.a}:$A$${rows.b}` : null);
    put(PIN_COLS, cols ? `${q}!$${columnName(cols.a)}$1:$${columnName(cols.b)}$1` : null);
  } catch (e) {
    showToast(`Не получилось: ${e}`);
  }
  model.mergeUndoSince(mark);
  syncPins();
  focusGrid();
}

function syncPins(): void {
  const p = readPins();
  grid.setPins(p.rows, p.cols);
}

let lastPointer = { x: 200, y: 160 };
document.addEventListener('pointerdown', (e) => { lastPointer = { x: e.clientX, y: e.clientY }; }, true);

// --- Отменить / Повторить: видно, можно ли, и что сработало ---

const qaSave = $<HTMLButtonElement>('qa-save');
const qaPrint = $<HTMLButtonElement>('qa-print');
qaSave.innerHTML = icon('save', 20);
qaPrint.innerHTML = icon('print', 20);
for (const [b, id] of [[qaSave, 'file.save'], [qaPrint, 'file.print']] as const) {
  b.addEventListener('mousedown', (e) => e.preventDefault());
  b.addEventListener('click', () => run(id));
}

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

const SCOPES: ['book' | 'sheet' | 'others', string][] = [
  ['book', 'Вся книга'], ['sheet', 'Только этот лист'], ['others', 'Все листы, кроме этого'],
];
const qaScope = $<HTMLButtonElement>('qa-scope');
function showScope(): void {
  const label = SCOPES.find(([k]) => k === settings.undoScope)![1];
  qaScope.textContent = settings.undoScope === 'book' ? '▾' : `▾ ${settings.undoScope === 'sheet' ? 'лист' : 'др.'}`;
  qaScope.title = `Что отменяют стрелки: ${label}`;
}
showScope();
qaScope.addEventListener('mousedown', (e) => e.preventDefault());
qaScope.addEventListener('click', () => {
  const r = qaScope.getBoundingClientRect();
  showMenu(SCOPES.map(([k, text]) => ({
    label: text, checked: settings.undoScope === k,
    run: () => { settings.undoScope = k; saveSettings(); showScope(); syncUndoState(); },
  })), r.left, r.bottom + 2);
});

/** 0 — вся книга, иначе режим отмены по листам. */
const scopeMode = () => (settings.undoScope === 'sheet' ? 0 : 1);

function undoRedo(kind: 'undo' | 'redo'): void {
  if (settings.undoScope !== 'book') {
    const code = kind === 'undo' ? model.undoScoped(grid.sheet, scopeMode()) : model.redoScoped(grid.sheet, scopeMode());
    if (code === 1) { showToast(kind === 'undo' ? 'Отменять нечего (в выбранной области).' : 'Повторять нечего (в выбранной области).'); return; }
    if (code === 2) {
      showToast('Дальше отмена по листам невозможна: следующее действие меняет строки, столбцы или листы целиком. Переключите область на «Вся книга».');
      return;
    }
    wink(kind);
    grid.invalidateLayout();
    updateFormulaBar();
    updateStatus();
    syncUndoState();
    return;
  }
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
  syncOutline();
  syncPins();
  const u = model.canUndo(); // с отменой по листам кнопка живая, пока есть что отменять: при препятствии объяснит причину
  const r = model.canRedo();
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
async function sortBy(column: number, descending: boolean, custom = false, target?: SheetFilter): Promise<void> {
  const s = grid.sel;
  // Таблица с фильтром: та, где активная ячейка (или та, чья кнопка нажата).
  const f = target ?? filterAt(grid.sheet, s.row, s.column);
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
    // При фильтре, как в Excel, переставляются только видимые строки.
    model.sortRangeBy(area, keys, BLANKS_CODE[request.blanks], !!f);
  } catch (e) {
    showToast(`Не удалось отсортировать: ${e}`);
  }
  grid.render();
  updateFormulaBar();
  focusGrid();
}

function toggleFilter(): void {
  const s = grid.sel;
  // В таблице с фильтром — снять его; иначе включить для таблицы вокруг выделения.
  const existing = filterAt(grid.sheet, s.row, s.column);
  if (existing) { removeFilter(grid.sheet, existing); return; }
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
  addFilter(grid.sheet, { r1: header, c1, r2: Math.max(region.r2, header), c2, criteria: new Map() });
}

initFilter({
  model: () => model,
  sheet: () => grid.sheet,
  sort: (column, descending, f) => sortBy(column, descending, false, f),
  changed: () => { syncFilter(); updateStatus(); },
});
/** Выделение для правил: целые строки и столбцы — до конца данных (но не меньше активной ячейки). */
function selectionRect(): Rect {
  const s = grid.sel, d = model.getSheetDimension(grid.sheet);
  return {
    r1: s.r1, c1: s.c1,
    r2: s.r2 === LAST_ROW ? Math.max(d[2], s.r1) : s.r2,
    c2: s.c2 === LAST_COLUMN ? Math.max(d[3], s.c1) : s.c2,
  };
}

initConditional({
  model: () => model,
  sheet: () => grid.sheet,
  sheetName: () => model.getWorksheetsProperties()[grid.sheet].name,
  selection: selectionRect,
  render: () => grid.render(),
  toast: showToast,
  focus: focusGrid,
});

/** Меню «Условное форматирование» — под кнопкой на ленте (или у указателя). */
function openConditionalMenu(): void {
  const b = document.querySelector<HTMLElement>('.btn[data-command="format.conditional"]');
  const r = b?.getBoundingClientRect();
  if (r && r.width) conditionalMenu(r.left, r.bottom + 2);
  else conditionalMenu(lastPointer.x, lastPointer.y + 8);
}

// После текущего щелчка: иначе тот же mousedown сразу закрыл бы окно.
grid.onFilterButton = (row, column, x, y) => setTimeout(() => openFilterPopup(row, column, x, y));

initNames({
  model: () => model,
  sheet: () => grid.sheet,
  selectionRef: () => {
    const s = grid.sel, name = model.getWorksheetsProperties()[grid.sheet].name;
    const q = /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
    const cell = (r: number, c: number) => `$${columnName(c)}$${r}`;
    return `${q}!${cell(s.r1, s.c1)}${s.r1 === s.r2 && s.c1 === s.c2 ? '' : `:${cell(s.r2, s.c2)}`}`;
  },
  goTo: (formula) => { const r = refsIn(formula.toUpperCase())[0]; if (r) goTo(r); },
  changed: () => { grid.render(); updateFormulaBar(); updateStatus(); },
  toast: showToast,
});

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
  const items = pasteSpecialItems();
  if (items) showMenu(items, x, y);
}

/** Пункты специальной вставки; без скопированного в ENS-Table — одна подсказка. */
function pasteSpecialItems(): MenuItem[] {
  const c = clip;
  if (!c) {
    return [{ label: 'Сначала скопируйте ячейки (Ctrl+C)', disabled: true, run: () => {} }];
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
  // Формулы: как «Вставить всё», но оформление ячеек остаётся прежним.
  const area = () => ({ r: grid.sel.r1, c: grid.sel.c1 });
  const keepFormats = (f: () => void) => {
    const { r, c: col } = area();
    const kept = [];
    for (let i = 0; i < r2 - r1 + 1; i++) {
      const row = [];
      for (let j = 0; j < c2 - c1 + 1; j++) row.push(model.getCellStyle(grid.sheet, r + i, col + j).style);
      kept.push(row);
    }
    f();
    model.setSelectedCell(r, col);
    model.setSelectedRange(r, col, r, col);
    model.onPasteStyles(kept);
  };
  return [
    { label: 'Вставить всё', keys: 'Ctrl+V', run: () => paste(c.text) },
    { label: 'Формулы', run: () => run2(() => keepFormats(() => { at(); model.pasteFromClipboard(c.sheet, c.range, c.data, false); }), r2 - r1 + 1, c2 - c1 + 1) },
    { label: 'Текст формулы', run: () => run2(() => {
      const { r, c: col } = area();
      for (let i = 0; i <= r2 - r1; i++) {
        for (let j = 0; j <= c2 - c1; j++) {
          const text = model.getCellContent(c.sheet, r1 + i, c1 + j);
          model.setUserInput(grid.sheet, r + i, col + j, text.startsWith('=') ? `'${text}` : text);
        }
      }
    }, r2 - r1 + 1, c2 - c1 + 1) },
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
  ];
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
  if (!menu.hidden && !menu.contains(e.target as Node) && !subMenu.contains(e.target as Node)) { menu.hidden = true; subMenu.hidden = true; }
});
document.addEventListener('keydown', (e) => {
  const menu = $('context-menu');
  if (e.key === 'Escape' && !menu.hidden) { menu.hidden = true; subMenu.hidden = true; e.stopPropagation(); }
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
  grid.clearMarks();
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

interface Clip {
  data: unknown; sheet: number; range: [number, number, number, number]; text: string; cut: boolean;
  /** Вырезаны целые строки или столбцы: их можно вставить между другими. */
  lines?: { kind: 'row' | 'column'; start: number; count: number };
}
let clip: Clip | null = null;

/** Скопировать выделение: во внутренний буфер и текстом — в системный. */
function takeClip(cut: boolean): string {
  const s = grid.sel;
  model.setSelectedCell(s.r1, s.c1);
  model.setSelectedRange(s.r1, s.c1, s.r2, s.c2);
  const c = model.copyToClipboard();
  const text = toTsv();
  const fullRows = s.c1 === 1 && s.c2 === LAST_COLUMN, fullCols = s.r1 === 1 && s.r2 === LAST_ROW;
  const lines = !cut || fullRows === fullCols ? undefined
    : fullRows ? { kind: 'row' as const, start: s.r1, count: s.r2 - s.r1 + 1 }
    : { kind: 'column' as const, start: s.c1, count: s.c2 - s.c1 + 1 };
  clip = { data: c.data, sheet: grid.sheet, range: c.range, text, cut, lines };
  grid.clipRange = { r1: s.r1, c1: s.c1, r2: Math.min(s.r2, c.range[2]), c2: Math.min(s.c2, c.range[3]), sheet: grid.sheet };
  grid.render();
  return text;
}

/** «Вставить вырезанные строки (столбцы)»: перед линией `before`; остальные сдвигаются,
 *  формулы следуют за ячейками, высота, ширина и скрытие переезжают вместе с ними. */
function insertCutLines(before: number): void {
  const L = clip?.lines;
  if (!clip || !L || clip.sheet !== grid.sheet) return;
  const end = L.start + L.count - 1;
  if (before >= L.start && before <= end + 1) { showToast('Вырезанное уже стоит здесь.'); return; }
  const delta = before > end ? before - L.start - L.count : before - L.start;
  let ok = true;
  structure(() => {
    try {
      if (L.kind === 'row') model.moveRowsExact(grid.sheet, L.start, L.count, delta);
      else model.moveColumnsExact(grid.sheet, L.start, L.count, delta);
    } catch (e) { ok = false; throw e; }
  });
  dropClip();
  if (ok) grid.selectLines(L.kind, L.start + delta, L.start + delta + L.count - 1);
}

function dropClip(): void {
  clip = null;
  grid.clipRange = null;
  grid.render();
}

const sameText = (a: string, b: string) => a.replace(/\r\n?/g, '\n').trim() === b.replace(/\r\n?/g, '\n').trim();

/** Вставка: своё — через движок (формулы, форматы), чужое — как текст. */
function paste(text: string): void {
  if (grid.extraRanges.length) {
    // Несколько диапазонов (Ctrl): скопированное — в каждый; вырезанное — только в одно место.
    if (clip?.cut) { showToast('Вырезанное вставляется в одно место: выделите один диапазон.'); return; }
    const extra = grid.extraRanges, keep = grid.sel;
    forEachRange(() => pasteHere(text));
    grid.sel = keep;
    grid.extraRanges = extra;
    grid.render();
    return;
  }
  pasteHere(text);
}

function pasteHere(text: string): void {
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
  e.clipboardData?.setData('text/plain', visibleCopyText() ?? takeClip(false));
});

/** Копирование с отобранным фильтром (или «только видимые» в настройке): только видимые
 *  строки и столбцы, значениями, как в Excel. Вопрос задать нельзя (буфер ждать не будет),
 *  поэтому при «спрашивать» — как в Excel. null — копировать как обычно. */
function visibleCopyText(): string | null {
  const s = grid.sel;
  if (grid.extraRanges.length) return null;
  const info = hiddenIn({ r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 });
  if (!info.rows && !info.cols) return null;
  const visible = settings.hiddenPolicy === 'visible' || (settings.hiddenPolicy !== 'all' && info.filtered);
  if (!visible) return null;
  const parts = visibleParts({ r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 });
  const rowsSet = [...new Set(parts.flatMap((p) => Array.from({ length: p.r2 - p.r1 + 1 }, (_, i) => p.r1 + i)))].sort((a, b) => a - b);
  const colsSet = [...new Set(parts.flatMap((p) => Array.from({ length: p.c2 - p.c1 + 1 }, (_, i) => p.c1 + i)))].sort((a, b) => a - b);
  const text = rowsSet.map((r) => colsSet.map((c) => model.getFormattedCellValue(grid.sheet, r, c)).join('\t')).join('\n');
  dropClip();
  showToast(`Скопированы только видимые строки (${rowsSet.length}) — значениями, как в Excel. Изменить: «Вид → Скрытое».`);
  return text;
}
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
  if (editing && richStyle(path, value(false))) return;
  const s = grid.sel;
  const current = flag ? flag(model.getCellStyle(grid.sheet, s.row, s.column).style) : false;
  const v = value(current);
  void onTargets('Оформление', () => model.updateRangeStyle(selectionArea(), path, v)).then(() => {
    grid.render();
    refreshRibbon();
    syncNumberFormat();
  });
}


// --- Группировка строк и столбцов (структура), как в Excel ---
// Уровни хранит движок (outlineLevel в xlsx); «свернуть» — скрыть строки группы
// и поставить метку «свёрнуто» у итоговой строки (под группой).

let outlineKey = '';
function syncOutline(force = false): void {
  let rows: ArrayLike<number> = [], cols: ArrayLike<number> = [];
  try { rows = model.getRowOutline(grid.sheet); cols = model.getColumnOutline(grid.sheet); } catch { /* старый движок */ }
  const [below, right] = outlineSummary();
  const key = `${grid.sheet}|${below}${right}|${Array.from(rows).join()}|${Array.from(cols).join()}`;
  if (!force && key === outlineKey) return;
  outlineKey = key;
  grid.setOutline(rows, cols, below === 1, right === 1);
}

/** Где «+/−» групп листа: [под строками, справа от столбцов], 1 — да (по умолчанию — сверху и слева). */
function outlineSummary(): [number, number] {
  try { const a = model.getOutlineSummary(grid.sheet); return [a[0], a[1]]; } catch { return [0, 0]; }
}

/** Итоговая строка/столбец группы a..b — та, где стоит «+/−» и метка «свёрнуто». */
function summaryIndex(kind: 'row' | 'column', a: number, b: number): number {
  const after = outlineSummary()[kind === 'row' ? 0 : 1] === 1;
  return after || a <= 1 ? b + 1 : a - 1;
}

function summaryMenu(x: number, y: number): void {
  const [below, right] = outlineSummary();
  const set = (b: number, r: number) => {
    const mark = model.undoMark();
    const moved = [...grid.rowGroups, ...grid.colGroups].filter((g) => g.collapsed);
    model.setOutlineSummary(grid.sheet, b === 1, r === 1);
    // Метка «свёрнуто» переезжает к новой итоговой строке (столбцу).
    for (const g of moved) {
      const now = summaryIndex(g.kind, g.a, g.b);
      if (now === g.s) continue;
      setLinesCollapsed(g.kind, g.s, false);
      setLinesCollapsed(g.kind, now, true);
    }
    model.mergeUndoSince(mark);
    syncOutline(true);
    focusGrid();
  };
  showMenu([
    { label: 'Итоги строк под данными', checked: below === 1, run: () => set(below === 1 ? 0 : 1, right) },
    { label: 'Итоги столбцов справа от данных', checked: right === 1, run: () => set(below, right === 1 ? 0 : 1) },
  ], x, y);
}

function outlineMap(kind: 'row' | 'column'): Map<number, { level: number; collapsed: boolean }> {
  const a = kind === 'row' ? model.getRowOutline(grid.sheet) : model.getColumnOutline(grid.sheet);
  const m = new Map<number, { level: number; collapsed: boolean }>();
  for (let i = 0; i + 2 < a.length; i += 3) m.set(a[i], { level: a[i + 1], collapsed: a[i + 2] === 1 });
  return m;
}

const hideLines = (kind: 'row' | 'column', a: number, b: number, hidden: boolean) =>
  kind === 'row' ? model.setRowsHidden(grid.sheet, a, b, hidden) : model.setColumnsHidden(grid.sheet, a, b, hidden);
const setLinesCollapsed = (kind: 'row' | 'column', i: number, c: boolean) =>
  kind === 'row' ? model.setRowCollapsed(grid.sheet, i, c) : model.setColumnCollapsed(grid.sheet, i, c);

/** Что группировать: выделены столбцы целиком — столбцы, иначе строки. */
function outlineTarget(): { kind: 'row' | 'column'; a: number; b: number } {
  const s = grid.sel;
  if (s.r1 === 1 && s.r2 === LAST_ROW && !(s.c1 === 1 && s.c2 === LAST_COLUMN)) return { kind: 'column', a: s.c1, b: s.c2 };
  return { kind: 'row', a: s.r1, b: Math.min(s.r2, LAST_ROW - 1) };
}

/** Сгруппировать (delta = 1) или разгруппировать (−1) строки/столбцы a..b. */
function changeOutline(delta: 1 | -1): void {
  const { kind, a, b } = outlineTarget();
  if (b - a > 100_000) { showToast('Слишком большой диапазон для группировки.'); return; }
  const levels = outlineMap(kind);
  const mark = model.undoMark();
  try {
    if (delta < 0 && levels.get(summaryIndex(kind, a, b))?.collapsed) {
      // Разгруппировать свёрнутую группу — строки снова видны.
      hideLines(kind, a, b, false);
      setLinesCollapsed(kind, summaryIndex(kind, a, b), false);
    }
    for (let i = a; i <= b;) {
      const lv = levels.get(i)?.level ?? 0;
      let j = i;
      while (j + 1 <= b && (levels.get(j + 1)?.level ?? 0) === lv) j++;
      const nl = Math.max(0, Math.min(7, lv + delta));
      if (nl !== lv) {
        if (kind === 'row') model.setRowsOutline(grid.sheet, i, j, nl);
        else model.setColumnsOutline(grid.sheet, i, j, nl);
      }
      i = j + 1;
    }
    if (delta > 0 && [...levels.values()].every((x) => x.level < 7) === false) showToast('Больше 7 уровней групп нельзя (как в Excel).');
  } catch (e) {
    showToast(`Не получилось: ${e}`);
  }
  model.mergeUndoSince(mark);
  grid.invalidateLayout();
  syncOutline(true);
  focusGrid();
}

/** «+/−»: свернуть или развернуть группу (вложенные свёрнутые остаются свёрнутыми). */
function toggleGroup(g: OutlineGroup, collapse: boolean): void {
  const mark = model.undoMark();
  try {
    if (collapse) {
      hideLines(g.kind, g.a, g.b, true);
      setLinesCollapsed(g.kind, g.s, true);
    } else {
      hideLines(g.kind, g.a, g.b, false);
      setLinesCollapsed(g.kind, g.s, false);
      const groups = g.kind === 'row' ? grid.rowGroups : grid.colGroups;
      for (const n of groups) {
        if (n.level > g.level && n.a >= g.a && n.b <= g.b && n.collapsed) hideLines(g.kind, n.a, n.b, true);
      }
    }
  } catch (e) {
    showToast(`Не получилось: ${e}`);
  }
  model.mergeUndoSince(mark);
  grid.invalidateLayout();
  syncOutline(true);
}

/** Кнопки уровней «1 2 3»: видны уровни меньше n, группы уровня n и глубже — свёрнуты. */
function showOutlineLevel(kind: 'row' | 'column', n: number): void {
  const groups = kind === 'row' ? grid.rowGroups : grid.colGroups;
  const mark = model.undoMark();
  try {
    for (const g of groups) hideLines(kind, g.a, g.b, false);
    for (const g of groups) setLinesCollapsed(kind, g.s, g.level >= n);
    for (const g of groups) if (g.level >= n) hideLines(kind, g.a, g.b, true);
  } catch (e) {
    showToast(`Не получилось: ${e}`);
  }
  model.mergeUndoSince(mark);
  grid.invalidateLayout();
  syncOutline(true);
}

grid.onOutlineToggle = (g, collapse) => toggleGroup(g, collapse);
grid.onOutlineLevel = (kind, n) => showOutlineLevel(kind, n);

// --- Скрытые строки и столбцы: что делать, когда действие их задевает ---
// По умолчанию — как в Excel (скрытые фильтром не трогаем, свёрнутые и
// скрытые вручную — вместе со всеми), но с вопросом; выбор можно запомнить.

type HiddenChoice = 'all' | 'visible';

function hiddenIn(r: Rect): { rows: number; cols: number; filtered: boolean } {
  const d = model.getSheetDimension(grid.sheet);
  const r2 = Math.min(r.r2, Math.max(d[2], r.r1)), c2 = Math.min(r.c2, Math.max(d[3], r.c1));
  let rows = 0, cols = 0, filtered = false;
  for (let i = r.r1; i <= Math.min(r2, r.r1 + 200_000); i++) {
    if (model.getRowHeight(grid.sheet, i) === 0) {
      rows++;
      const f = filterAt(grid.sheet, i, r.c1);
      if (f && f.criteria.size) filtered = true;
    }
  }
  for (let j = r.c1; j <= Math.min(c2, r.c1 + 16_384); j++) if (model.getColumnWidth(grid.sheet, j) === 0) cols++;
  return { rows, cols, filtered };
}

/** Видимые части прямоугольника: полосы видимых строк × полосы видимых столбцов. */
function visibleParts(r: Rect): Rect[] {
  const d = model.getSheetDimension(grid.sheet);
  const r2 = Math.min(r.r2, Math.max(d[2], r.r1)), c2 = Math.min(r.c2, Math.max(d[3], r.c1));
  const runs = (a: number, b: number, hidden: (i: number) => boolean) => {
    const out: [number, number][] = [];
    let start = 0;
    for (let i = a; i <= b + 1; i++) {
      const vis = i <= b && !hidden(i);
      if (vis && !start) start = i;
      if (!vis && start) { out.push([start, i - 1]); start = 0; }
    }
    return out;
  };
  const rr = runs(r.r1, r2, (i) => model.getRowHeight(grid.sheet, i) === 0);
  const cc = runs(r.c1, c2, (j) => model.getColumnWidth(grid.sheet, j) === 0);
  return rr.flatMap(([a, b]) => cc.map(([x, y]) => ({ r1: a, c1: x, r2: b, c2: y })));
}

function askHidden(what: string, info: { rows: number; cols: number; filtered: boolean }): Promise<HiddenChoice | null> {
  const excel: HiddenChoice = info.filtered ? 'visible' : 'all';
  if (settings.hiddenPolicy === 'all' || settings.hiddenPolicy === 'visible') return Promise.resolve(settings.hiddenPolicy);
  if (settings.hiddenPolicy === 'excel') return Promise.resolve(excel);
  const parts = [info.rows ? `строк: ${info.rows}` : '', info.cols ? `столбцов: ${info.cols}` : ''].filter(Boolean).join(', ');
  const o = $('recovery');
  o.innerHTML = `<div class="modal small hidden-ask">
    <h2>${escapeHtml(what)}: есть скрытое</h2>
    <p>В выделении скрыто ${parts}${info.filtered ? ' (часть — фильтром)' : ''}.
       Что делать с ними? Как в Excel — «${excel === 'all' ? 'вместе со скрытыми' : 'только видимые'}».</p>
    <label class="sort-choice"><input type="checkbox" data-remember> Запомнить выбор (изменить можно: «Вид → Скрытое»)</label>
    <div class="rec-actions">
      <button data-act="all" class="${excel === 'all' ? 'primary' : ''}">Вместе со скрытыми</button>
      <button data-act="visible" class="${excel === 'visible' ? 'primary' : ''}">Только видимые</button>
      <button data-act="cancel">Отмена</button>
    </div></div>`;
  o.hidden = false;
  const m = o.querySelector<HTMLElement>('.modal')!;
  return new Promise((resolve) => {
    const done = (v: HiddenChoice | null) => {
      if (v && m.querySelector<HTMLInputElement>('[data-remember]')!.checked) { settings.hiddenPolicy = v; saveSettings(); refreshRibbon(); }
      o.hidden = true;
      resolve(v);
      focusGrid();
    };
    m.addEventListener('click', (e) => {
      const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
      if (act === 'all' || act === 'visible') done(act);
      else if (act === 'cancel') done(null);
    });
    m.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') done(null);
      if (e.key === 'Enter') { e.preventDefault(); done(excel); }
    });
    m.querySelector<HTMLButtonElement>(`[data-act="${excel}"]`)!.focus();
  });
}

/** Действие на выделение с учётом скрытого (и выделения с Ctrl): fn — для одного диапазона. */
async function onTargets(what: string, fn: () => void): Promise<void> {
  const keepSel = grid.sel, keepExtra = grid.extraRanges;
  const ranges: Rect[] = [...keepExtra, { r1: keepSel.r1, c1: keepSel.c1, r2: keepSel.r2, c2: keepSel.c2 }];
  const info = ranges.map(hiddenIn).reduce((x, y) => ({ rows: x.rows + y.rows, cols: x.cols + y.cols, filtered: x.filtered || y.filtered }));
  let choice: HiddenChoice = 'all';
  if (info.rows || info.cols) {
    const c = await askHidden(what, info);
    if (!c) return;
    choice = c;
  }
  if (choice === 'all') { forEachRange(fn); return; }
  const parts = ranges.flatMap(visibleParts);
  if (!parts.length) return;
  const last = parts[parts.length - 1];
  grid.extraRanges = parts.slice(0, -1);
  grid.sel = { row: last.r1, column: last.c1, ...last };
  try {
    forEachRange(fn);
  } finally {
    grid.sel = keepSel;
    grid.extraRanges = keepExtra;
  }
  grid.render();
}

/** Удалить строки (столбцы) выделения — со скрытыми или только видимые. */
async function deleteLines(kind: 'row' | 'column'): Promise<void> {
  const s = grid.sel;
  const rect: Rect = kind === 'row' ? { r1: s.r1, c1: 1, r2: s.r2, c2: 1 } : { r1: 1, c1: s.c1, r2: 1, c2: s.c2 };
  const info = kind === 'row' ? { ...hiddenIn({ ...rect, c2: 1 }), cols: 0 } : { ...hiddenIn({ ...rect, r2: 1 }), rows: 0 };
  let choice: HiddenChoice = 'all';
  if (info.rows || info.cols) {
    const c = await askHidden(kind === 'row' ? 'Удаление строк' : 'Удаление столбцов', info);
    if (!c) return;
    choice = c;
  }
  const mark = model.undoMark();
  structure(() => {
    if (choice === 'all') {
      if (kind === 'row') model.deleteRows(grid.sheet, s.r1, s.r2 - s.r1 + 1);
      else model.deleteColumns(grid.sheet, s.c1, s.c2 - s.c1 + 1);
      return;
    }
    // Только видимые: полосы с конца, чтобы номера впереди не сдвигались.
    const parts = visibleParts(kind === 'row' ? { r1: s.r1, c1: 1, r2: s.r2, c2: 1 } : { r1: 1, c1: s.c1, r2: 1, c2: s.c2 });
    for (const p of parts.reverse()) {
      if (kind === 'row') model.deleteRows(grid.sheet, p.r1, p.r2 - p.r1 + 1);
      else model.deleteColumns(grid.sheet, p.c1, p.c2 - p.c1 + 1);
    }
  });
  model.mergeUndoSince(mark);
  syncOutline(true);
}

// --- Направление текста (поворот), как в Excel ---
const ROTATIONS: [number, string][] = [
  [0, 'Обычный текст'], [45, 'Против часовой стрелки (45°)'], [135, 'По часовой стрелке (−45°)'],
  [255, 'Вертикальный текст (буквы столбиком)'], [90, 'Повернуть текст вверх (90°)'], [180, 'Повернуть текст вниз (−90°)'],
];

function rotationMenu(x: number, y: number): void {
  const cur = model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.text_rotation ?? 0;
  showMenu([
    ...ROTATIONS.map(([v, label]) => ({
      label, checked: v === cur,
      run: () => { style('alignment.text_rotation', () => String(v)); fitRows(grid.sel.r1, grid.sel.r2); },
    })),
    {
      label: 'Свой угол…', separator: true,
      run: () => {
        const a = prompt('Угол поворота текста, градусы (от −90 до 90; плюс — против часовой стрелки):', '30');
        const n = Math.round(Number(a?.replace(',', '.')));
        if (a === null || !Number.isFinite(n) || n < -90 || n > 90) return;
        style('alignment.text_rotation', () => String(n >= 0 ? n : 90 - n));
        fitRows(grid.sel.r1, grid.sel.r2);
      },
    },
  ], x, y);
}

/** Кнопки «Разряды», «Разрядность»: новый код от текущего формата активной ячейки — на всё выделение. */
function numFormat(change: (code: string, shown: string) => string | null, what: string): void {
  const s = grid.sel;
  const code = model.getCellStyle(grid.sheet, s.row, s.column).style.num_fmt;
  const next = change(code, model.getFormattedCellValue(grid.sheet, s.row, s.column));
  if (next === null) { showToast(`${what}: для дат, времени и текста не меняется.`); return; }
  style('num_fmt', () => next);
}

/** «Формат ячеек» (Ctrl+1): образец считает движок по значению активной ячейки. */
function openFormatCells(): void {
  if (editing) commitEdit(0, 0);
  const s = grid.sel;
  void formatCellsDialog({
    model: () => model,
    sheet: () => grid.sheet,
    active: () => ({ row: s.row, column: s.column }),
    apply: (code) => style('num_fmt', () => code),
  }, model.getCellStyle(grid.sheet, s.row, s.column).style.num_fmt).then(() => focusGrid());
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
    if (grid.selObj >= 0) { deleteSelectedObject(); return; }
    void onTargets('Очистка', () => { const s = grid.sel; model.rangeClearContents(grid.sheet, s.r1, s.c1, s.r2, s.c2); })
      .then(() => { grid.render(); updateFormulaBar(); });
  } },
  { id: 'edit.autosum', label: 'Автосумма', keywords: ['sum', 'сумма', 'сумм', 'итог'], keys: ['Alt+Equal'], run: () => void autosum() },
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
  ...([
    ['fn.finance', 'Финансовые', ['Финансовые']], ['fn.logic', 'Логические', ['Логические']], ['fn.text', 'Текстовые', ['Текстовые']],
    ['fn.date', 'Дата и время', ['Дата и время']], ['fn.lookup', 'Ссылки и массивы', ['Ссылки и массивы']], ['fn.math', 'Математические', ['Математические']],
    ['fn.other', 'Другие функции', ['Статистические', 'Инженерные', 'Информационные', 'Совместимость', 'Работа с базой данных']],
  ] as [string, string, string[]][]).map(([id, label, cats]) => ({
    id, label: `Функции: ${label}`, keywords: ['функции', 'библиотека функций', label.toLowerCase()],
    run: () => {
      const b = document.querySelector<HTMLElement>(`.btn[data-command="${id}"]`)?.getBoundingClientRect();
      const toItem = ([ru, en]: [string, string]): MenuItem => ({ label: ru, keys: en === ru ? undefined : en, run: () => insertFunctionText(ru) });
      const items: MenuItem[] = cats.length === 1
        ? (FUNCTION_CATEGORIES[cats[0]] ?? []).map(toItem)
        : cats.map((c) => ({ label: c, run: () => {}, sub: (FUNCTION_CATEGORIES[c] ?? []).map(toItem) }));
      showMenu(items, b ? b.left : lastPointer.x, b ? b.bottom + 2 : lastPointer.y + 8);
    },
  })),
  { id: 'chart.bar', label: 'Диаграмма: столбцы', keywords: ['chart', 'column chart', 'диаграмма', 'гистограмма', 'график'], run: () => insertChart(0) },
  { id: 'chart.line', label: 'Диаграмма: линии', keywords: ['line chart', 'график', 'диаграмма'], run: () => insertChart(1) },
  { id: 'chart.pie', label: 'Диаграмма: круговая', keywords: ['pie chart', 'круговая', 'диаграмма'], run: () => insertChart(2) },
  { id: 'chart.edit', label: 'Изменить диаграмму', keywords: ['edit chart'], run: () => editSelectedChart() },
  { id: 'insert.image', label: 'Вставить рисунок', keywords: ['picture', 'image', 'рисунок', 'картинка', 'изображение', 'фото'],
    run: () => imageInput.click() },
  { id: 'protect.sheet', label: 'Защитить лист / снять защиту', keywords: ['protect sheet', 'защита листа', 'защитить', 'блокировка'],
    checked: () => model.isSheetProtected(grid.sheet), run: () => toggleSheetProtection() },
  { id: 'protect.book', label: 'Защитить структуру книги / снять защиту', keywords: ['protect workbook', 'защита книги', 'структура книги'],
    checked: () => model.isStructureLocked(), run: () => {
      try { model.setStructureLocked(!model.isStructureLocked()); showToast(model.isStructureLocked() ? 'Структура книги защищена.' : 'Защита структуры снята.'); } catch (e) { showToast(String(e)); }
      renderSheetTabs(); refreshRibbon();
    } },
  { id: 'protect.cells', label: 'Защита ячеек…', keywords: ['locked', 'hidden', 'скрыть формулы', 'защищаемая ячейка'],
    run: () => cellProtectionDialog() },
  { id: 'comment.edit', label: 'Примечание (создать или изменить)', keywords: ['comment', 'note', 'примечание', 'комментарий'], keys: ['Shift+F2'],
    run: () => editComment(grid.sel.row, grid.sel.column) },
  { id: 'comment.delete', label: 'Удалить примечание', keywords: ['delete comment', 'удалить комментарий'],
    run: () => deleteComment(grid.sel.row, grid.sel.column) },
  { id: 'page.breakInsert', label: 'Вставить разрыв страницы', keywords: ['page break', 'разрыв страницы', 'печать'],
    run: () => pageBreak(true) },
  { id: 'page.breakRemove', label: 'Убрать разрыв страницы', keywords: ['remove page break', 'убрать разрыв'],
    run: () => pageBreak(false) },
  { id: 'page.breakReset', label: 'Сбросить все разрывы страниц', keywords: ['reset page breaks'],
    run: () => { model.clearPageBreaks(grid.sheet); grid.render(); } },
  { id: 'insert.symbol', label: 'Вставить символ', keywords: ['symbol', 'символ', 'спецсимвол', 'знак', 'греческие буквы', 'валюта', 'стрелка'],
    run: () => openSymbols(insertSymbol) },
  { id: 'formulas.names', label: 'Диспетчер имён', keywords: ['name manager', 'имена', 'именованный диапазон', 'присвоить имя'], keys: ['Ctrl+F3'],
    run: () => openNameManager() },
  { id: 'format.strike', label: 'Зачёркнутый', keywords: ['strikethrough', 'зачеркнуть', 'зачёркивание'], keys: ['Ctrl+Digit5'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.font.strike,
    run: () => style('font.strike', (cur) => String(!cur), (st) => st.font.strike) },
  { id: 'format.superscript', label: 'Надстрочный', keywords: ['superscript', 'верхний индекс', 'степень'], checked: () => vertAlignOf() === 'superscript',
    run: () => style('font.vert_align', () => (vertAlignOf() === 'superscript' ? 'baseline' : 'superscript')) },
  { id: 'format.subscript', label: 'Подстрочный', keywords: ['subscript', 'нижний индекс'], checked: () => vertAlignOf() === 'subscript',
    run: () => style('font.vert_align', () => (vertAlignOf() === 'subscript' ? 'baseline' : 'subscript')) },
  { id: 'align.justify', label: 'По ширине', keywords: ['justify', 'выровнять по ширине'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.horizontal === 'justify',
    run: () => style('alignment.horizontal', () => 'justify') },
  { id: 'format.fill', label: 'Цвет заливки', keywords: ['fill', 'заливка', 'цвет', 'фон'], run: (c) => style('fill.fg_color', () => c ?? '') },
  { id: 'format.fontColor', label: 'Цвет шрифта', keywords: ['font color', 'цвет текста'], run: (c) => style('font.color', () => c ?? '') },
  { id: 'format.fontName', label: 'Шрифт', keywords: ['font', 'гарнитура', 'начертание'], run: (f) => { if (f) style('font.name', () => f); } },
  { id: 'format.fontSize', label: 'Размер шрифта', keywords: ['font size', 'кегль'], run: (n) => { if (n) style('font.size', () => n); } },
  { id: 'align.left', label: 'По левому краю', keywords: ['align left'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.horizontal === 'left', run: () => style('alignment.horizontal', () => 'left') },
  { id: 'align.center', label: 'По центру', keywords: ['align center', 'центр'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.horizontal === 'center', run: () => style('alignment.horizontal', () => 'center') },
  { id: 'align.right', label: 'По правому краю', keywords: ['align right'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.horizontal === 'right', run: () => style('alignment.horizontal', () => 'right') },
  { id: 'align.top', label: 'По верхнему краю', keywords: ['align top', 'вверху'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.vertical === 'top', run: () => style('alignment.vertical', () => 'top') },
  { id: 'align.middle', label: 'По середине (по вертикали)', keywords: ['middle', 'по вертикали'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.vertical === 'center', run: () => style('alignment.vertical', () => 'center') },
  { id: 'align.bottom', label: 'По нижнему краю', keywords: ['align bottom', 'внизу'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.vertical === 'bottom', run: () => style('alignment.vertical', () => 'bottom') },
  { id: 'view.formulaBar', label: 'Развернуть или свернуть строку формул', keywords: ['formula bar', 'строка формул', 'высота'], keys: ['Ctrl+Shift+KeyU'],
    run: () => toggleFormulaHeight() },
  { id: 'data.group', label: 'Сгруппировать', keywords: ['group', 'группа', 'свернуть', 'структура'], keys: ['Shift+Alt+ArrowRight'],
    run: () => changeOutline(1) },
  { id: 'data.summary', label: 'Положение «+» у групп', keywords: ['итоги под данными', 'итоги над данными', 'summary below', 'структура'],
    run: () => { const b = document.querySelector<HTMLElement>('.btn[data-command="data.summary"]')?.getBoundingClientRect(); summaryMenu(b ? b.left : lastPointer.x, b ? b.bottom + 2 : lastPointer.y + 8); } },
  { id: 'view.markRow', label: 'Метка строки (поставить или снять)', keywords: ['bookmark', 'закладка', 'метка', 'отметить строку'], keys: ['Ctrl+Alt+KeyB'],
    run: () => grid.toggleMark('row', grid.sel.row) },
  { id: 'view.markColumn', label: 'Метка столбца (поставить или снять)', keywords: ['bookmark', 'закладка', 'метка', 'отметить столбец'], keys: ['Ctrl+Alt+Shift+KeyB'],
    run: () => grid.toggleMark('column', grid.sel.column) },
  { id: 'view.markNext', label: 'К следующей метке', keywords: ['next bookmark', 'закладка', 'следующая метка'], keys: ['Ctrl+Alt+ArrowDown'],
    run: () => jumpMark(1) },
  { id: 'view.markPrev', label: 'К предыдущей метке', keywords: ['previous bookmark', 'закладка', 'предыдущая метка'], keys: ['Ctrl+Alt+ArrowUp'],
    run: () => jumpMark(-1) },
  { id: 'data.ungroup', label: 'Разгруппировать', keywords: ['ungroup', 'группа', 'структура'], keys: ['Shift+Alt+ArrowLeft'],
    run: () => changeOutline(-1) },
  { id: 'format.rotate', label: 'Направление текста', keywords: ['text direction', 'поворот текста', 'вертикально', 'под углом', 'ориентация'],
    run: () => {
      const b = document.querySelector<HTMLElement>('.btn[data-command="format.rotate"]')?.getBoundingClientRect();
      rotationMenu(b ? b.left : lastPointer.x, b ? b.bottom + 2 : lastPointer.y + 8);
    } },
  { id: 'view.hidden', label: 'Скрытое: что делать при действиях', keywords: ['скрытые строки', 'видимые', 'hidden'],
    run: (v) => { if (v) { settings.hiddenPolicy = v as typeof settings.hiddenPolicy; saveSettings(); } } },
  { id: 'format.wrap', label: 'Перенос текста', keywords: ['wrap text', 'перенос', 'переносить по словам'],
    checked: () => !!model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.alignment?.wrap_text,
    run: () => { style('alignment.wrap_text', (cur) => String(!cur), (st) => !!st.alignment?.wrap_text); fitRows(grid.sel.r1, grid.sel.r2); } },
  { id: 'format.merge', label: 'Объединить и поместить в центр', keywords: ['merge', 'объединить ячейки', 'объединение'],
    checked: () => inMerge(),
    run: () => merge('center') },
  { id: 'format.unmerge', label: 'Отменить объединение ячеек', keywords: ['unmerge', 'разъединить'], run: () => merge('undo') },
  { id: 'format.borders', label: 'Границы (последний выбор)', keywords: ['borders', 'рамка', 'границы', 'обрамление'],
    run: () => applyBorders() },
  { id: 'format.borders:menu', label: 'Границы: выбрать вид и линию', keywords: ['тип линии', 'толщина границы'],
    run: (at) => {
      const [x, y] = (at ?? '').split(',').map(Number);
      bordersMenu(Number.isFinite(x) ? x : lastPointer.x, Number.isFinite(y) ? y : lastPointer.y + 8);
    } },
  { id: 'format.conditional', label: 'Условное форматирование', keywords: ['conditional formatting', 'выделить ячейки', 'правила', 'гистограмма', 'цветовая шкала', 'значки'],
    run: () => openConditionalMenu() },
  { id: 'format.conditional:new', label: 'Условное форматирование: создать правило', keywords: ['new rule', 'условие'],
    run: () => { void editRule(null, { range: rectRef(selectionRect()) }).then(() => focusGrid()); } },
  { id: 'format.conditional:manage', label: 'Условное форматирование: управление правилами', keywords: ['manage rules', 'удалить правила'],
    run: () => { void manageRules(); } },
  { id: 'view.theme', label: 'Тема оформления', keywords: ['theme', 'тема', 'оформление', 'пустошь', 'постапокалипсис', 'дизайн'],
    run: () => themeMenu() },
  { id: 'format.cells', label: 'Формат ячеек…', keywords: ['format cells', 'числовой формат', 'дата', 'процент', 'разрядность'], keys: ['Ctrl+Digit1'],
    run: () => openFormatCells() },
  { id: 'format.decimalsMore', label: 'Увеличить разрядность', keywords: ['increase decimal', 'знаки после запятой', 'десятичные'],
    run: () => numFormat((c, shown) => changeDecimals(c, 1, shown), 'Разрядность') },
  { id: 'format.decimalsLess', label: 'Уменьшить разрядность', keywords: ['decrease decimal', 'знаки после запятой', 'округлить вид'],
    run: () => numFormat((c, shown) => changeDecimals(c, -1, shown), 'Разрядность') },
  { id: 'format.grouping', label: 'Разделитель групп разрядов', keywords: ['thousands separator', 'разряды', 'пробел между тысячами'], keys: ['Ctrl+Shift+Digit1'],
    run: () => numFormat(toggleGrouping, 'Разделитель разрядов') },
  { id: 'format.percent', label: 'Процентный формат', keywords: ['percent', 'процент'], keys: ['Ctrl+Shift+Digit5'],
    checked: () => /%/.test(model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.num_fmt),
    run: () => style('num_fmt', () => '0%') },
  { id: 'format.currency', label: 'Денежный формат (₽)', keywords: ['currency', 'рубли', 'деньги'], keys: ['Ctrl+Shift+Digit4'],
    checked: () => /₽/.test(model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.num_fmt),
    run: () => style('num_fmt', () => '#,##0.00 "₽";[Red]-#,##0.00 "₽"') },
  { id: 'format.number', label: 'Числовой формат', keywords: ['number format', 'формат ячеек'], run: (fmt) => {
    if (fmt === ':more') { syncNumberFormat(); openFormatCells(); return; }
    if (fmt) style('num_fmt', () => fmt);
  } },
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
  { id: 'edit.selectAll', label: 'Выделить всё', keywords: ['select all', 'выделить все', 'выделить всё'], keys: ['Ctrl+KeyA'],
    run: () => {
      // Как Excel: сначала текущая область данных, повторное нажатие — весь лист.
      const s = grid.sel;
      const single = s.r1 === s.r2 && s.c1 === s.c2;
      const region = currentRegion(s.row, s.column, single ? undefined : s);
      const same = region.r1 === s.r1 && region.c1 === s.c1 && region.r2 === s.r2 && region.c2 === s.c2;
      if (same || (single && region.r1 === region.r2 && region.c1 === region.c2)) grid.selectRange(1, 1, LAST_ROW, LAST_COLUMN);
      else grid.selectRange(region.r1, region.c1, region.r2, region.c2);
    } },
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
    checked: () => !!filterAt(grid.sheet, grid.sel.row, grid.sel.column), run: toggleFilter },
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
onRibbonDrawn(syncBordersButton);
setHiddenThemes(settings.hiddenThemes);
const setTheme = () => applyTheme(settings.theme, grid, settings.themeMotion, settings.themeMotionForced);
setCustomThemes(settings.customThemes);
setTheme();

/** «Загрузить тему из файла…»: файл .json (образец — «Сохранить образец темы»). */
function loadThemeFile(): void {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json,.enstheme,application/json';
  input.addEventListener('change', async () => {
    const file = input.files?.[0];
    if (!file) return;
    const t = parseThemeFile(await file.text());
    if (typeof t === 'string') { showToast(t); return; }
    settings.customThemes = [...settings.customThemes.filter((x) => x.id !== t.id), t];
    settings.theme = t.id;
    saveSettings();
    setCustomThemes(settings.customThemes);
    setTheme();
    showToast(`Тема «${t.name}» добавлена и включена.`);
  });
  input.click();
}

/** «Вид → Тема»: выбор темы оформления (книга и печать не меняются). */
const THEMES_BUILTIN = new Set(THEMES.map((t) => t.id));

function themeMenu(): void {
  const b = document.querySelector<HTMLElement>('.btn[data-command="view.theme"]')?.getBoundingClientRect();
  const reduce = systemReducesMotion();
  const moving = settings.themeMotion && (!reduce || settings.themeMotionForced);
  const mine = settings.customThemes;
  showMenu([
    ...allThemes().map((t) => ({
      label: `${t.name} — ${t.note}`, checked: t.id === settings.theme, icon: '',
      run: () => { settings.theme = t.id; saveSettings(); setTheme(); },
    })),
    { label: 'Загрузить тему из файла…', separator: true, run: loadThemeFile },
    {
      label: 'Сохранить образец темы (в буфер обмена)',
      run: () => {
        void copyText(themeTemplate(settings.theme))
          .then(() => showToast('Образец темы скопирован. Вставьте в Блокнот, поменяйте цвета и сохраните как файл .json.'))
          .catch((e) => showToast(String(e)));
      },
    },
    { label: 'Скрыть выбранную тему из списка', separator: true, disabled: settings.theme === 'classic' || !THEMES_BUILTIN.has(settings.theme), run: () => {
      settings.hiddenThemes = [...settings.hiddenThemes, settings.theme]; setHiddenThemes(settings.hiddenThemes);
      settings.theme = 'classic'; saveSettings(); setTheme();
    } },
    ...(settings.hiddenThemes.length ? [{ label: `Показать скрытые темы (${settings.hiddenThemes.length})`, run: () => {
      settings.hiddenThemes = []; setHiddenThemes([]); saveSettings();
    } }] : []),
    ...mine.map((t, i) => ({
      label: `Удалить свою тему «${t.name}»`, separator: i === 0,
      run: () => {
        settings.customThemes = settings.customThemes.filter((x) => x.id !== t.id);
        if (settings.theme === t.id) settings.theme = 'classic';
        saveSettings();
        setCustomThemes(settings.customThemes);
        setTheme();
      },
    })),
    {
      // Анимация идёт, если включена и система не просит «уменьшить движение» (или включили вопреки ей).
      label: moving ? 'Выключить анимацию темы' : reduce ? 'Включить анимацию темы (в системе она уменьшена)' : 'Включить анимацию темы',
      separator: true,
      run: () => {
        settings.themeMotion = !moving;
        settings.themeMotionForced = !moving && reduce;
        saveSettings();
        setTheme();
      },
    },
  ], b ? b.left : lastPointer.x, b ? b.bottom + 2 : lastPointer.y + 8);
}
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
  setSelect('view.hidden', settings.hiddenPolicy);
  const select = document.querySelector<HTMLSelectElement>('select[data-command="format.number"]');
  if (!select) return;
  const fmt = st.num_fmt;
  const known = [...select.options].some((o) => o.value === fmt);
  if (!known) {
    select.querySelector('option[data-custom]')?.remove();
    const custom = new Option(`${formatName(fmt)}: ${fmt}`, fmt);
    custom.dataset.custom = '1';
    select.add(custom, select.querySelector<HTMLOptionElement>('option[value=":more"]'));
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

async function autosum(): Promise<void> {
  if (editing) commitEdit(0, 0);
  const s = grid.sel;
  // Скрытые строки в суммируемом: СУММ — со всеми, ПРОМЕЖУТОЧНЫЕ.ИТОГИ(109) — только видимые.
  const sumOf = async (r1: number, c1: number, r2: number, c2: number): Promise<string | null> => {
    const info = hiddenIn({ r1, c1, r2, c2 });
    if (!info.rows && !info.cols) return 'СУММ(';
    const c = await askHidden('Автосумма', info);
    if (!c) return null;
    return c === 'visible' ? 'ПРОМЕЖУТОЧНЫЕ.ИТОГИ(109;' : 'СУММ(';
  };
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
    const fn = await sumOf(s.r1, s.c1, bottom, s.c2);
    if (!fn) return;
    const mark = model.undoMark();
    for (let c = s.c1; c <= s.c2; c++) model.setUserInput(sheet, target, c, `=${fn}${ref(s.r1, c, bottom, c)})`);
    model.mergeUndoSince(mark);
    grid.select(target, s.c1);
    grid.render();
    return;
  }
  // Одна ячейка: сумма чисел над ней, а если над ней пусто — слева.
  let top = s.row - 1;
  while (top >= 1 && isNum(top, s.column)) top--;
  if (top < s.row - 1) {
    const fn = await sumOf(top + 1, s.column, s.row - 1, s.column);
    if (fn) startEdit(`=${fn}${ref(top + 1, s.column, s.row - 1, s.column)})`);
    return;
  }
  let left = s.column - 1;
  while (left >= 1 && isNum(s.row, left)) left--;
  if (left < s.column - 1) {
    const fn = await sumOf(s.row, left + 1, s.row, s.column - 1);
    if (fn) startEdit(`=${fn}${ref(s.row, left + 1, s.row, s.column - 1)})`);
    return;
  }
  startEdit('=СУММ()');
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
    newBookFont.name = settings.defaultFont;
    newBookFont.size = settings.defaultFontSize;
    grid.render(); updateFormulaBar(); updateStatus(); refreshRibbon(); syncNumberFormat();
  },
  recoveryFolder: () => desk.recoveryFolder(),
  chooseRecoveryFolder: () => desk.chooseRecoveryFolder(),
  printPreview: (el, opts) => renderPreview(el, model, grid.sheet, opts, grid.sel, doc.title),
  printArea: () => {
    const a = printAreaOf(model, grid.sheet);
    return a ? `${columnName(a.c1)}${a.r1}:${columnName(a.c2)}${a.r2}` : '';
  },
  setPrintArea: (clear) => {
    const s = grid.sel;
    try { model.setPrintArea(grid.sheet, clear ? [] : [s.r1, s.c1, Math.min(s.r2, 1_048_576), Math.min(s.c2, 16_384)]); } catch (e) { showToast(String(e)); }
  },
  printTitles: () => {
    const t = printTitles(model, grid.sheet);
    return { rows: t[0] ? (t[0] === t[1] ? String(t[0]) : `${t[0]}:${t[1]}`) : '', cols: t[2] ? (t[2] === t[3] ? columnName(t[2]) : `${columnName(t[2])}:${columnName(t[3])}`) : '' };
  },
  setPrintTitles: (rows, cols) => {
    const r = rows.trim().replace(/\$/g, '') ? /^(\d+)(?::(\d+))?$/.exec(rows.trim().replace(/\$/g, '')) : [];
    const c = cols.trim().replace(/\$/g, '') ? /^([A-Za-z]{1,3})(?::([A-Za-z]{1,3}))?$/.exec(cols.trim().replace(/\$/g, '')) : [];
    if (!r) return 'Сквозные строки: укажите номера, например 1:3.';
    if (!c) return 'Сквозные столбцы: укажите буквы, например A:B.';
    const r1 = r[1] ? Number(r[1]) : 0, r2 = r[1] ? Number(r[2] ?? r[1]) : 0;
    const c1 = c[1] ? columnNumber(c[1].toUpperCase()) : 0, c2 = c[1] ? columnNumber((c[2] ?? c[1]).toUpperCase()) : 0;
    if ((r1 && (r2 < r1 || r2 > 1_048_576)) || (c1 && (c2 < c1 || c2 > 16_384))) return 'Диапазон указан неверно: начало должно быть не больше конца.';
    try { model.setPrintTitles(grid.sheet, [r1, r2, c1, c2]); } catch (e) { return String(e); }
    return '';
  },
  showChangeLog: () => showLog(model, { goTo: goToArea }),
  printOptions: () => { try { return fromSaved(model.getPageSetup(grid.sheet)); } catch { return { ...DEFAULT_PRINT }; } },
  savePrintOptions: (opts) => { try { model.setPageSetup(grid.sheet, toSaved(opts)); } catch { /* старый движок */ } },
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
  goTo: goToArea,
});
