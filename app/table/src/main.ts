import './styles.css';
import { allCommands, handleKey, keyLabel, matchScore, register, run } from './commands';
import { columnName, columnNumber, loadEngine, newWorkbook, openXlsx, type Area, type Model } from './engine';
import { Grid } from './grid';
import { icon, renderRibbon } from './ribbon';
import { FUNCTIONS } from './functions';
import { localeId, saveSettings, settings } from './settings';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

await loadEngine();
let model: Model = newWorkbook(localeId());
let fileName = 'Книга1';

const grid = new Grid($('grid'), () => model);
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
  const s = grid.sel;
  const r = grid.cellRect(s.row, s.column);
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
  measure.font = getComputedStyle(editor).font;
  const need = measure.measureText(editor.value).width + 14;
  const max = grid.scroller.clientWidth - r.x - 4;
  editor.style.width = `${Math.min(Math.max(r.w, need), Math.max(r.w, max))}px`;
}

/** Режим ввода, как в Excel: «ввод» (начали печатать — стрелки завершают ввод)
 *  или «правка» (F2, двойной щелчок, строка формул — стрелки двигают курсор). */
let editMode: 'enter' | 'edit' = 'enter';

function focusGrid(): void {
  placeEditor();
  editor.focus({ preventScroll: true });
}

function startEdit(initial?: string): void {
  editMode = initial === undefined ? 'edit' : 'enter';
  placeEditor();
  editor.value = initial ?? model.getCellContent(grid.sheet, grid.sel.row, grid.sel.column);
  formula.value = editor.value;
  editing = true;
  editor.classList.add('editing');
  editor.focus();
  editor.setSelectionRange(editor.value.length, editor.value.length);
}

function commitEdit(dRow: number, dCol: number): void {
  if (!editing) return;
  endPoint();
  editing = false;
  editor.classList.remove('editing');
  const s = grid.sel;
  const before = model.getCellStyle(grid.sheet, s.row, s.column).style.num_fmt;
  try {
    model.setUserInput(grid.sheet, s.row, s.column, editor.value);
  } catch (err) {
    alert(`Не удалось записать значение: ${err}`);
  }
  const after = model.getCellStyle(grid.sheet, s.row, s.column).style.num_fmt;
  if (after !== before) offerFormatUndo(s.row, s.column, before, after);
  editor.value = '';
  grid.render();
  grid.move(dRow, dCol, false);
  focusGrid();
}

function cancelEdit(): void {
  endPoint();
  editing = false;
  editor.classList.remove('editing');
  editor.value = '';
  updateFormulaBar();
  focusGrid();
}

editor.addEventListener('input', () => {
  if (!editing) {
    // Первый введённый символ начинает редактирование, как в Excel.
    editing = true;
    editMode = 'enter';
    editor.classList.add('editing');
  }
  endPoint();
  formula.value = editor.value;
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
  return r1 === r2 && c1 === c2 ? a : `${a}:${columnName(c2)}${r2}`;
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
    const from = point ? { row: point.row, col: point.col } : { row: grid.sel.row, col: grid.sel.column };
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
editor.addEventListener('blur', () => { if (editing && document.activeElement !== formula) commitEdit(0, 0); });

formula.addEventListener('focus', () => { if (!editing) { startEdit(); formula.focus(); } });
formula.addEventListener('input', () => { endPoint(); editor.value = formula.value; });
formula.addEventListener('keydown', (e) => editKey(e));

// §3: Ctrl+щелчок по ссылке в строке формул — переход к ячейке.
const REF = /(?:'([^']+)'|([A-Za-zА-Яа-я_][\wА-Яа-я.]*))?!?\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?/g;

interface Ref { sheet?: string; row: number; column: number; text: string; start: number; end: number }

function refsIn(text: string): Ref[] {
  const out: Ref[] = [];
  for (const m of text.matchAll(REF)) {
    const sheet = m[1] ?? (m[0].includes('!') ? m[2] : undefined);
    out.push({ sheet, column: columnNumber(m[3]), row: Number(m[4]), text: m[0], start: m.index!, end: m.index! + m[0].length });
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
  grid.select(ref.row, ref.column);
  focusGrid();
}

formula.addEventListener('click', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const at = formula.selectionStart ?? 0;
  const ref = refsIn(formula.value).find((r) => at >= r.start && at <= r.end);
  if (ref) { e.preventDefault(); cancelEdit(); goTo(ref); }
});

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

const LAST_ROW = 1_048_576;
const LAST_COLUMN = 16_384;

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
  }
  // Печатные символы не перехватываем: они попадают в поле ввода
  // и начинают редактирование через событие input.
}

// --- Листы ---

function setSheet(index: number): void {
  grid.sheet = index;
  model.setSelectedSheet(index);
  grid.invalidateLayout();
  grid.select(1, 1);
  renderSheetTabs();
}

function renderSheetTabs(): void {
  const host = $('sheet-tabs');
  host.replaceChildren();
  model.getWorksheetsProperties().forEach((p, i) => {
    if (p.state !== 'visible') return;
    const b = document.createElement('button');
    b.className = `sheet-tab${i === grid.sheet ? ' active' : ''}`;
    b.textContent = p.name;
    b.addEventListener('click', () => setSheet(i));
    host.append(b);
  });
  const add = document.createElement('button');
  add.className = 'sheet-add';
  add.title = 'Новый лист (Shift+F11)';
  add.textContent = '+';
  add.addEventListener('click', () => run('sheet.new'));
  host.append(add);
}

// --- Файлы ---
// В браузере файл нельзя перезаписать на месте: «Сохранить» скачивает копию.
// Надёжное сохранение (временный файл → проверка → замена) — в оболочке Tauri (ADR-005).

const fileInput = $<HTMLInputElement>('file-input');
fileInput.addEventListener('change', async () => {
  const file = fileInput.files?.[0];
  if (!file) return;
  try {
    const next = openXlsx(new Uint8Array(await file.arrayBuffer()), file.name, localeId());
    model.free();
    model = next;
    fileName = file.name.replace(/\.xlsx$/i, '');
    document.title = `${fileName} — ENS-Table`;
    $('doc-name').textContent = fileName;
    setSheet(0);
  } catch (err) {
    alert(`Не удалось открыть «${file.name}»:\n${err}`);
  }
  fileInput.value = '';
});

function save(): void {
  const bytes = model.toXlsx();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  a.download = `${fileName}.xlsx`;
  a.click();
  URL.revokeObjectURL(a.href);
}

// --- Буфер обмена: через события браузера (не нужны разрешения) ---

editor.addEventListener('copy', (e) => {
  if (editing) return;
  e.preventDefault();
  e.clipboardData?.setData('text/plain', toTsv());
});
editor.addEventListener('cut', (e) => {
  if (editing) return;
  e.preventDefault();
  e.clipboardData?.setData('text/plain', toTsv());
  run('edit.clear');
});
editor.addEventListener('paste', (e) => {
  if (editing) return; // в режиме ввода — обычная вставка текста
  e.preventDefault();
  pasteText(e.clipboardData?.getData('text/plain') ?? '');
});

function pasteText(text: string): void {
  if (!text) return;
  model.pasteCsvText(selectionArea(), text.replace(/\r\n?/g, '\n').replace(/\n$/, ''));
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
  { id: 'file.open', label: 'Открыть', keywords: ['open', 'файл', 'xlsx'], keys: ['Ctrl+KeyO'], run: () => fileInput.click() },
  { id: 'file.save', label: 'Сохранить копию (xlsx)', keywords: ['save', 'скачать'], keys: ['Ctrl+KeyS'], run: save },
  { id: 'edit.undo', label: 'Отменить', keywords: ['undo'], keys: ['Ctrl+KeyZ'], run: () => { model.undo(); grid.render(); updateFormulaBar(); } },
  { id: 'edit.redo', label: 'Повторить', keywords: ['redo'], keys: ['Ctrl+KeyY'], run: () => { model.redo(); grid.render(); updateFormulaBar(); } },
  { id: 'edit.clear', label: 'Очистить содержимое', keywords: ['delete', 'удалить'], keys: ['Delete'], run: () => {
    const s = grid.sel; model.rangeClearContents(grid.sheet, s.r1, s.c1, s.r2, s.c2); grid.render(); updateFormulaBar();
  } },
  { id: 'edit.autosum', label: 'Автосумма', keywords: ['sum', 'сумма', 'сумм'], keys: ['Alt+Equal'], run: () => {
    const s = grid.sel;
    let top = s.row - 1;
    while (top >= 1 && model.getCellType(grid.sheet, top, s.column) === 1) top--;
    startEdit(top < s.row - 1 ? `=SUM(${columnName(s.column)}${top + 1}:${columnName(s.column)}${s.row - 1})` : '=SUM()');
  } },
  { id: 'clipboard.copy', label: 'Копировать', keywords: ['copy'], keys: ['Ctrl+KeyC'], nativeKeys: true,
    run: () => { focusGrid(); document.execCommand('copy'); } },
  { id: 'clipboard.cut', label: 'Вырезать', keywords: ['cut'], keys: ['Ctrl+KeyX'], nativeKeys: true,
    run: () => { focusGrid(); document.execCommand('cut'); } },
  { id: 'clipboard.paste', label: 'Вставить', keywords: ['paste'], keys: ['Ctrl+KeyV'], nativeKeys: true, run: async () => {
    try {
      pasteText(await navigator.clipboard.readText());
    } catch {
      showToast('Браузер не дал прочитать буфер обмена с кнопки. Нажмите Ctrl+V — так вставка работает всегда.');
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
  { id: 'format.number', label: 'Числовой формат', keywords: ['number format', 'формат ячеек'], run: (fmt) => { if (fmt) style('num_fmt', () => fmt); } },
  { id: 'sheet.new', label: 'Новый лист', keywords: ['new sheet', 'лист'], keys: ['Shift+F11'], run: () => {
    model.newSheet(); setSheet(model.getWorksheetsProperties().length - 1);
  } },
  { id: 'view.focusCell', label: 'Фокус ячейки (перекрестие)', keywords: ['focus cell', 'подсветка строки', 'перекрестие', 'crosshair'],
    keys: ['Ctrl+Shift+KeyK'], checked: () => grid.focusCell.enabled,
    run: () => { grid.focusCell.enabled = !grid.focusCell.enabled; grid.render(); refreshRibbon(); } },
  { id: 'view.gridlines', label: 'Сетка', keywords: ['gridlines', 'линии сетки'], checked: () => model.getShowGridLines(grid.sheet),
    run: () => { model.setShowGridLines(grid.sheet, !model.getShowGridLines(grid.sheet)); grid.render(); refreshRibbon(); } },
  { id: 'nav.precedents', label: 'Перейти к влияющим ячейкам', keywords: ['trace precedents', 'ссылки формулы'], keys: ['Ctrl+BracketLeft'],
    run: () => {
      const refs = refsIn(model.getCellContent(grid.sheet, grid.sel.row, grid.sel.column).startsWith('=')
        ? model.getCellContent(grid.sheet, grid.sel.row, grid.sel.column) : '');
      if (refs.length === 1) goTo(refs[0]);
      else if (refs.length > 1) showPicker(refs.map((r) => ({ label: r.text, run: () => goTo(r) })));
    } },
  { id: 'nav.back', label: 'Назад к предыдущей ячейке', keywords: ['back', 'назад'], keys: ['Alt+ArrowLeft'], run: () => {
    const prev = backStack.pop();
    if (!prev) return;
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
    checked: () => model.getRecognizeDates(),
    run: () => { model.setRecognizeDates(!model.getRecognizeDates()); refreshRibbon(); } },
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

function searchResults(q: string) {
  const found: { score: number; label: string; hint?: string; run: () => void }[] = [];
  for (const c of allCommands()) {
    if (c.id === 'search.focus') continue;
    const score = Math.max(matchScore(q, c.label), ...(c.keywords ?? []).map((k) => matchScore(q, k) * 0.9));
    if (score > 0) found.push({ score: score + 5, label: c.label, hint: c.keys?.[0] && keyLabel(c.keys[0]), run: () => c.run() });
  }
  for (const f of FUNCTIONS) {
    const score = Math.max(matchScore(q, f.ru), matchScore(q, f.en), matchScore(q, f.description) * 0.7);
    if (score > 0) {
      found.push({ score, label: `ƒ ${f.ru} (${f.en}) — ${f.description}`, run: () => startEdit(`=${f.en}(`) });
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

const refreshRibbon = renderRibbon($('tabs'), $('ribbon'), () => run('file.open'));
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
$('doc-name').textContent = fileName;
renderSheetTabs();
grid.select(1, 1);
focusGrid();
