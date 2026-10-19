import './styles.css';
import { allCommands, handleKey, keyLabel, matchScore, register, run } from './commands';
import { columnName, columnNumber, loadEngine, newWorkbook, openXlsx, type Area, type Model } from './engine';
import { Grid } from './grid';
import { icon, renderRibbon } from './ribbon';
import { FUNCTIONS } from './functions';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

await loadEngine();
let model: Model = newWorkbook();
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
  editor.style.minWidth = `${r.w}px`;
  editor.style.height = `${r.h}px`;
}

function focusGrid(): void {
  placeEditor();
  editor.focus({ preventScroll: true });
}

function startEdit(initial?: string): void {
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
  editing = false;
  editor.classList.remove('editing');
  const s = grid.sel;
  try {
    model.setUserInput(grid.sheet, s.row, s.column, editor.value);
  } catch (err) {
    alert(`Не удалось записать значение: ${err}`);
  }
  editor.value = '';
  grid.render();
  grid.move(dRow, dCol, false);
  focusGrid();
}

function cancelEdit(): void {
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
    editor.classList.add('editing');
  }
  formula.value = editor.value;
});
editor.addEventListener('keydown', (e) => {
  if (!editing) { gridKey(e); return; }
  if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); commitEdit(e.shiftKey ? -1 : 1, 0); }
  else if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); }
  else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
});
editor.addEventListener('blur', () => { if (editing && document.activeElement !== formula) commitEdit(0, 0); });

formula.addEventListener('focus', () => { if (!editing) { startEdit(); formula.focus(); } });
formula.addEventListener('input', () => (editor.value = formula.value));
formula.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') { e.preventDefault(); commitEdit(1, 0); }
  else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
});

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
          const v = Number(text.replace(/[^\d.,\-E]/g, '').replace(',', '.'));
          if (Number.isFinite(v)) { numbers++; sum += v; }
        }
      }
    }
  }
  const fmt = (v: number) => v.toLocaleString('ru-RU', { maximumFractionDigits: 10 });
  status.textContent = count > 1
    ? `Среднее: ${numbers ? fmt(sum / numbers) : '—'}    Количество: ${count}    Сумма: ${fmt(sum)}`
    : '';
}

grid.onEditRequest = (initial) => startEdit(initial);

// --- Клавиатура сетки ---

function gridKey(e: KeyboardEvent): void {
  if (handleKey(e)) return;
  const shift = e.shiftKey;
  const ctrl = e.ctrlKey || e.metaKey;
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
    const next = openXlsx(new Uint8Array(await file.arrayBuffer()), file.name);
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

// --- Команды ---

function style(path: string, value: (current: boolean) => string, flag?: (s: ReturnType<Model['getCellStyle']>['style']) => boolean): void {
  const s = grid.sel;
  const current = flag ? flag(model.getCellStyle(grid.sheet, s.row, s.column).style) : false;
  model.updateRangeStyle(selectionArea(), path, value(current));
  grid.render();
  refreshRibbon();
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
  { id: 'clipboard.copy', label: 'Копировать', keywords: ['copy'], keys: ['Ctrl+KeyC'], run: () => void navigator.clipboard.writeText(toTsv()) },
  { id: 'clipboard.cut', label: 'Вырезать', keywords: ['cut'], keys: ['Ctrl+KeyX'], run: () => {
    void navigator.clipboard.writeText(toTsv()).then(() => run('edit.clear'));
  } },
  { id: 'clipboard.paste', label: 'Вставить', keywords: ['paste'], keys: ['Ctrl+KeyV'], run: async () => {
    const text = await navigator.clipboard.readText();
    model.pasteCsvText(selectionArea(), text);
    grid.render(); updateFormulaBar();
  } },
  { id: 'format.bold', label: 'Полужирный', keywords: ['bold', 'жирный'], keys: ['Ctrl+KeyB'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.font.b,
    run: () => style('font.b', (cur) => String(!cur), (st) => st.font.b) },
  { id: 'format.italic', label: 'Курсив', keywords: ['italic'], keys: ['Ctrl+KeyI'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.font.i,
    run: () => style('font.i', (cur) => String(!cur), (st) => st.font.i) },
  { id: 'format.underline', label: 'Подчёркнутый', keywords: ['underline'], keys: ['Ctrl+KeyU'], checked: () => model.getCellStyle(grid.sheet, grid.sel.row, grid.sel.column).style.font.u,
    run: () => style('font.u', (cur) => String(!cur), (st) => st.font.u) },
  { id: 'format.fill', label: 'Цвет заливки (жёлтый)', keywords: ['fill', 'заливка', 'цвет'], run: () => style('fill.fg_color', () => '#FFF2CC') },
  { id: 'format.fontColor', label: 'Цвет шрифта (красный)', keywords: ['font color'], run: () => style('font.color', () => '#C00000') },
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
grid.onSelect = () => { updateFormulaBar(); updateStatus(); refreshRibbon(); if (!editing) placeEditor(); };
grid.scroller.addEventListener('scroll', () => { if (!editing) placeEditor(); });
grid.scroller.addEventListener('mouseup', () => { if (!editing) focusGrid(); });
grid.scroller.addEventListener('focus', () => { if (!editing) focusGrid(); });
$('doc-name').textContent = fileName;
renderSheetTabs();
grid.select(1, 1);
focusGrid();
