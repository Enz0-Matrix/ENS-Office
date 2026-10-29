// ENS-Tekst: текстовый редактор ENS-Office. Лента, команды, файлы, восстановление.

import '../../table/src/styles.css';
import './text.css';
import './icons';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { redo, redoDepth, undo, undoDepth } from 'prosemirror-history';
import {
  addColumnAfter, addColumnBefore, addRowAfter, addRowBefore, deleteColumn, deleteRow, deleteTable, isInTable, mergeCells, splitCell,
} from 'prosemirror-tables';
import type { Node as PMNode } from 'prosemirror-model';
import { register, run, handleKey } from '../../table/src/commands';
import { FONT_SIZES, FONTS, lastColor, renderRibbon, type RibbonTab } from '../../table/src/ribbon';
import { fontStack, loadFonts } from '../../table/src/fonts';
import { ICONS } from '../../table/src/icons';
import { APP_VERSION } from './version';
import { schema } from './schema';
import {
  clearFormatting, createState, currentLevel, currentStyle, findKey, findStep, indent, insertImage, insertPageBreak, insertTable, linkAt,
  listKind, markActive, paraAttr, replaceAll, replaceMatch, setLink, setParaAttrs, setStyle, setTextStyle, toggle, toggleList, toggles,
} from './editor';
import { A4, exportDocx, importDocx, newMeta, type DocMeta, type PageSetup } from './docx';
import { clearRecovery, listRecovery, saveRecovery, type RecoveryRecord } from './recovery';
import * as desktop from './desktop';
import { readText } from '../../table/src/platform';
import { dialog, el, setRestoreFocus, toast } from './ui';
import { PAGE_GAP, PX, geometry, pageCount, pagination, replan, stats as pagStats } from './pagination';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

// ---------------------------------------------------------------- состояние документа

let meta: DocMeta = newMeta();
let savedMeta = JSON.stringify(meta.page) + JSON.stringify(meta.defaults);
let docName = 'Документ1';
let savedDoc: PMNode = schema.nodes.doc.createAndFill()!;
const docId = `doc-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
let zoom = 1;

function emptyDoc(): PMNode {
  return schema.nodes.doc.createAndFill()!;
}

const metaKey = () => JSON.stringify(meta.page) + JSON.stringify(meta.defaults);
const isDirty = () => !view.state.doc.eq(savedDoc) || metaKey() !== savedMeta;

// ---------------------------------------------------------------- страница

function applyPage(): void {
  const p = meta.page;
  const sheet = $('sheet');
  sheet.style.width = `${p.w}pt`;
  view.dom.style.padding = `${p.top}pt ${p.right}pt ${p.bottom}pt ${p.left}pt`;
  drawPapers(pageCount(view));
  const d = meta.defaults;
  const root = document.documentElement.style;
  root.setProperty('--doc-font', fontStack(d.font));
  root.setProperty('--doc-size', `${d.size}pt`);
  root.setProperty('--doc-line', d.line === 1 ? 'normal' : String((d.line * 1.15).toFixed(3)));
  root.setProperty('--doc-after', `${d.after}pt`);
  root.setProperty('--doc-before', `${d.before}pt`);
  let st = document.getElementById('print-page') as HTMLStyleElement | null;
  if (!st) { st = document.createElement('style'); st.id = 'print-page'; document.head.append(st); }
  st.textContent = `@page { size: ${p.w}pt ${p.h}pt; margin: ${p.top}pt ${p.right}pt ${p.bottom}pt ${p.left}pt; }`;
}

/** Листы-подложки: по одному на страницу; высота страницы и зазоры — как у расчёта разрывов. */
function drawPapers(pages: number): void {
  const g = geometry(meta.page);
  const papers = $('papers');
  if (papers.childElementCount !== pages || papers.dataset.h !== String(g.H)) {
    papers.replaceChildren(...Array.from({ length: pages }, (_, i) => {
      const d = document.createElement('div');
      d.className = 'paper';
      d.style.cssText = `top:${i * g.P}px;height:${g.H}px`;
      return d;
    }));
    papers.dataset.h = String(g.H);
  }
  const total = pages * g.H + (pages - 1) * PAGE_GAP;
  $('sheet').style.minHeight = `${total}px`;
  view.dom.style.minHeight = `${total}px`;
  updatePageStatus();
}

/** «Страница X из N» по положению курсора. */
function updatePageStatus(): void {
  const n = pageCount(view);
  let page = 1;
  try {
    const box = view.dom.getBoundingClientRect();
    const scale = box.width / (view.dom as HTMLElement).offsetWidth || 1;
    const c = view.coordsAtPos(view.state.selection.head);
    page = Math.min(n, Math.max(1, Math.floor((((c.top - box.top) / scale) + 1) / geometry(meta.page).P) + 1));
  } catch { /* вид ещё не готов */ }
  $('status-page').textContent = `Страница ${page} из ${n}`;
}

function setZoom(z: number): void {
  zoom = Math.min(3, Math.max(0.3, Math.round(z * 10) / 10));
  ($('sheet-wrap').style as CSSStyleDeclaration & { zoom: string }).zoom = String(zoom);
  $('zoom-val').textContent = `${Math.round(zoom * 100)}%`;
}

// ---------------------------------------------------------------- редактор

const paginationPlugin = pagination(() => meta.page, (n) => drawPapers(n));

const view = new EditorView($('sheet'), {
  state: createState(emptyDoc(), [paginationPlugin]),
  attributes: { spellcheck: 'true', lang: 'ru', 'aria-label': 'Текст документа' },
  dispatchTransaction(tr) {
    const before = view.state;
    view.updateState(before.apply(tr));
    afterChange(tr.docChanged);
  },
  handlePaste(v, event) {
    const files = [...(event.clipboardData?.files ?? [])].filter((f) => f.type.startsWith('image/'));
    if (!files.length) return false;
    void insertImageFile(files[0]);
    return true;
  },
  handleDrop(v, event) {
    const files = [...(event.dataTransfer?.files ?? [])].filter((f) => f.type.startsWith('image/'));
    if (!files.length) return false;
    void insertImageFile(files[0]);
    return true;
  },
  handleClick(v, pos, event) {
    if (!event.ctrlKey && !event.metaKey) return false;
    const $pos = v.state.doc.resolve(pos);
    const link = schema.marks.link.isInSet($pos.marks());
    if (!link) return false;
    if (desktop.isDesktop) {
      desktop.openTarget(link.attrs.href).catch((e) => toast(String((e as Error)?.message ?? e)));
      return true;
    }
    const a = document.createElement('a');
    a.href = link.attrs.href;
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.click();
    return true;
  },
});

let statsTimer = 0;
let recoveryTimer = 0;

function afterChange(docChanged: boolean): void {
  updateUi();
  if (docChanged) {
    clearTimeout(statsTimer);
    statsTimer = window.setTimeout(updateStats, 250);
    scheduleRecovery();
  } else {
    updateSelStats();
  }
}

function updateStats(): void {
  const text = view.state.doc.textBetween(0, view.state.doc.content.size, '\n', ' ');
  const words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-]*/gu) ?? []).length;
  $('status-words').textContent = `Слов: ${words}  ·  Знаков: ${text.replace(/\n/g, '').length}`;
  updateSelStats();
}

function updateSelStats(): void {
  updatePageStatus();
  const { from, to, empty } = view.state.selection;
  if (empty) { $('status-sel').textContent = ''; return; }
  const text = view.state.doc.textBetween(from, to, '\n', ' ');
  const words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-]*/gu) ?? []).length;
  $('status-sel').textContent = `Выделено слов: ${words}  ·  знаков: ${text.replace(/\n/g, '').length}`;
}

// ---------------------------------------------------------------- команды

const dispatchCmd = (cmd: (s: EditorState, d?: typeof view.dispatch) => boolean): void => {
  cmd(view.state, view.dispatch);
  view.focus();
};

const ALIGN_LABEL = { left: 'По левому краю', center: 'По центру', right: 'По правому краю', justify: 'По ширине' } as const;
const FORMAT_LABEL: Record<string, string> = {
  bold: 'Полужирный', italic: 'Курсив', underline: 'Подчёркнутый', strike: 'Зачёркнутый', superscript: 'Надстрочный', subscript: 'Подстрочный',
};
const FORMAT_KEYS: Record<string, string[]> = { bold: ['Ctrl+KeyB'], italic: ['Ctrl+KeyI'], underline: ['Ctrl+KeyU'] };

for (const name of Object.keys(toggles)) {
  register({
    id: `format.${name}`, label: FORMAT_LABEL[name], keys: FORMAT_KEYS[name],
    run: () => dispatchCmd(toggle(name)),
    checked: () => markActive(view.state, toggles[name]),
  });
}
register({ id: 'format.superscript', label: 'Надстрочный', keys: ['Ctrl+Shift+Equal'], run: () => dispatchCmd(toggle('superscript')), checked: () => markActive(view.state, toggles.superscript) });
register({ id: 'format.subscript', label: 'Подстрочный', keys: ['Ctrl+Equal'], run: () => dispatchCmd(toggle('subscript')), checked: () => markActive(view.state, toggles.subscript) });

register(
  { id: 'edit.undo', label: 'Отменить', keys: ['Ctrl+KeyZ'], run: () => { undo(view.state, view.dispatch); view.focus(); } },
  { id: 'edit.redo', label: 'Повторить', keys: ['Ctrl+KeyY', 'Ctrl+Shift+KeyZ'], run: () => { redo(view.state, view.dispatch); view.focus(); } },
  {
    id: 'clipboard.paste', label: 'Вставить (Ctrl+V)',
    run: () => {
      view.focus();
      (desktop.isDesktop ? readText() : navigator.clipboard.readText()).then((t) => { if (t) view.dispatch(view.state.tr.insertText(t).scrollIntoView()); })
        .catch(() => toast('Браузер не дал доступ к буферу обмена — нажмите Ctrl+V'));
    },
  },
  { id: 'clipboard.cut', label: 'Вырезать (Ctrl+X)', run: () => { view.focus(); document.execCommand('cut'); } },
  { id: 'clipboard.copy', label: 'Копировать (Ctrl+C)', run: () => { view.focus(); document.execCommand('copy'); } },
  { id: 'format.fontName', label: 'Шрифт', run: (v) => v && dispatchCmd(setTextStyle({ font: v })) },
  { id: 'format.fontSize', label: 'Размер шрифта', run: (v) => v && dispatchCmd(setTextStyle({ size: parseFloat(v) })) },
  { id: 'format.fontColor', label: 'Цвет текста', run: (v) => dispatchCmd(setTextStyle({ color: v || null })) },
  { id: 'format.highlight', label: 'Цвет выделения', run: (v) => dispatchCmd(setTextStyle({ bg: v || null })) },
  { id: 'format.clear', label: 'Очистить форматирование', keys: ['Ctrl+Space'], run: () => dispatchCmd(clearFormatting) },
  { id: 'format.sizeUp', label: 'Увеличить размер', keys: ['Ctrl+Shift+Period'], run: () => stepSize(1) },
  { id: 'format.sizeDown', label: 'Уменьшить размер', keys: ['Ctrl+Shift+Comma'], run: () => stepSize(-1) },
);

function stepSize(dir: 1 | -1): void {
  const sizes = FONT_SIZES.map((o) => +o.value);
  const cur = currentStyle(view.state).size ?? meta.defaults.size;
  const next = dir > 0 ? sizes.find((s) => s > cur) : [...sizes].reverse().find((s) => s < cur);
  if (next) dispatchCmd(setTextStyle({ size: next }));
}

for (const a of ['left', 'center', 'right', 'justify'] as const) {
  register({
    id: `align.${a}`, label: ALIGN_LABEL[a], keys: [{ left: 'Ctrl+KeyL', center: 'Ctrl+KeyE', right: 'Ctrl+KeyR', justify: 'Ctrl+KeyJ' }[a]],
    run: () => dispatchCmd(setParaAttrs({ align: a === 'left' ? null : a })),
    checked: () => (paraAttr<string>(view.state, 'align') ?? 'left') === a,
  });
}
register(
  { id: 'list.bullet', label: 'Маркированный список', keys: ['Ctrl+Shift+KeyL'], run: () => dispatchCmd(toggleList(schema.nodes.bullet_list)), checked: () => listKind(view.state) === 'bullet' },
  { id: 'list.ordered', label: 'Нумерованный список', run: () => dispatchCmd(toggleList(schema.nodes.ordered_list)), checked: () => listKind(view.state) === 'ordered' },
  { id: 'indent.inc', label: 'Увеличить отступ', keys: ['Ctrl+KeyM'], run: () => dispatchCmd(indent(1)) },
  { id: 'indent.dec', label: 'Уменьшить отступ', keys: ['Ctrl+Shift+KeyM'], run: () => dispatchCmd(indent(-1)) },
  { id: 'para.line', label: 'Межстрочный интервал', run: (v) => dispatchCmd(setParaAttrs({ line: v && v !== '1' ? parseFloat(v) : null })) },
  { id: 'para.after', label: 'Интервал после абзаца', run: (v) => dispatchCmd(setParaAttrs({ after: v === '' || v === undefined ? null : parseFloat(v) })) },
  { id: 'para.style', label: 'Стиль абзаца', run: (v) => dispatchCmd(setStyle(+(v ?? 0))) },
  { id: 'edit.find', label: 'Найти', keys: ['Ctrl+KeyF'], run: () => openFind(false) },
  { id: 'edit.replace', label: 'Заменить', keys: ['Ctrl+KeyH'], run: () => openFind(true) },
  { id: 'insert.table', label: 'Таблица', run: () => tableDialog() },
  { id: 'insert.image', label: 'Рисунок', run: () => $('image-input').click() },
  { id: 'insert.link', label: 'Гиперссылка', keys: ['Ctrl+KeyK'], run: () => linkDialog() },
  { id: 'insert.pageBreak', label: 'Разрыв страницы', run: () => dispatchCmd(insertPageBreak) },
  { id: 'table.rowBefore', label: 'Вставить строку выше', run: () => dispatchCmd(addRowBefore) },
  { id: 'table.rowAfter', label: 'Вставить строку ниже', run: () => dispatchCmd(addRowAfter) },
  { id: 'table.colBefore', label: 'Вставить столбец слева', run: () => dispatchCmd(addColumnBefore) },
  { id: 'table.colAfter', label: 'Вставить столбец справа', run: () => dispatchCmd(addColumnAfter) },
  { id: 'table.delRow', label: 'Удалить строку', run: () => dispatchCmd(deleteRow) },
  { id: 'table.delCol', label: 'Удалить столбец', run: () => dispatchCmd(deleteColumn) },
  { id: 'table.delete', label: 'Удалить таблицу', run: () => dispatchCmd(deleteTable) },
  { id: 'table.merge', label: 'Объединить ячейки', run: () => dispatchCmd(mergeCells) },
  { id: 'table.split', label: 'Разделить ячейку', run: () => dispatchCmd(splitCell) },
  { id: 'table.borders', label: 'Границы таблицы', run: () => toggleBorders() },
  { id: 'page.size', label: 'Размер страницы', run: (v) => setPageSize(v ?? 'a4') },
  { id: 'page.orientation', label: 'Ориентация', run: (v) => setOrientation(v === 'landscape') },
  { id: 'page.margins', label: 'Поля', run: (v) => setMargins(v ?? 'normal') },
  { id: 'view.zoomIn', label: 'Увеличить масштаб', run: () => setZoom(zoom + 0.1) },
  { id: 'view.zoomOut', label: 'Уменьшить масштаб', run: () => setZoom(zoom - 0.1) },
  { id: 'view.zoomReset', label: 'Масштаб 100%', run: () => setZoom(1) },
  { id: 'file.new', label: 'Создать', keys: ['Ctrl+KeyN'], run: () => (desktop.isDesktop ? void desktop.newWindow() : void guardUnsaved(newDocument)) },
  { id: 'file.open', label: 'Открыть', keys: ['Ctrl+KeyO'], run: () => (desktop.isDesktop ? void desktop.openDialog() : void guardUnsaved(() => $('file-input').click())) },
  { id: 'file.save', label: 'Сохранить', keys: ['Ctrl+KeyS'], run: () => void save() },
  { id: 'file.saveAs', label: 'Сохранить как', keys: ['Ctrl+Shift+KeyS'], run: () => void save(true) },
  { id: 'file.print', label: 'Печать', keys: ['Ctrl+KeyP'], run: () => window.print() },
);

// ---------------------------------------------------------------- вкладки ленты

const PAGE_SIZES = [
  { value: 'a4', label: 'A4' }, { value: 'a5', label: 'A5' }, { value: 'a3', label: 'A3' },
  { value: 'letter', label: 'Letter' }, { value: 'legal', label: 'Legal' },
];
const MARGINS = [
  { value: 'normal', label: 'Поля: обычные' }, { value: 'narrow', label: 'Поля: узкие' },
  { value: 'moderate', label: 'Поля: средние' }, { value: 'wide', label: 'Поля: широкие' },
];
const LINES = [
  { value: '1', label: 'Интервал 1,0' }, { value: '1.15', label: 'Интервал 1,15' }, { value: '1.5', label: 'Интервал 1,5' }, { value: '2', label: 'Интервал 2,0' },
];
const STYLES = [
  { value: '0', label: 'Обычный' }, { value: '1', label: 'Заголовок 1' }, { value: '2', label: 'Заголовок 2' }, { value: '3', label: 'Заголовок 3' },
  { value: '4', label: 'Заголовок 4' }, { value: '5', label: 'Заголовок 5' }, { value: '6', label: 'Заголовок 6' },
];

const TABS: RibbonTab[] = [
  {
    id: 'home', label: 'Главная',
    groups: [
      { name: 'Отмена', items: [{ type: 'medium', command: 'edit.undo', icon: 'arrow_undo' }, { type: 'medium', command: 'edit.redo', icon: 'arrow_redo' }] },
      {
        name: 'Буфер обмена',
        items: [
          { type: 'large', command: 'clipboard.paste', icon: 'clipboard_paste', label: 'Вставить' },
          { type: 'small', command: 'clipboard.cut', icon: 'cut' },
          { type: 'small', command: 'clipboard.copy', icon: 'copy' },
        ],
      },
      {
        name: 'Шрифт',
        items: [
          { type: 'select', command: 'format.fontName', options: FONTS, width: 150 },
          { type: 'select', command: 'format.fontSize', options: FONT_SIZES, width: 56 },
          { type: 'small', command: 'format.bold', icon: 'text_bold' },
          { type: 'small', command: 'format.italic', icon: 'text_italic' },
          { type: 'small', command: 'format.underline', icon: 'text_underline' },
          { type: 'small', command: 'format.strike', icon: 'text_strike' },
          { type: 'small', command: 'format.subscript', icon: 'text_sub' },
          { type: 'small', command: 'format.superscript', icon: 'text_sup' },
          { type: 'color', command: 'format.highlight', icon: 'highlight', kind: 'fill' },
          { type: 'color', command: 'format.fontColor', icon: 'text_color', kind: 'font' },
          { type: 'small', command: 'format.clear', icon: 'clear_format' },
        ],
      },
      {
        name: 'Абзац',
        items: [
          { type: 'small', command: 'list.bullet', icon: 'text_bullet_list' },
          { type: 'small', command: 'list.ordered', icon: 'text_number_list' },
          { type: 'small', command: 'indent.dec', icon: 'text_indent_less' },
          { type: 'small', command: 'indent.inc', icon: 'text_indent_more' },
          { type: 'select', command: 'para.line', options: LINES, width: 120 },
          { type: 'small', command: 'align.left', icon: 'text_align_left' },
          { type: 'small', command: 'align.center', icon: 'text_align_center' },
          { type: 'small', command: 'align.right', icon: 'text_align_right' },
          { type: 'small', command: 'align.justify', icon: 'text_align_justify' },
        ],
      },
      { name: 'Стили', items: [{ type: 'select', command: 'para.style', options: STYLES, width: 130 }] },
      {
        name: 'Редактирование',
        items: [
          { type: 'medium', command: 'edit.find', icon: 'search' },
          { type: 'medium', command: 'edit.replace', icon: 'arrow_sort' },
        ],
      },
    ],
  },
  {
    id: 'insert', label: 'Вставка',
    groups: [
      {
        name: 'Страницы',
        items: [{ type: 'large', command: 'insert.pageBreak', icon: 'page_break', label: 'Разрыв страницы' }],
      },
      {
        name: 'Таблицы и рисунки',
        items: [
          { type: 'large', command: 'insert.table', icon: 'table', label: 'Таблица' },
          { type: 'large', command: 'insert.image', icon: 'image', label: 'Рисунок' },
        ],
      },
      { name: 'Ссылки', items: [{ type: 'large', command: 'insert.link', icon: 'link', label: 'Ссылка' }] },
    ],
  },
  {
    id: 'layout', label: 'Макет',
    groups: [
      {
        name: 'Параметры страницы',
        items: [
          { type: 'select', command: 'page.size', options: PAGE_SIZES, width: 90 },
          { type: 'select', command: 'page.orientation', options: [{ value: 'portrait', label: 'Книжная' }, { value: 'landscape', label: 'Альбомная' }], width: 100 },
          { type: 'select', command: 'page.margins', options: MARGINS, width: 150 },
        ],
      },
      {
        name: 'Абзац',
        items: [{ type: 'select', command: 'para.after', options: [{ value: '', label: 'После абзаца: как в документе' }, { value: '0', label: 'После: 0 пт' }, { value: '6', label: 'После: 6 пт' }, { value: '8', label: 'После: 8 пт' }, { value: '12', label: 'После: 12 пт' }], width: 210 }],
      },
    ],
  },
  {
    id: 'table', label: 'Таблица',
    groups: [
      {
        name: 'Строки и столбцы',
        items: [
          { type: 'large', command: 'table.rowBefore', icon: 'row_add', label: 'Строка выше' },
          { type: 'large', command: 'table.rowAfter', icon: 'row_add', label: 'Строка ниже' },
          { type: 'large', command: 'table.colBefore', icon: 'col_add', label: 'Столбец слева' },
          { type: 'large', command: 'table.colAfter', icon: 'col_add', label: 'Столбец справа' },
          { type: 'large', command: 'table.delRow', icon: 'row_del', label: 'Удалить строку' },
          { type: 'large', command: 'table.delCol', icon: 'col_del', label: 'Удалить столбец' },
        ],
      },
      {
        name: 'Ячейки',
        items: [
          { type: 'large', command: 'table.merge', icon: 'merge', label: 'Объединить' },
          { type: 'large', command: 'table.split', icon: 'cell_split', label: 'Разделить' },
          { type: 'large', command: 'table.borders', icon: 'border_all', label: 'Границы' },
          { type: 'large', command: 'table.delete', icon: 'eraser', label: 'Удалить таблицу' },
        ],
      },
    ],
  },
  {
    id: 'view', label: 'Вид',
    groups: [
      {
        name: 'Масштаб',
        items: [
          { type: 'large', command: 'view.zoomIn', icon: 'zoom_in', label: 'Увеличить' },
          { type: 'large', command: 'view.zoomOut', icon: 'zoom_in', label: 'Уменьшить' },
          { type: 'large', command: 'view.zoomReset', icon: 'document_text', label: '100%' },
        ],
      },
    ],
  },
];

const refreshRibbon = renderRibbon($('tabs'), $('ribbon'), () => fileMenu(), TABS);

function setSelect(cmd: string, value: string): void {
  const sel = document.querySelector<HTMLSelectElement>(`select[data-command="${cmd}"]`);
  if (!sel || sel.value === value) return;
  if (![...sel.options].some((o) => o.value === value)) sel.add(new Option(value, value));
  sel.value = value;
}

const TABLE_CMDS = new Set(['table.rowBefore', 'table.rowAfter', 'table.colBefore', 'table.colAfter', 'table.delRow', 'table.delCol', 'table.merge', 'table.split', 'table.borders', 'table.delete']);

function updateUi(): void {
  refreshRibbon();
  const dirty = isDirty();
  $('doc-name').textContent = `${docName}${dirty ? ' •' : ''} — ENS-Tekst`;
  document.title = `${docName}${dirty ? '*' : ''} — ENS-Tekst`;
  const canUndo = undoDepth(view.state) > 0;
  const canRedo = redoDepth(view.state) > 0;
  ($('qa-undo') as HTMLButtonElement).disabled = !canUndo;
  ($('qa-redo') as HTMLButtonElement).disabled = !canRedo;
  document.querySelector('.btn[data-command="edit.undo"]')?.classList.toggle('disabled', !canUndo);
  document.querySelector('.btn[data-command="edit.redo"]')?.classList.toggle('disabled', !canRedo);
  const inTable = isInTable(view.state);
  document.querySelectorAll<HTMLElement>('.btn[data-command]').forEach((b) => {
    if (TABLE_CMDS.has(b.dataset.command!)) b.classList.toggle('disabled', !inTable);
  });
  const st = currentStyle(view.state);
  setSelect('format.fontName', st.font ?? meta.defaults.font);
  setSelect('format.fontSize', String(st.size ?? meta.defaults.size));
  setSelect('para.style', String(currentLevel(view.state)));
  setSelect('para.line', String(paraAttr<number>(view.state, 'line') ?? meta.defaults.line));
  setSelect('page.orientation', meta.page.w > meta.page.h ? 'landscape' : 'portrait');
  setSelect('page.size', sizeKey());
  setSelect('page.margins', marginsKey());
}

// --- Страница: размеры и поля

const SIZES: Record<string, [number, number]> = { a4: [595.3, 841.9], a5: [419.5, 595.3], a3: [841.9, 1190.6], letter: [612, 792], legal: [612, 1008] };
const cm = (v: number) => Math.round(v * 28.3465 * 10) / 10;
const MARGIN_SETS: Record<string, [number, number, number, number]> = {
  normal: [cm(2), cm(2), cm(3), cm(1.5)], narrow: [cm(1.27), cm(1.27), cm(1.27), cm(1.27)],
  moderate: [cm(2.54), cm(2.54), cm(1.91), cm(1.91)], wide: [cm(2.54), cm(2.54), cm(5.08), cm(5.08)],
};

function sizeKey(): string {
  const w = Math.min(meta.page.w, meta.page.h);
  const h = Math.max(meta.page.w, meta.page.h);
  for (const [k, [sw, sh]] of Object.entries(SIZES)) if (Math.abs(sw - w) < 2 && Math.abs(sh - h) < 2) return k;
  return 'a4';
}

function marginsKey(): string {
  const p = meta.page;
  for (const [k, [t, b, l, r]] of Object.entries(MARGIN_SETS)) {
    if ([p.top - t, p.bottom - b, p.left - l, p.right - r].every((d) => Math.abs(d) < 1.5)) return k;
  }
  return 'normal';
}

function changePage(next: Partial<PageSetup>): void {
  meta.page = { ...meta.page, ...next };
  applyPage();
  replan(view);
  updateUi();
  scheduleRecovery();
}

function setPageSize(key: string): void {
  const [w, h] = SIZES[key] ?? SIZES.a4;
  const land = meta.page.w > meta.page.h;
  changePage(land ? { w: h, h: w } : { w, h });
}

function setOrientation(landscape: boolean): void {
  const lo = Math.min(meta.page.w, meta.page.h);
  const hi = Math.max(meta.page.w, meta.page.h);
  changePage(landscape ? { w: hi, h: lo } : { w: lo, h: hi });
}

function setMargins(key: string): void {
  const [top, bottom, left, right] = MARGIN_SETS[key] ?? MARGIN_SETS.normal;
  changePage({ top, bottom, left, right });
}

// ---------------------------------------------------------------- окна

function tableDialog(): void {
  dialog((box, close) => {
    const rows = el('input', { type: 'number', value: '3', min: '1', max: '100' });
    const cols = el('input', { type: 'number', value: '3', min: '1', max: '30' });
    const borders = el('input', { type: 'checkbox', checked: true });
    const ok = el('button', { class: 'primary', textContent: 'Вставить' });
    const go = () => {
      const r = Math.min(100, Math.max(1, +rows.value || 1));
      const c = Math.min(30, Math.max(1, +cols.value || 1));
      close();
      dispatchCmd(insertTable(r, c, borders.checked));
    };
    ok.addEventListener('click', go);
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
    box.append(
      el('h2', { textContent: 'Вставка таблицы' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Строк' }), rows),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Столбцов' }), cols),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'С границами' }), borders),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), ok),
    );
  });
}

function linkDialog(): void {
  dialog((box, close) => {
    const url = el('input', { type: 'text', value: linkAt(view.state) ?? 'https://', placeholder: 'Адрес ссылки' });
    const ok = el('button', { class: 'primary', textContent: 'Применить' });
    const go = () => {
      const v = url.value.trim();
      close();
      dispatchCmd(setLink(v === 'https://' ? '' : v));
    };
    ok.addEventListener('click', go);
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
    box.append(
      el('h2', { textContent: 'Гиперссылка' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Адрес' }), url),
      el('p', { textContent: 'Пустой адрес убирает ссылку с выделенного текста. Открыть ссылку: Ctrl+щелчок.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), ok),
    );
  });
}

function toggleBorders(): void {
  const { $from } = view.state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type === schema.nodes.table) {
      view.dispatch(view.state.tr.setNodeMarkup($from.before(d), undefined, { ...n.attrs, borders: !n.attrs.borders }));
      view.focus();
      return;
    }
  }
}

// ---------------------------------------------------------------- рисунки

async function insertImageFile(file: File): Promise<void> {
  const src = await new Promise<string>((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result as string);
    r.onerror = () => rej(r.error);
    r.readAsDataURL(file);
  });
  const dim = await new Promise<{ w: number; h: number }>((res) => {
    const img = new Image();
    img.onload = () => res({ w: img.naturalWidth || 200, h: img.naturalHeight || 150 });
    img.onerror = () => res({ w: 200, h: 150 });
    img.src = src;
  });
  // Не шире рабочей области страницы (96 dpi: 1 pt = 4/3 px).
  const maxW = ((meta.page.w - meta.page.left - meta.page.right) * 4) / 3;
  const k = dim.w > maxW ? maxW / dim.w : 1;
  dispatchCmd(insertImage(src, Math.round(dim.w * k), Math.round(dim.h * k), file.name.replace(/\.[^.]+$/, '')));
}

$('image-input').addEventListener('change', (e) => {
  const input = e.target as HTMLInputElement;
  const f = input.files?.[0];
  input.value = '';
  if (f) void insertImageFile(f);
});

// ---------------------------------------------------------------- поиск и замена

const findbar = $('findbar');

function openFind(withReplace: boolean): void {
  findbar.hidden = false;
  findbar.replaceChildren();
  const q = el('input', { type: 'text', placeholder: 'Найти', value: findKey.getState(view.state)?.query ?? '' });
  const r = el('input', { type: 'text', placeholder: 'Заменить на' });
  const mc = el('input', { type: 'checkbox' });
  const count = el('span', { class: 'count' });
  const sync = () => {
    view.dispatch(view.state.tr.setMeta(findKey, { query: q.value, matchCase: mc.checked, current: -1 }));
    const s = findKey.getState(view.state)!;
    count.textContent = q.value ? `Найдено: ${s.matches.length}` : '';
  };
  q.addEventListener('input', sync);
  mc.addEventListener('change', sync);
  q.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); findStep(view, e.shiftKey ? -1 : 1); }
    if (e.key === 'Escape') closeFind();
  });
  r.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeFind(); });
  const replaceOne = () => {
    const s = findKey.getState(view.state)!;
    let idx = s.current;
    const sel = view.state.selection;
    if (idx < 0 || s.matches[idx]?.from !== sel.from || s.matches[idx]?.to !== sel.to) { findStep(view, 1); return; }
    replaceMatch(view, idx, r.value);
    sync();
    findStep(view, 1);
    idx = -1;
  };
  const buttons = [
    el('button', { textContent: '↑', title: 'Предыдущее (Shift+Enter)', onclick: () => findStep(view, -1) }),
    el('button', { textContent: '↓', title: 'Следующее (Enter)', onclick: () => findStep(view, 1) }),
  ];
  const kids: (Node | string)[] = [q, ...buttons, el('label', {}, mc, ' Учитывать регистр'), count];
  if (withReplace) {
    kids.push(r, el('button', { textContent: 'Заменить', onclick: replaceOne }), el('button', {
      textContent: 'Заменить все',
      onclick: () => { const n = replaceAll(view, r.value); sync(); toast(`Заменено: ${n}`); },
    }));
  }
  kids.push(el('button', { textContent: '✕', title: 'Закрыть (Esc)', onclick: closeFind }));
  findbar.append(...kids);
  const sel = view.state.selection;
  if (!sel.empty && sel.$from.sameParent(sel.$to)) q.value = view.state.doc.textBetween(sel.from, sel.to);
  sync();
  q.focus();
  q.select();
}

function closeFind(): void {
  findbar.hidden = true;
  view.dispatch(view.state.tr.setMeta(findKey, { query: '', current: -1 }));
  view.focus();
}

// ---------------------------------------------------------------- файлы

function showNotice(text: string, actions: { label: string; run: () => void }[] = []): void {
  const n = $('notice');
  n.replaceChildren(el('span', { textContent: text }), ...actions.map((a) => el('button', { textContent: a.label, onclick: a.run })));
  n.hidden = false;
}

function hideNotice(): void {
  $('notice').hidden = true;
}

function loadDoc(doc: PMNode, newMetaValue: DocMeta, name: string, dirty = false): void {
  meta = newMetaValue;
  docName = name;
  applyPage();
  view.updateState(createState(doc, [paginationPlugin]));
  savedDoc = dirty ? emptyDoc() : doc;
  savedMeta = dirty ? '' : metaKey();
  hideNotice();
  if (meta.lossy.length) {
    const tail = desktop.isDesktop
      ? 'Поверх исходного файла сохранять нельзя — «Сохранить» предложит сохранить копию, в ней этих элементов не будет.'
      : 'Оригинал не изменится: «Сохранить» скачивает новую копию, в ней этих элементов не будет.';
    showNotice(`В этой версии не перенесено: ${meta.lossy.join('; ')}. ${tail}`, [{ label: 'Понятно', run: hideNotice }]);
  }
  replan(view);
  updateStats();
  updateUi();
  view.focus();
}

function newDocument(): void {
  void clearRecovery(docId);
  loadDoc(emptyDoc(), newMeta(), 'Документ1');
}

/** Если есть несохранённые правки — спрашивает; продолжает, когда можно. */
async function guardUnsaved(next: () => void): Promise<void> {
  if (!isDirty()) { next(); return; }
  dialog((box, close) => {
    box.append(
      el('h2', { textContent: 'Сохранить изменения?' }),
      el('p', { textContent: `В документе «${docName}» есть несохранённые изменения.` }),
      el('div', { class: 'dlg-actions' },
        el('button', { class: 'primary', textContent: 'Сохранить', onclick: async () => { close(); await save(); next(); } }),
        el('button', { class: 'danger', textContent: 'Не сохранять', onclick: () => { close(); next(); } }),
        el('button', { textContent: 'Отмена', onclick: close })),
    );
  });
}

$('file-input').addEventListener('change', async (e) => {
  const input = e.target as HTMLInputElement;
  const f = input.files?.[0];
  input.value = '';
  if (!f) return;
  try {
    const { doc, meta: m } = importDocx(new Uint8Array(await f.arrayBuffer()));
    void clearRecovery(docId);
    loadDoc(doc, m, f.name.replace(/\.docx$/i, ''));
  } catch (err) {
    dialog((box, close) => {
      box.append(
        el('h2', { textContent: 'Не удалось открыть файл' }),
        el('p', { textContent: `${f.name}: ${(err as Error).message}` }),
        el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })),
      );
    });
  }
});

/** Сохранение: .docx собирается целиком и проверяется чтением, прежде чем отдаётся пользователю. */
async function save(asNew = false): Promise<void> {
  if (desktop.isDesktop) { await desktop.save(asNew); return; }
  try {
    let name = docName;
    if (asNew || docName === 'Документ1') {
      const entered = await askName(docName);
      if (!entered) return;
      name = entered;
    }
    const bytes = exportDocx(view.state.doc, meta, name);
    importDocx(bytes); // проверка: только что собранный файл должен читаться
    const blob = new Blob([bytes as Uint8Array<ArrayBuffer>], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${name}.docx`;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
    docName = name;
    savedDoc = view.state.doc;
    savedMeta = metaKey();
    void clearRecovery(docId);
    updateUi();
    toast(`Сохранено: ${name}.docx`);
  } catch (err) {
    dialog((box, close) => {
      box.append(
        el('h2', { textContent: 'Не удалось сохранить' }),
        el('p', { textContent: `${(err as Error).message}. Документ остался открытым и не потерян.` }),
        el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })),
      );
    });
  }
}

function askName(current: string): Promise<string | null> {
  return new Promise((resolve) => {
    dialog((box, close) => {
      const name = el('input', { type: 'text', value: current });
      const done = (v: string | null) => { close(); resolve(v); };
      const ok = el('button', { class: 'primary', textContent: 'Сохранить', onclick: () => done(name.value.trim().replace(/[\\/:*?"<>|]/g, '_') || 'Документ') });
      box.addEventListener('keydown', (e) => { if (e.key === 'Enter') ok.click(); if (e.key === 'Escape') resolve(null); });
      box.append(
        el('h2', { textContent: 'Сохранить как' }),
        el('div', { class: 'dlg-row' }, el('label', { textContent: 'Имя файла' }), name, el('span', { textContent: '.docx' })),
        el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: () => done(null) }), ok),
      );
    });
  });
}

function fileMenu(): void {
  dialog((box, close) => {
    const item = (label: string, id: string) => el('button', { textContent: label, onclick: () => { close(); run(id); } });
    box.append(
      el('h2', { textContent: 'Файл' }),
      el('div', { class: 'menu-list' },
        item('Создать', 'file.new'), item('Открыть…', 'file.open'), item('Сохранить', 'file.save'), item('Сохранить как…', 'file.saveAs'), item('Печать…', 'file.print'),
        el('button', { textContent: 'О программе', onclick: () => { close(); about(); } })),
    );
  });
}

function about(): void {
  dialog((box, close) => {
    box.append(
      el('h2', { textContent: `ENS-Tekst ${APP_VERSION}` }),
      el('p', { textContent: 'Текстовый редактор ENS-Office. Работает без интернета. Формат файлов — .docx.' }),
      el('p', { textContent: 'Версия 0.1 — первый этап: абзацы, заголовки, списки, таблицы, изображения, ссылки, страницы, печать. Колонтитулы, сноски и совместная работа — в планах.' }),
      el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })),
    );
  });
}

// ---------------------------------------------------------------- восстановление

function scheduleRecovery(): void {
  clearTimeout(recoveryTimer);
  recoveryTimer = window.setTimeout(persistRecovery, 1500);
}

function persistRecovery(): void {
  if (desktop.isDesktop) { void desktop.persist(); return; }
  if (!isDirty()) { void clearRecovery(docId); return; }
  void saveRecovery({ id: docId, name: docName, savedAt: Date.now(), doc: view.state.doc.toJSON(), meta });
}

document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') persistRecovery(); });
// В настольной версии вопрос о сохранении и копию для восстановления ведёт desktop.ts.
if (!desktop.isDesktop) {
  window.addEventListener('beforeunload', (e) => {
    if (isDirty()) { persistRecovery(); e.preventDefault(); e.returnValue = ''; }
  });
}

async function offerRecovery(): Promise<void> {
  let list: RecoveryRecord[] = [];
  try { list = await listRecovery(); } catch { return; }
  list = list.filter((r) => r.id !== docId).sort((a, b) => b.savedAt - a.savedAt);
  if (!list.length) return;
  dialog((box, close) => {
    box.append(el('h2', { textContent: 'Восстановление документов' }), el('p', { textContent: 'Найдены документы, закрытые без сохранения. Восстановленная версия откроется в этом окне; исходные файлы не затронуты.' }));
    for (const r of list) {
      const when = new Date(r.savedAt).toLocaleString('ru-RU');
      box.append(el('div', { class: 'rec-item' },
        el('div', { class: 'rec-title', textContent: r.name }),
        el('div', { class: 'rec-meta', textContent: `Последняя копия: ${when}` }),
        el('div', { class: 'rec-buttons' },
          el('button', {
            class: 'primary', textContent: 'Восстановить',
            onclick: () => {
              try {
                const doc = schema.nodeFromJSON(r.doc);
                doc.check();
                close();
                loadDoc(doc, r.meta, `${r.name} (восстановлено)`, true);
                void clearRecovery(r.id);
                persistRecovery();
              } catch {
                toast('Копию прочитать не удалось');
              }
            },
          }),
          el('button', { class: 'danger', textContent: 'Удалить копию', onclick: async () => { await clearRecovery(r.id); close(); void offerRecovery(); } }),
        )));
    }
    box.append(el('div', { class: 'rec-actions' }, el('button', { textContent: 'Решить позже', onclick: close })));
  }, false);
}

// ---------------------------------------------------------------- запуск

$('qa-save').innerHTML = ICONS.save_20 ?? '';
$('qa-print').innerHTML = ICONS.print_20 ?? '';
$('qa-undo').innerHTML = ICONS.arrow_undo_20 ?? '';
$('qa-redo').innerHTML = ICONS.arrow_redo_20 ?? '';
$('qa-save').addEventListener('click', () => run('file.save'));
$('qa-print').addEventListener('click', () => run('file.print'));
$('qa-undo').addEventListener('click', () => run('edit.undo'));
$('qa-redo').addEventListener('click', () => run('edit.redo'));
$('zoom-in').addEventListener('click', () => setZoom(zoom + 0.1));
$('zoom-out').addEventListener('click', () => setZoom(zoom - 0.1));
$('zoom-val').addEventListener('click', () => setZoom(1));

document.addEventListener('keydown', (e) => {
  if (e.isComposing) return;
  handleKey(e);
}, true);
// Масштаб страницы браузера запрещён: Ctrl+колесо меняет масштаб документа.
$('desk').addEventListener('wheel', (e) => {
  if (!e.ctrlKey) return;
  e.preventDefault();
  setZoom(zoom + (e.deltaY < 0 ? 0.1 : -0.1));
}, { passive: false });
window.addEventListener('wheel', (e) => { if (e.ctrlKey) e.preventDefault(); }, { passive: false });

// Выпадающие списки ленты должны отдавать фокус обратно в текст.
$('ribbon').addEventListener('change', () => setTimeout(() => view.focus(), 0));

applyPage();
setZoom(1);
updateStats();
updateUi();

await loadFonts();
applyPage();
setRestoreFocus(() => view.focus());
if (desktop.isDesktop) {
  await desktop.start({
    name: () => docName,
    isDirty,
    isPristine: () => !isDirty() && view.state.doc.childCount === 1 && view.state.doc.firstChild!.content.size === 0,
    docxBytes: (name) => {
      const bytes = exportDocx(view.state.doc, meta, name);
      importDocx(bytes); // проверка: только что собранный файл должен читаться
      return bytes;
    },
    snapshotJson: () => JSON.stringify({ doc: view.state.doc.toJSON(), meta, name: docName }),
    openBytes: (bytes, name) => {
      const { doc, meta: m } = importDocx(bytes);
      loadDoc(doc, m, name);
      return m.lossy;
    },
    restoreJson: (json, title) => {
      const r = JSON.parse(json) as { doc: Record<string, unknown>; meta: DocMeta };
      const doc = schema.nodeFromJSON(r.doc);
      doc.check();
      loadDoc(doc, r.meta, title, true);
    },
    markSaved: (name) => {
      docName = name;
      savedDoc = view.state.doc;
      savedMeta = metaKey();
      updateUi();
      toast(`Сохранено: ${name}.docx`);
    },
  });
} else {
  void offerRecovery();
}

// Для автоматических проверок (Playwright): доступ к редактору из консоли.
if (import.meta.env.DEV || import.meta.env.MODE === 'test') {
  (window as unknown as Record<string, unknown>).ens = {
    view, schema, pageCount: () => pageCount(view), pagStats, geometry: () => geometry(meta.page), PX, importDocx, exportDocx, loadDoc, getMeta: () => meta, isDirty, run, A4, lastColor,
  };
}
