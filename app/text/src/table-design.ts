// SPDX-License-Identifier: AGPL-3.0-or-later
// «Конструктор таблиц» и «Макет» (по образцу Word): стили с параметрами, перо и ластик границ, поля ячеек,
// направление текста, выравнивание в ячейке, выделение строки/столбца/таблицы, автоподбор, сортировка,
// повтор строки заголовка, преобразование в текст.

import { Fragment, type Node as PMNode } from 'prosemirror-model';
import { Plugin, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { CellSelection, TableMap } from 'prosemirror-tables';
import type { EditorView } from 'prosemirror-view';
import { schema, type CellBorders, type Edge } from './schema';
import { tableAt } from './tables-ui';
import { dialog, el, toast } from './ui';

// --- Стили таблиц

export interface Look { style: string; header: boolean; total: boolean; banded: boolean; firstCol: boolean }
interface StyleDef { label: string; head: string | null; headText?: string; band: string | null; total: string | null; first: string | null; line: string }

export const TABLE_STYLES: Record<string, StyleDef> = {
  grid: { label: 'Сетка', head: null, band: null, total: null, first: null, line: '#000000' },
  light: { label: 'Светлая', head: '#D9E2F3', band: '#F2F2F2', total: '#D9E2F3', first: '#F2F2F2', line: '#8EAADB' },
  blue: { label: 'Синяя', head: '#2F5496', headText: '#FFFFFF', band: '#D9E2F3', total: '#B4C6E7', first: '#D9E2F3', line: '#2F5496' },
  green: { label: 'Зелёная', head: '#548235', headText: '#FFFFFF', band: '#E2EFD9', total: '#C5E0B3', first: '#E2EFD9', line: '#548235' },
  gray: { label: 'Серая (ГОСТ)', head: '#D9D9D9', band: null, total: '#D9D9D9', first: null, line: '#000000' },
};
export const DEFAULT_LOOK: Look = { style: 'light', header: true, total: false, banded: true, firstCol: false };

/** Перекрасить таблицу по стилю и параметрам: заливка ячеек, полужирный в строке заголовка и итогов, линии. */
export function applyLook(state: EditorState, look: Look): Transaction | null {
  const t = tableAt(state);
  if (!t) return null;
  const def = TABLE_STYLES[look.style] ?? TABLE_STYLES.grid;
  const prev = t.node.attrs.look as Look | null;
  const tr = state.tr;
  const map = TableMap.get(t.node);
  const start = t.pos + 1;
  const edge: Edge = { w: 0.5, c: def.line, s: 'solid' };
  const rows = t.node.childCount;
  const strong = schema.marks.strong;
  const colorMark = (c: string) => schema.marks.textStyle.create({ color: c });
  for (let r = 0; r < rows; r++) {
    const isHead = look.header && r === 0, isTotal = look.total && r === rows - 1 && rows > 1;
    const wasHead = !!prev?.header && r === 0, wasTotal = !!prev?.total && r === rows - 1 && rows > 1;
    for (let c = 0; c < map.width; c++) {
      const rel = map.map[r * map.width + c];
      const rect = map.findCell(rel);
      if (rect.left !== c || rect.top !== r) continue;
      const pos = start + rel;
      const cell = tr.doc.nodeAt(tr.mapping.map(pos))!;
      let bg: string | null = null;
      if (isHead) bg = def.head;
      else if (isTotal) bg = def.total;
      else if (look.firstCol && c === 0) bg = def.first;
      else if (look.banded && (look.header ? r % 2 === 0 : r % 2 === 1)) bg = def.band;
      const borders: CellBorders = { t: edge, r: edge, b: edge, l: edge };
      const at = tr.mapping.map(pos);
      tr.setNodeMarkup(at, undefined, { ...cell.attrs, background: bg, borders: def.line === '#000000' ? null : borders });
      const from = at + 1, to = at + cell.nodeSize - 1;
      // Полужирный: в строках заголовка, итогов и в первом столбце — ставим; где их больше нет — снимаем.
      const bold = isHead || isTotal || (look.firstCol && c === 0 && !!def.first);
      const wasBold = wasHead || wasTotal || (!!prev?.firstCol && c === 0);
      if (bold) tr.addMark(from, to, strong.create());
      else if (wasBold) tr.removeMark(from, to, strong);
      if (isHead && def.headText) tr.addMark(from, to, colorMark(def.headText));
      else if (wasHead && prev && TABLE_STYLES[prev.style]?.headText) tr.removeMark(from, to, schema.marks.textStyle);
    }
  }
  tr.setNodeMarkup(t.pos, undefined, { ...tr.doc.nodeAt(t.pos)!.attrs, look, borders: true });
  // Строка заголовка повторяется на страницах — как в Word при стиле с заголовком.
  const head = tr.doc.nodeAt(t.pos)!.firstChild;
  if (head) tr.setNodeMarkup(t.pos + 1, undefined, { ...head.attrs, header: look.header });
  return tr;
}

export function currentLook(state: EditorState): Look | null {
  const t = tableAt(state);
  return t ? (t.node.attrs.look as Look | null) : null;
}

// --- Выделение

export function selectPart(view: EditorView, what: 'cell' | 'row' | 'col' | 'table'): boolean {
  const t = tableAt(view.state);
  if (!t) return false;
  const { $from } = view.state.selection;
  let cellPos = -1;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type === schema.nodes.table_cell || n.type === schema.nodes.table_header) { cellPos = $from.before(d); break; }
  }
  if (cellPos < 0) return false;
  const $cell = view.state.doc.resolve(cellPos);
  let sel: CellSelection;
  if (what === 'cell') sel = new CellSelection($cell);
  else if (what === 'row') sel = CellSelection.rowSelection($cell);
  else if (what === 'col') sel = CellSelection.colSelection($cell);
  else {
    const map = TableMap.get(t.node);
    sel = new CellSelection(view.state.doc.resolve(t.pos + 1 + map.map[0]), view.state.doc.resolve(t.pos + 1 + map.map[map.map.length - 1]));
  }
  view.dispatch(view.state.tr.setSelection(sel));
  return true;
}

/** Ячейки выделения (или ячейка под курсором). */
function selectedCells(state: EditorState): { pos: number; node: PMNode }[] {
  const sel = state.selection;
  const out: { pos: number; node: PMNode }[] = [];
  if (sel instanceof CellSelection) { sel.forEachCell((node, pos) => out.push({ node, pos })); return out; }
  const { $from } = sel;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type === schema.nodes.table_cell || n.type === schema.nodes.table_header) return [{ node: n, pos: $from.before(d) }];
  }
  return out;
}

// --- Ячейки: выравнивание (9 вариантов), направление текста, поля

export function alignCells(state: EditorState, h: 'left' | 'center' | 'right', v: 'top' | 'middle' | 'bottom'): Transaction | null {
  const cells = selectedCells(state);
  if (!cells.length) return null;
  const tr = state.tr;
  for (const c of cells) {
    tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, valign: v === 'top' ? null : v });
    c.node.descendants((n, p) => {
      if (n.isTextblock) tr.setNodeMarkup(c.pos + 1 + p, undefined, { ...n.attrs, align: h === 'left' ? null : h });
      return !n.isTextblock;
    });
  }
  return tr;
}

/** Направление текста по кругу: обычное → снизу вверх → сверху вниз. */
export function cycleDirection(state: EditorState): Transaction | null {
  const cells = selectedCells(state);
  if (!cells.length) return null;
  const order = [null, 'btLr', 'tbRl'];
  const next = order[(order.indexOf(cells[0].node.attrs.dir as string | null) + 1) % order.length];
  const tr = state.tr;
  for (const c of cells) tr.setNodeMarkup(c.pos, undefined, { ...c.node.attrs, dir: next });
  return tr;
}

export function setRepeatHeader(state: EditorState, on?: boolean): Transaction | null {
  const t = tableAt(state);
  if (!t || !t.node.firstChild) return null;
  const head = t.node.firstChild;
  return state.tr.setNodeMarkup(t.pos + 1, undefined, { ...head.attrs, header: on ?? !head.attrs.header });
}

export const repeatHeaderOn = (state: EditorState) => !!tableAt(state)?.node.firstChild?.attrs.header;

export function setAutofit(state: EditorState, mode: 'content' | 'fixed'): Transaction | null {
  const t = tableAt(state);
  if (!t) return null;
  const tr = state.tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, autofit: mode === 'content' });
  if (mode === 'content') {
    // По содержимому: ширины столбцов убираем — браузер (и Word) подбирают их сами.
    t.node.descendants((n, p) => {
      if (n.type === schema.nodes.table_cell || n.type === schema.nodes.table_header) tr.setNodeMarkup(t.pos + 1 + p, undefined, { ...n.attrs, colwidth: null });
      return n.type === schema.nodes.table_row || n.type === schema.nodes.table;
    });
  }
  return tr;
}

// --- Перо и ластик границ

export interface Pen { mode: 'off' | 'draw' | 'erase'; w: number; s: Edge['s']; c: string }
export const pen: Pen = { mode: 'off', w: 0.5, s: 'solid', c: '#000000' };
const NONE: Edge = { w: 0, c: '#000000', s: 'none' };

/** Граница ячейки у точки щелчка (ближайшая сторона в пределах 8 px) или null. */
function sideAt(td: HTMLElement, x: number, y: number): 't' | 'r' | 'b' | 'l' | null {
  const r = td.getBoundingClientRect();
  const d = { t: Math.abs(y - r.top), b: Math.abs(r.bottom - y), l: Math.abs(x - r.left), r: Math.abs(r.right - x) };
  const best = (Object.entries(d) as [keyof typeof d, number][]).sort((a, b) => a[1] - b[1])[0];
  return best[1] <= 8 ? best[0] : null;
}

/** Задать сторону ячейки и общую сторону соседа (иначе сосед дорисует свою линию). */
export function paintEdge(state: EditorState, cellPos: number, side: 't' | 'r' | 'b' | 'l', e: Edge): Transaction | null {
  const $cell = state.doc.resolve(cellPos);
  const table = $cell.node(-1);
  const tableStart = $cell.start(-1);
  const map = TableMap.get(table);
  const rect = map.findCell(cellPos - tableStart);
  const tr = state.tr;
  const set = (pos: number, s: 't' | 'r' | 'b' | 'l') => {
    const node = tr.doc.nodeAt(pos)!;
    const cur = (node.attrs.borders as CellBorders | null) ?? (table.attrs.borders ? { t: null, r: null, b: null, l: null } : { t: NONE, r: NONE, b: NONE, l: NONE });
    tr.setNodeMarkup(pos, undefined, { ...node.attrs, borders: { ...cur, [s]: e } });
  };
  set(cellPos, side);
  const opp = { t: 'b', b: 't', l: 'r', r: 'l' } as const;
  const nr = side === 't' ? rect.top - 1 : side === 'b' ? rect.bottom : rect.top;
  const nc = side === 'l' ? rect.left - 1 : side === 'r' ? rect.right : rect.left;
  if (nr >= 0 && nr < map.height && nc >= 0 && nc < map.width) {
    const npos = tableStart + map.map[nr * map.width + nc];
    if (npos !== cellPos) set(npos, opp[side]);
  }
  return tr;
}

/** Плагин пера: в режиме рисования или ластика щелчок у границы ячейки меняет её. */
export const penPlugin = new Plugin({
  props: {
    handleDOMEvents: {
      mousedown(view, event) {
        if (pen.mode === 'off' || event.button !== 0) return false;
        const td = (event.target as HTMLElement).closest('td, th') as HTMLElement | null;
        if (!td || !view.dom.contains(td)) return false;
        const side = sideAt(td, event.clientX, event.clientY);
        if (!side) return false;
        const pos = view.posAtDOM(td, 0) - 1;
        const node = view.state.doc.nodeAt(pos);
        if (!node || (node.type !== schema.nodes.table_cell && node.type !== schema.nodes.table_header)) return false;
        const tr = paintEdge(view.state, pos, side, pen.mode === 'erase' ? NONE : { w: pen.w, s: pen.s, c: pen.c });
        if (tr) view.dispatch(tr);
        event.preventDefault();
        return true;
      },
    },
    attributes: (): Record<string, string> => (pen.mode === 'off' ? {} : { class: `pen-${pen.mode}` }),
  },
});

// --- Сортировка и преобразование в текст

/** Сортировка строк по столбцу (как «Сортировка» в Word): числа — как числа, текст — по алфавиту (русскому). */
export function sortRows(state: EditorState, col: number, desc: boolean, hasHeader: boolean): Transaction | string {
  const t = tableAt(state);
  if (!t) return 'Курсор не в таблице.';
  const map = TableMap.get(t.node);
  // Объединения по вертикали ломают порядок строк — такие таблицы не сортируем (как и Word).
  for (let i = 0; i < map.map.length; i++) {
    const rect = map.findCell(map.map[i]);
    if (rect.bottom - rect.top > 1) return 'В таблице есть ячейки, объединённые по вертикали: сортировка невозможна.';
  }
  const rows: PMNode[] = [];
  t.node.forEach((r) => rows.push(r));
  const head = hasHeader ? rows.slice(0, 1) : [];
  const body = rows.slice(head.length);
  const key = (r: PMNode) => {
    const idx = map.map[rows.indexOf(r) * map.width + Math.min(col, map.width - 1)];
    return t.node.nodeAt(idx)?.textContent.trim() ?? '';
  };
  const num = (s: string) => { const v = Number(s.replace(/\s/g, '').replace(',', '.')); return s && !isNaN(v) ? v : null; };
  const sorted = [...body].sort((a, b) => {
    const x = key(a), y = key(b), nx = num(x), ny = num(y);
    const c = nx != null && ny != null ? nx - ny : nx != null ? -1 : ny != null ? 1 : x.localeCompare(y, 'ru', { sensitivity: 'base' });
    return desc ? -c : c;
  });
  const table = t.node.copy(Fragment.fromArray([...head, ...sorted]));
  // Курсор остаётся в таблице (в первой ячейке), чтобы команды «Макета» продолжали работать.
  const tr = state.tr.replaceWith(t.pos, t.pos + t.node.nodeSize, table);
  return tr.setSelection(TextSelection.near(tr.doc.resolve(t.pos + 4)));
}

export function tableToText(state: EditorState, sep: string): Transaction | null {
  const t = tableAt(state);
  if (!t) return null;
  const paras: PMNode[] = [];
  t.node.forEach((row) => {
    const parts: string[] = [];
    row.forEach((cell) => parts.push(cell.textBetween(0, cell.content.size, ' ', ' ').trim()));
    const line = parts.join(sep);
    paras.push(schema.nodes.paragraph.create(null, line ? schema.text(line) : null));
  });
  const tr = state.tr.replaceWith(t.pos, t.pos + t.node.nodeSize, paras);
  return tr.setSelection(TextSelection.create(tr.doc, t.pos + 1));
}

// --- Окна

export function sortDialog(view: EditorView): void {
  const t = tableAt(view.state);
  if (!t) { toast('Поставьте курсор в таблицу.'); return; }
  const map = TableMap.get(t.node);
  const firstRow = t.node.firstChild!;
  dialog((box, close) => {
    const head = el('input', { type: 'checkbox', checked: !!(t.node.attrs.look as Look | null)?.header || !!firstRow.attrs.header });
    const col = el('select', { ariaLabel: 'Столбец' });
    const fill = () => {
      col.replaceChildren(...Array.from({ length: map.width }, (_, i) => {
        const name = head.checked ? t.node.nodeAt(map.map[i])?.textContent.trim() : '';
        return new Option(name ? `${name}` : `Столбец ${i + 1}`, String(i));
      }));
    };
    head.addEventListener('change', fill);
    fill();
    const order = el('select', { ariaLabel: 'Порядок' }, new Option('По возрастанию (А→Я, 1→9)', 'asc'), new Option('По убыванию (Я→А, 9→1)', 'desc'));
    const status = el('p', { role: 'status' });
    box.append(el('h2', { textContent: 'Сортировка' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Сортировать по' }), col),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Порядок' }), order),
      el('label', {}, head, ' первая строка — заголовок (не сортируется)'), status,
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Сортировать', onclick: () => {
        const r = sortRows(view.state, Number(col.value), order.value === 'desc', head.checked);
        if (typeof r === 'string') { status.textContent = r; return; }
        close(); view.dispatch(r); view.focus();
      } })));
  });
}

export function toTextDialog(view: EditorView): void {
  if (!tableAt(view.state)) { toast('Поставьте курсор в таблицу.'); return; }
  dialog((box, close) => {
    const sep = el('select', { ariaLabel: 'Разделитель' }, new Option('Знак табуляции', '\t'), new Option('Точка с запятой', '; '), new Option('Запятая', ', '), new Option('Пробел', ' '));
    box.append(el('h2', { textContent: 'Преобразовать в текст' }),
      el('p', { textContent: 'Каждая строка таблицы станет абзацем, ячейки разделятся выбранным знаком. Отменить — Ctrl+Z.' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Разделитель' }), sep),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Преобразовать', onclick: () => {
        const tr = tableToText(view.state, sep.value);
        close(); if (tr) view.dispatch(tr); view.focus();
      } })));
  });
}

/** Поля ячеек (как «Параметры таблицы → Поля ячеек» в Word), pt. */
export function cellMarginsDialog(view: EditorView): void {
  const t = tableAt(view.state);
  if (!t) { toast('Поставьте курсор в таблицу.'); return; }
  const cur = (t.node.attrs.pad as number[] | null) ?? [1, 5.4, 1, 5.4];
  dialog((box, close) => {
    const inputs = ['Верхнее', 'Правое', 'Нижнее', 'Левое'].map((l, i) => {
      const input = el('input', { type: 'number', step: '0.1', min: '0', max: '72', value: String(Math.round(cur[i] / 28.3465 * 100) / 100), ariaLabel: l });
      return { input, row: el('div', { class: 'dlg-row' }, el('label', { textContent: `${l}, см` }), input) };
    });
    box.append(el('h2', { textContent: 'Поля ячеек' }), ...inputs.map((i) => i.row),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Применить', onclick: () => {
        const pad = inputs.map((i) => Math.max(0, Number(i.input.value) || 0) * 28.3465);
        const tt = tableAt(view.state);
        close();
        if (tt) view.dispatch(view.state.tr.setNodeMarkup(tt.pos, undefined, { ...tt.node.attrs, pad: pad.map((v) => Math.round(v * 10) / 10) }));
        view.focus();
      } })));
  });
}
