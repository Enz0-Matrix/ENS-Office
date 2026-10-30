// Таблицы: ширина столбцов, границы и заливка ячеек, свойства таблицы. Функции ниже нужны и линейке.

import type { Node as PMNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { TableMap } from 'prosemirror-tables';
import type { EditorView } from 'prosemirror-view';

/** Таблица, в которой стоит курсор (в том числе вложенная — ближайшая). */
export function tableAt(state: EditorState): { pos: number; node: PMNode } | null {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const n = $from.node(d);
    if (n.type.name === 'table') return { pos: $from.before(d), node: n };
  }
  return null;
}

/** Ширины столбцов в px (по уже нарисованным ячейкам). */
export function columnWidths(view: EditorView, tablePos: number): number[] {
  const table = view.state.doc.nodeAt(tablePos);
  if (!table) return [];
  const map = TableMap.get(table);
  const widths: (number | null)[] = Array(map.width).fill(null);
  const start = tablePos + 1;
  const seen = new Set<number>();
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i];
    if (seen.has(pos)) continue;
    seen.add(pos);
    const cell = table.nodeAt(pos);
    const dom = view.nodeDOM(start + pos) as HTMLElement | null;
    if (!cell || !dom) continue;
    const span = (cell.attrs.colspan as number) || 1;
    const col = i % map.width;
    if (span === 1) widths[col] = dom.offsetWidth;
  }
  // Столбцы без одиночных ячеек: делим ширину объединённой ячейки поровну.
  const known = widths.filter((w): w is number => w != null);
  const avg = known.length ? known.reduce((a, b) => a + b, 0) / known.length : 100;
  return widths.map((w) => w ?? avg);
}

/** Записать ширины столбцов во все ячейки таблицы (colwidth по числу занимаемых столбцов). */
export function setColumnWidths(view: EditorView, tablePos: number, cols: number[]): void {
  const table = view.state.doc.nodeAt(tablePos);
  if (!table) return;
  const map = TableMap.get(table);
  const tr = view.state.tr;
  const seen = new Set<number>();
  for (let i = 0; i < map.map.length; i++) {
    const pos = map.map[i];
    if (seen.has(pos)) continue;
    seen.add(pos);
    const cell = table.nodeAt(pos);
    if (!cell) continue;
    const span = (cell.attrs.colspan as number) || 1;
    const col = i % map.width;
    tr.setNodeMarkup(tablePos + 1 + pos, undefined, { ...cell.attrs, colwidth: cols.slice(col, col + span).map((w) => Math.round(w)) });
  }
  view.dispatch(tr);
}

// ---------------------------------------------------------------- границы и заливка

import { selectedRect } from 'prosemirror-tables';
import type { CellBorders, Edge } from './schema';
import { dialog, el, toast } from './ui';

const DEF_EDGE: Edge = { w: 0.5, c: '#000000', s: 'solid' };
const NO_EDGE: Edge = { w: 0, c: '#000000', s: 'none' };

export type EdgeSet = { top: boolean; bottom: boolean; left: boolean; right: boolean; insideH: boolean; insideV: boolean };

/** Нарисовать (или убрать, если edge = null) границы в выбранных ячейках или во всей таблице. */
export function applyBorders(view: EditorView, scope: 'cells' | 'table', sides: EdgeSet, edge: Edge | null): boolean {
  const st = view.state;
  let rect;
  try { rect = selectedRect(st); } catch { toast('Поставьте курсор в ячейку таблицы.'); return false; }
  const { map, table, tableStart } = rect;
  const area = scope === 'table' ? { left: 0, top: 0, right: map.width, bottom: map.height } : { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
  const e = edge ?? NO_EDGE;
  const tr = st.tr;
  const cur = new Map<number, CellBorders>();
  const bordersOf = (pos: number): CellBorders => {
    let b = cur.get(pos);
    if (!b) {
      const cell = table.nodeAt(pos)!;
      const own = cell.attrs.borders as CellBorders | null;
      const base = table.attrs.borders ? DEF_EDGE : null;
      b = own ? { ...own } : { t: base, r: base, b: base, l: base };
      cur.set(pos, b);
    }
    return b;
  };
  const set = (pos: number, side: 't' | 'r' | 'b' | 'l') => { bordersOf(pos)[side] = e; };
  const at = (row: number, col: number): number | null => (row < 0 || col < 0 || row >= map.height || col >= map.width ? null : map.map[row * map.width + col]);
  for (let r = area.top; r < area.bottom; r++) {
    for (let c = area.left; c < area.right; c++) {
      const pos = at(r, c)!;
      const cr = map.findCell(pos);
      // Верх и низ ячейки.
      if (cr.top === r) {
        const above = at(r - 1, c);
        if (r === area.top ? sides.top : sides.insideH) { set(pos, 't'); if (above != null) set(above, 'b'); }
      }
      if (cr.bottom === r + 1) {
        const below = at(r + 1, c);
        if (r === area.bottom - 1 ? sides.bottom : sides.insideH) { set(pos, 'b'); if (below != null) set(below, 't'); }
      }
      if (cr.left === c) {
        const left = at(r, c - 1);
        if (c === area.left ? sides.left : sides.insideV) { set(pos, 'l'); if (left != null) set(left, 'r'); }
      }
      if (cr.right === c + 1) {
        const right = at(r, c + 1);
        if (c === area.right - 1 ? sides.right : sides.insideV) { set(pos, 'r'); if (right != null) set(right, 'l'); }
      }
    }
  }
  for (const [pos, b] of cur) {
    const cell = table.nodeAt(pos)!;
    tr.setNodeMarkup(tableStart + pos, undefined, { ...cell.attrs, borders: b });
  }
  view.dispatch(tr);
  return true;
}

/** Ячейки области: выделенные или вся таблица. */
function cellsOf(view: EditorView, scope: 'cells' | 'table'): { tr: Transaction; list: { pos: number; node: PMNode }[]; tableStart: number } | null {
  let rect;
  try { rect = selectedRect(view.state); } catch { return null; }
  const { map, table, tableStart } = rect;
  const area = scope === 'table' ? { left: 0, top: 0, right: map.width, bottom: map.height } : { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom };
  const list = map.cellsInRect(area).map((pos: number) => ({ pos, node: table.nodeAt(pos)! }));
  return { tr: view.state.tr, list, tableStart };
}

export function setCellsAttr(view: EditorView, scope: 'cells' | 'table', name: string, value: unknown): boolean {
  const c = cellsOf(view, scope);
  if (!c) { toast('Поставьте курсор в ячейку таблицы.'); return false; }
  for (const { pos, node } of c.list) c.tr.setNodeMarkup(c.tableStart + pos, undefined, { ...node.attrs, [name]: value });
  view.dispatch(c.tr);
  return true;
}

export function bordersDialog(view: () => EditorView): void {
  if (!tableAt(view().state)) { toast('Поставьте курсор в ячейку таблицы.'); return; }
  dialog((box, close) => {
    const row = (label: string, ...k: Node[]) => el('div', { class: 'dlg-row' }, el('label', { textContent: label }), ...k);
    const scope = el('select', {}, new Option('выделенным ячейкам', 'cells'), new Option('всей таблице', 'table'));
    const sides: Record<keyof EdgeSet, HTMLInputElement> = {
      top: el('input', { type: 'checkbox', checked: true }), bottom: el('input', { type: 'checkbox', checked: true }),
      left: el('input', { type: 'checkbox', checked: true }), right: el('input', { type: 'checkbox', checked: true }),
      insideH: el('input', { type: 'checkbox', checked: true }), insideV: el('input', { type: 'checkbox', checked: true }),
    };
    const preset = (name: string, on: Partial<EdgeSet>) => el('button', { textContent: name, onclick: () => { for (const k of Object.keys(sides) as (keyof EdgeSet)[]) sides[k].checked = !!on[k]; } });
    const kind = el('select', {}, new Option('сплошная', 'solid'), new Option('пунктир', 'dashed'), new Option('точки', 'dotted'), new Option('двойная', 'double'));
    const width = el('select', {}, ...[0.25, 0.5, 0.75, 1, 1.5, 2.25, 3].map((w) => new Option(`${String(w).replace('.', ',')} пт`, String(w))));
    width.value = '0.5';
    const color = el('input', { type: 'color', value: '#000000' });
    const fillOn = el('input', { type: 'checkbox' });
    const fill = el('input', { type: 'color', value: '#ffff99' });
    const apply = (remove: boolean) => {
      const v = view();
      const s: EdgeSet = { top: sides.top.checked, bottom: sides.bottom.checked, left: sides.left.checked, right: sides.right.checked, insideH: sides.insideH.checked, insideV: sides.insideV.checked };
      const edge: Edge | null = remove ? null : { w: parseFloat(width.value), c: color.value.toUpperCase(), s: kind.value as Edge['s'] };
      if (Object.values(s).some(Boolean) && !applyBorders(v, scope.value as 'cells' | 'table', s, edge)) return;
      if (fillOn.checked) setCellsAttr(v, scope.value as 'cells' | 'table', 'background', fill.value.toUpperCase());
      close();
      v.focus();
    };
    box.append(
      el('h2', { textContent: 'Границы и заливка' }),
      row('Применить к', scope),
      row('Быстрый выбор', preset('Все', { top: true, bottom: true, left: true, right: true, insideH: true, insideV: true }), preset('Внешние', { top: true, bottom: true, left: true, right: true }), preset('Внутренние', { insideH: true, insideV: true })),
      row('Стороны', el('label', {}, sides.top, ' верх '), el('label', {}, sides.bottom, ' низ '), el('label', {}, sides.left, ' лево '), el('label', {}, sides.right, ' право ')),
      row('Внутри', el('label', {}, sides.insideH, ' горизонтальные '), el('label', {}, sides.insideV, ' вертикальные ')),
      row('Линия', kind, width, color),
      row('Заливка', el('label', {}, fillOn, ' залить '), fill),
      el('p', { textContent: '«Убрать границы» стирает линии на отмеченных сторонах. Заливка без галочки не меняется.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { textContent: 'Убрать границы', onclick: () => apply(true) }), el('button', { class: 'primary', textContent: 'Применить', onclick: () => apply(false) })),
    );
  }, false);
}

export function clearFill(view: EditorView): void { setCellsAttr(view, 'cells', 'background', null); }
export const setValign = (view: EditorView, v: string): boolean => setCellsAttr(view, 'cells', 'valign', v === 'top' ? null : v);

// ---------------------------------------------------------------- столбцы и строки

export function equalColumns(view: EditorView): void {
  const t = tableAt(view.state);
  if (!t) { toast('Поставьте курсор в ячейку таблицы.'); return; }
  const cols = columnWidths(view, t.pos);
  const total = cols.reduce((a, b) => a + b, 0);
  setColumnWidths(view, t.pos, cols.map(() => total / cols.length));
}

/** Ширина таблицы по ширине текста (px): столбцы растягиваются пропорционально. */
export function fitTable(view: EditorView, textWidth: number): void {
  const t = tableAt(view.state);
  if (!t) { toast('Поставьте курсор в ячейку таблицы.'); return; }
  const cols = columnWidths(view, t.pos);
  const total = cols.reduce((a, b) => a + b, 0) || 1;
  setColumnWidths(view, t.pos, cols.map((w) => (w * textWidth) / total));
}

export function equalRows(view: EditorView): void {
  const t = tableAt(view.state);
  if (!t) { toast('Поставьте курсор в ячейку таблицы.'); return; }
  const tr = view.state.tr;
  let off = 0;
  let max = 0;
  const heights: number[] = [];
  t.node.forEach((row, o) => {
    const dom = view.nodeDOM(t.pos + 1 + o) as HTMLElement | null;
    const h = dom ? dom.offsetHeight : 0;
    heights.push(h);
    max = Math.max(max, h);
    off += row.nodeSize;
  });
  void off;
  const pt = Math.round((max * 0.75) * 10) / 10;
  t.node.forEach((row, o) => tr.setNodeMarkup(t.pos + 1 + o, undefined, { ...row.attrs, height: pt }));
  view.dispatch(tr);
}

/** Свойства таблицы: положение на странице, ширины столбцов, высота выделенных строк. */
export function tableDialog(view: () => EditorView, textWidth: () => number): void {
  const v0 = view();
  const t0 = tableAt(v0.state);
  if (!t0) { toast('Поставьте курсор в ячейку таблицы.'); return; }
  const CM = 96 / 2.54;
  const cols0 = columnWidths(v0, t0.pos);
  let rowsSel: number[] = [];
  try { const r = selectedRect(v0.state); for (let i = r.top; i < r.bottom; i++) rowsSel.push(i); } catch { rowsSel = []; }
  const rowH0 = rowsSel.length ? (t0.node.child(rowsSel[0]).attrs.height as number | null) : null;
  dialog((box, close) => {
    const row = (label: string, ...k: Node[]) => el('div', { class: 'dlg-row' }, el('label', { textContent: label }), ...k);
    const align = el('select', {}, new Option('слева', ''), new Option('по центру', 'center'), new Option('справа', 'right'));
    align.value = (t0.node.attrs.align as string | null) ?? '';
    const colInputs = cols0.map((w) => el('input', { type: 'text', value: String(Math.round((w / CM) * 100) / 100), size: 5 }));
    const rowH = el('input', { type: 'text', value: rowH0 ? String(Math.round((rowH0 / 28.3465) * 100) / 100) : '', size: 6, placeholder: 'авто' });
    const total = el('div', { class: 'tpl-hint' });
    const upd = () => { const s = colInputs.reduce((a, i) => a + (parseFloat(i.value.replace(',', '.')) || 0), 0); total.textContent = `Ширина таблицы: ${Math.round(s * 100) / 100} см (текст: ${Math.round((textWidth() / CM) * 100) / 100} см)`; };
    for (const i of colInputs) i.addEventListener('input', upd);
    upd();
    const ok = el('button', { class: 'primary', textContent: 'Применить', onclick: () => {
      const v = view();
      const t = tableAt(v.state);
      if (!t) { close(); return; }
      const widths = colInputs.map((i, k) => { const x = parseFloat(i.value.replace(',', '.')); return x > 0.2 ? x * CM : cols0[k]; });
      const tr = v.state.tr;
      tr.setNodeMarkup(t.pos, undefined, { ...t.node.attrs, align: align.value || null });
      const hv = parseFloat(rowH.value.replace(',', '.'));
      for (const r of rowsSel) {
        let off = 0;
        for (let i = 0; i < r; i++) off += t.node.child(i).nodeSize;
        tr.setNodeMarkup(t.pos + 1 + off, undefined, { ...t.node.child(r).attrs, height: hv > 0 ? Math.round(hv * 28.3465 * 10) / 10 : null });
      }
      v.dispatch(tr);
      setColumnWidths(v, t.pos, widths);
      close();
      v.focus();
    } });
    box.append(
      el('h2', { textContent: 'Свойства таблицы' }),
      row('Положение на странице', align),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Ширина столбцов, см' }), el('span', { class: 'tbl-cols' }, ...colInputs.flatMap((i, k) => [`${k + 1}: `, i, ' ']))),
      total,
      row('Высота выделенных строк, см', rowH),
      el('p', { textContent: 'Столбцы можно тянуть и мышью — за границу в таблице или за метку на верхней линейке. Строки — за метку на левой линейке.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), ok),
    );
  }, false);
}

