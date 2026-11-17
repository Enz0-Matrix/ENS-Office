// SPDX-License-Identifier: AGPL-3.0-or-later
// Перемещение таблицы мышью за ручку в левом верхнем углу (как в Word): таблица переносится между абзацами
// одной транзакцией — одна запись в истории отмены. Пока тянем, линия показывает, куда встанет таблица.
// Работает для таблиц верхнего уровня (не внутри надписи, списка или другой таблицы); обтекание — позже.

import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';

export interface TableMoveHost {
  view: () => EditorView;
  sheet: HTMLElement;
  scale: () => number;
  locked: () => boolean;
}

let H: TableMoveHost;
let handle: HTMLElement;
let line: HTMLElement;
/** Таблица, к которой сейчас приставлена ручка (позиция узла table в документе). */
let current: number | null = null;
let hovered: number | null = null;
let dragging = false;

/** Позиция таблицы верхнего уровня по её DOM (или null). */
function tablePosOf(tableEl: Element): number | null {
  const v = H.view();
  if (!v.dom.contains(tableEl)) return null;
  let pos: number;
  try { pos = v.posAtDOM(tableEl, 0); } catch { return null; }
  const $p = v.state.doc.resolve(pos);
  for (let d = $p.depth; d > 0; d--) {
    if ($p.node(d).type === schema.nodes.table) return d === 1 ? $p.before(d) : null;
  }
  return null;
}

/** Таблица верхнего уровня, в которой стоит курсор. */
function tableAtCursor(): number | null {
  const $f = H.view().state.selection.$from;
  return $f.depth >= 1 && $f.node(1).type === schema.nodes.table ? $f.before(1) : null;
}

function tableDom(pos: number): HTMLElement | null {
  const d = H.view().nodeDOM(pos);
  if (!(d instanceof HTMLElement)) return null;
  return d.matches('table') ? d : d.querySelector('table');
}

function place(): void {
  const pos = hovered ?? tableAtCursor();
  const t = pos != null && !H.locked() ? tableDom(pos) : null;
  if (!t || pos == null) { handle.hidden = true; current = null; return; }
  current = pos;
  const k = H.scale();
  const s = H.sheet.getBoundingClientRect();
  const r = t.getBoundingClientRect();
  handle.hidden = false;
  handle.style.left = `${(r.left - s.left) / k - 18}px`;
  handle.style.top = `${(r.top - s.top) / k - 18}px`;
}

/** Пересчитать ручку после изменения документа, выделения, масштаба. */
export function updateTableMove(): void {
  if (!H || dragging) return;
  if (hovered != null && H.view().state.doc.nodeAt(hovered)?.type !== schema.nodes.table) hovered = null;
  place();
}

/** Куда встанет таблица: позиция между блоками верхнего уровня и линия на экране (или null — на месте). */
function dropTarget(x: number, y: number, from: number): { pos: number; top: number; left: number; width: number } | null {
  const v = H.view();
  const doc = v.state.doc;
  const tbl = doc.nodeAt(from);
  if (!tbl) return null;
  const at = v.posAtCoords({ left: x, top: y });
  if (!at) return null;
  const $p = doc.resolve(at.inside >= 0 ? at.inside : at.pos);
  const idx = $p.depth >= 1 ? $p.index(0) : Math.min($p.index(0), doc.childCount - 1);
  const start = (() => { let o = 0; for (let i = 0; i < idx; i++) o += doc.child(i).nodeSize; return o; })();
  const block = doc.child(idx);
  const dom = v.nodeDOM(start);
  if (!(dom instanceof HTMLElement)) return null;
  const r = dom.getBoundingClientRect();
  const before = y < r.top + r.height / 2;
  const pos = before ? start : start + block.nodeSize;
  // На прежнем месте (сразу до или после самой таблицы) — не перемещаем.
  if (pos === from || pos === from + tbl.nodeSize) return null;
  const k = H.scale();
  const s = H.sheet.getBoundingClientRect();
  const vr = v.dom.getBoundingClientRect();
  const cs = getComputedStyle(v.dom);
  const left = vr.left + parseFloat(cs.paddingLeft) * k;
  const width = vr.width - (parseFloat(cs.paddingLeft) + parseFloat(cs.paddingRight)) * k;
  return { pos, top: ((before ? r.top : r.bottom) - s.top) / k, left: (left - s.left) / k, width: width / k };
}

function startDrag(e: MouseEvent): void {
  if (e.button !== 0 || current == null || H.locked()) return;
  e.preventDefault();
  e.stopPropagation();
  const from = current;
  dragging = true;
  H.sheet.classList.add('tbl-moving');
  let target: ReturnType<typeof dropTarget> = null;
  const mv = (ev: MouseEvent) => {
    target = dropTarget(ev.clientX, ev.clientY, from);
    line.hidden = !target;
    if (target) Object.assign(line.style, { top: `${target.top}px`, left: `${target.left}px`, width: `${target.width}px` });
  };
  const finish = (apply: boolean) => {
    document.removeEventListener('mousemove', mv);
    document.removeEventListener('mouseup', up);
    document.removeEventListener('keydown', key, true);
    dragging = false;
    H.sheet.classList.remove('tbl-moving');
    line.hidden = true;
    if (apply && target) moveTable(from, target.pos);
    place();
  };
  const up = () => finish(true);
  const key = (ev: KeyboardEvent) => { if (ev.key === 'Escape') { ev.preventDefault(); ev.stopPropagation(); finish(false); } };
  document.addEventListener('mousemove', mv);
  document.addEventListener('mouseup', up);
  document.addEventListener('keydown', key, true);
}

/** Перенести таблицу верхнего уровня с позиции from на позицию to (между блоками). Одна транзакция. */
export function moveTable(from: number, to: number): boolean {
  const v = H.view();
  const tbl = v.state.doc.nodeAt(from);
  if (!tbl || tbl.type !== schema.nodes.table) return false;
  let tr = v.state.tr.delete(from, from + tbl.nodeSize);
  // Документ не может остаться пустым: если таблица была единственным блоком, на её месте остаётся абзац.
  if (!tr.doc.childCount) tr = tr.insert(0, schema.nodes.paragraph.create());
  const at = tr.mapping.map(to, to > from ? -1 : 1);
  tr = tr.insert(at, tbl);
  tr = tr.setSelection(TextSelection.near(tr.doc.resolve(at + 1)));
  v.dispatch(tr.scrollIntoView());
  v.focus();
  return true;
}

export function initTableMove(host: TableMoveHost): void {
  H = host;
  handle = document.createElement('div');
  handle.className = 'tbl-move';
  handle.hidden = true;
  handle.title = 'Переместить таблицу: перетащите (Esc — отмена)';
  handle.setAttribute('aria-label', 'Переместить таблицу');
  handle.textContent = '✥';
  handle.addEventListener('mousedown', startDrag);
  line = document.createElement('div');
  line.className = 'tbl-drop';
  line.hidden = true;
  H.sheet.append(handle, line);
  H.sheet.addEventListener('mousemove', (e) => {
    if (dragging) return;
    const t = e.target as HTMLElement | null;
    if (t === handle) return;
    const tableEl = t?.closest('table');
    // Между таблицей и ручкой есть зазор: рядом с таблицей ручку не прячем, чтобы до неё можно было дотянуться.
    if (!tableEl && hovered != null) {
      const r = tableDom(hovered)?.getBoundingClientRect();
      if (r && e.clientX >= r.left - 28 && e.clientX <= r.right && e.clientY >= r.top - 28 && e.clientY <= r.bottom) return;
    }
    const pos = tableEl ? tablePosOf(tableEl) : null;
    if (pos === hovered) return;
    hovered = pos;
    place();
  });
  H.sheet.addEventListener('mouseleave', () => { if (!dragging && hovered != null) { hovered = null; place(); } });
  H.sheet.parentElement?.addEventListener('scroll', () => place());
  window.addEventListener('resize', () => place());
}
