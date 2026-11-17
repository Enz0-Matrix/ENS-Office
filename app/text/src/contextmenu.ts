// SPDX-License-Identifier: AGPL-3.0-or-later
// Контекстные меню ENS-Text (правый щелчок, Shift+F10, клавиша «Меню»), как в Word/Р7: набор пунктов зависит от того,
// где щёлкнули — текст, выделение, ссылка, рисунок или надпись, таблица. Меню WebView (в нём «Обновить», которое
// перезапускает программу и уносит несохранённое) в документе не показывается.
// Орфография: подсказки проверки правописания знает только сам браузер, странице он их не отдаёт. Поэтому
// Shift+правый щелчок по тексту открывает системное меню с вариантами исправления — об этом сказано в нашем меню.

import { NodeSelection, TextSelection, type EditorState } from 'prosemirror-state';
import { CellSelection } from 'prosemirror-tables';
import type { EditorView } from 'prosemirror-view';
import { command, keyLabel, run } from '../../table/src/commands';
import { schema } from './schema';

export interface CtxHost {
  view: () => EditorView;
  /** Лист документа: правый щелчок внутри него открывает наше меню. */
  sheet: HTMLElement;
  /** Открыть адрес ссылки (в оболочке — системной программой). */
  openLink: (href: string) => void;
  /** Положить текст в буфер обмена. */
  copyText: (text: string) => void;
}

export type CtxKind = 'text' | 'selection' | 'link' | 'image' | 'table';

interface Item { label: string; run: () => void; keys?: string; disabled?: boolean; hint?: boolean }
type Entry = Item | '-';

let H: CtxHost;
let menu: HTMLElement;
let lastKinds: CtxKind[] = [];

/** Ссылка под курсором: адрес и границы (весь фрагмент с той же ссылкой). */
function linkRange(state: EditorState): { href: string; from: number; to: number } | null {
  const $p = state.selection.$from;
  const type = schema.marks.link;
  const mark = type.isInSet($p.marks()) ?? ($p.nodeAfter ? type.isInSet($p.nodeAfter.marks) : null);
  if (!mark) return null;
  const parent = $p.parent, start = $p.start();
  let from = -1, to = -1, pos = start;
  // Непрерывный кусок абзаца с этой ссылкой, содержащий курсор.
  for (let i = 0; i < parent.childCount; i++) {
    const ch = parent.child(i);
    const end = pos + ch.nodeSize;
    if (mark.isInSet(ch.marks)) {
      if (from < 0) from = pos;
      to = end;
    } else if (to >= 0 && to < $p.pos) {
      from = -1; to = -1;
    } else if (to >= $p.pos) break;
    pos = end;
  }
  return from >= 0 && from <= $p.pos && to >= $p.pos ? { href: mark.attrs.href as string, from, to } : null;
}

/** Что под курсором (в порядке приоритета для пунктов меню). */
export function contextKinds(state: EditorState): CtxKind[] {
  const sel = state.selection;
  const kinds: CtxKind[] = [];
  if (sel instanceof NodeSelection && (sel.node.type === schema.nodes.image || sel.node.type === schema.nodes.frame)) return ['image'];
  if (linkRange(state)) kinds.push('link');
  const $f = sel.$from;
  let inTable = sel instanceof CellSelection;
  for (let d = $f.depth; d > 0 && !inTable; d--) if ($f.node(d).type === schema.nodes.table) inTable = true;
  if (inTable) kinds.push('table');
  kinds.push(sel.empty ? 'text' : 'selection');
  return kinds;
}

/** Пункт по команде из общего списка команд (подпись — наша, сочетание клавиш — из команды). */
function cmd(id: string, label: string, keys?: string, arg?: string): Item {
  const c = command(id);
  return { label, keys: keys ?? (c?.keys?.length ? keyLabel(c.keys[0]) : undefined), run: () => run(id, arg), disabled: !c };
}

function items(kinds: CtxKind[]): Entry[] {
  const v = H.view();
  const out: Entry[] = [];
  const empty = v.state.selection.empty;
  const clip: Entry[] = [
    { ...cmd('clipboard.cut', 'Вырезать', 'Ctrl+X'), disabled: empty },
    { ...cmd('clipboard.copy', 'Копировать', 'Ctrl+C'), disabled: empty },
    cmd('clipboard.paste', 'Вставить', 'Ctrl+V'),
  ];
  if (kinds[0] === 'image') {
    const frame = (v.state.selection as NodeSelection).node.type === schema.nodes.frame;
    out.push(...clip, '-');
    if (!frame) {
      out.push(cmd('image.format', 'Формат рисунка…'), cmd('image.crop', 'Обрезка'), cmd('image.reset', 'Исходный размер'), '-');
      out.push(cmd('image.wrap', 'В тексте', undefined, 'inline'), cmd('image.wrap', 'За текстом', undefined, 'behind'), cmd('image.wrap', 'Перед текстом', undefined, 'front'), '-');
      out.push(cmd('image.caption', 'Вставить название…'), cmd('image.replace', 'Заменить рисунок…'), cmd('image.delete', 'Удалить рисунок'));
    } else {
      out.push(cmd('image.front', 'На передний план'), cmd('image.back', 'На задний план'));
    }
    return out;
  }
  out.push(...clip);
  if (kinds.includes('link')) {
    const l = linkRange(v.state)!;
    out.push('-',
      { label: 'Открыть ссылку', keys: 'Ctrl+щелчок', run: () => H.openLink(l.href) },
      { label: 'Изменить ссылку…', run: () => { v.dispatch(v.state.tr.setSelection(TextSelection.create(v.state.doc, l.from, l.to))); run('insert.link'); } },
      { label: 'Копировать адрес ссылки', run: () => H.copyText(l.href) },
      { label: 'Удалить ссылку', run: () => { v.dispatch(v.state.tr.removeMark(l.from, l.to, schema.marks.link)); v.focus(); } });
  }
  if (kinds.includes('table')) {
    out.push('-',
      cmd('table.rowBefore', 'Вставить строку выше'), cmd('table.rowAfter', 'Вставить строку ниже'),
      cmd('table.colBefore', 'Вставить столбец слева'), cmd('table.colAfter', 'Вставить столбец справа'), '-',
      cmd('table.delRow', 'Удалить строку'), cmd('table.delCol', 'Удалить столбец'), cmd('table.delete', 'Удалить таблицу'), '-',
      { ...cmd('table.merge', 'Объединить ячейки'), disabled: !(v.state.selection instanceof CellSelection) }, cmd('table.split', 'Разделить ячейку'), '-',
      cmd('table.borderDialog', 'Границы и заливка…'), cmd('table.props', 'Свойства таблицы…'));
    return out;
  }
  if (kinds.includes('selection')) {
    out.push('-', cmd('format.bold', 'Полужирный'), cmd('format.italic', 'Курсив'), cmd('format.underline', 'Подчёркнутый'), cmd('format.clear', 'Очистить форматирование'),
      '-', cmd('typo.apply', 'Типограф для выделенного'));
  }
  if (!kinds.includes('link')) out.push('-', cmd('insert.link', 'Ссылка…'));
  out.push(cmd('list.bullet', 'Маркированный список'), cmd('list.ordered', 'Нумерованный список'), '-', cmd('review.inspector', 'Рентген форматирования'));
  out.push('-', { label: 'Орфография: Shift + правый щелчок по слову', run: () => undefined, disabled: true, hint: true });
  return out;
}

export function hideContextMenu(): void {
  if (menu) menu.hidden = true;
}

/** Показать меню для текущего выделения в точке экрана (x, y). */
export function showContextMenu(x: number, y: number): void {
  const kinds = contextKinds(H.view().state);
  lastKinds = kinds;
  menu.replaceChildren();
  let prevSep = true;
  for (const e of items(kinds)) {
    if (e === '-') {
      if (!prevSep) menu.append(Object.assign(document.createElement('hr'), { className: 'ctx-sep' }));
      prevSep = true;
      continue;
    }
    const b = document.createElement('button');
    b.type = 'button';
    b.disabled = !!e.disabled;
    if (e.hint) b.className = 'ctx-hint';
    b.setAttribute('role', 'menuitem');
    const span = document.createElement('span');
    span.textContent = e.label;
    b.append(span);
    if (e.keys) { const k = document.createElement('small'); k.textContent = e.keys; b.append(k); }
    b.addEventListener('click', () => { hideContextMenu(); H.view().focus(); e.run(); });
    menu.append(b);
    prevSep = false;
  }
  if (menu.lastElementChild?.tagName === 'HR') menu.lastElementChild.remove();
  menu.dataset.kind = kinds.join(' ');
  menu.hidden = false;
  const { width, height } = menu.getBoundingClientRect();
  menu.style.left = `${Math.max(4, Math.min(x, window.innerWidth - width - 4))}px`;
  menu.style.top = `${Math.max(4, y + height > window.innerHeight - 4 ? y - height : y)}px`;
  (menu.querySelector('button:not(:disabled)') as HTMLElement | null)?.focus({ preventScroll: true });
}

/** Правый щелчок: курсор встаёт в точку щелчка, если она вне выделения (как в Word); по рисунку — рисунок выделяется. */
function placeSelection(e: MouseEvent): void {
  const v = H.view();
  const t = e.target as HTMLElement;
  if (t.closest('.img-ov')) return; // рамка выделенного рисунка: работаем с ним
  const img = t.closest('img, .frame-box');
  if (img && v.dom.contains(img)) {
    let pos = v.posAtDOM(img, 0);
    if (img.classList.contains('frame-box')) {
      const $p = v.state.doc.resolve(pos);
      let d = $p.depth;
      while (d > 0 && $p.node(d).type !== schema.nodes.frame) d--;
      pos = d > 0 ? $p.before(d) : -1;
    }
    const n = pos >= 0 ? v.state.doc.nodeAt(pos) : null;
    if (n && (n.type === schema.nodes.image || n.type === schema.nodes.frame)) { v.dispatch(v.state.tr.setSelection(NodeSelection.create(v.state.doc, pos))); return; }
  }
  const sel = v.state.selection;
  if (sel instanceof CellSelection && t.closest('.selectedCell')) return;
  const at = v.posAtCoords({ left: e.clientX, top: e.clientY });
  if (!at) return;
  if (!sel.empty && !(sel instanceof NodeSelection) && at.pos >= sel.from && at.pos <= sel.to) return;
  v.dispatch(v.state.tr.setSelection(TextSelection.near(v.state.doc.resolve(at.pos))));
}

export function initContextMenu(host: CtxHost): void {
  H = host;
  menu = document.createElement('div');
  menu.className = 'context-menu ctx-text';
  menu.setAttribute('role', 'menu');
  menu.hidden = true;
  document.body.append(menu);

  document.addEventListener('contextmenu', (e) => {
    const t = e.target as HTMLElement | null;
    if (!t || t.closest('input, textarea, select')) return; // в полях ввода — обычное меню с «Вставить»
    if (menu.contains(t)) { e.preventDefault(); return; }
    const inHf = !!t.closest('.hf-edit');
    const inDoc = H.sheet.contains(t) && !inHf;
    // Shift+правый щелчок по тексту (и в колонтитуле) — системное меню: в нём варианты исправления орфографии.
    if (e.shiftKey && (inHf || (inDoc && H.view().dom.contains(t)))) { hideContextMenu(); return; }
    e.preventDefault();
    if (!inDoc || document.querySelector('.modal')) { hideContextMenu(); return; }
    placeSelection(e);
    H.view().focus();
    showContextMenu(e.clientX, e.clientY);
  });
  // Клавиатура: Shift+F10 и клавиша «Меню» — меню у курсора.
  H.sheet.addEventListener('keydown', (e) => {
    if (!(e.key === 'ContextMenu' || (e.key === 'F10' && e.shiftKey))) return;
    e.preventDefault();
    const v = H.view();
    const c = v.coordsAtPos(v.state.selection.head);
    showContextMenu(c.left, c.bottom);
  });
  menu.addEventListener('keydown', (e) => {
    const btns = [...menu.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
    const i = btns.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      btns[(i + (e.key === 'ArrowDown' ? 1 : -1) + btns.length) % btns.length]?.focus();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      hideContextMenu();
      H.view().focus();
    }
  });
  document.addEventListener('mousedown', (e) => { if (!menu.hidden && !menu.contains(e.target as Node)) hideContextMenu(); }, true);
  window.addEventListener('blur', hideContextMenu);
  window.addEventListener('resize', hideContextMenu);
  H.sheet.parentElement?.addEventListener('scroll', hideContextMenu);
}

/** Для тестов: что показало последнее меню. */
export const lastMenuKinds = (): CtxKind[] => lastKinds;
