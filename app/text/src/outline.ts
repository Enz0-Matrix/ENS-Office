// SPDX-License-Identifier: AGPL-3.0-or-later
// Панель «Структура» (Вид → Структура, как «Область навигации» в Word): заголовки деревом, переход щелчком,
// перенос главы целиком (перетаскиванием или кнопками ↑ ↓) вместе со всем текстом до следующего заголовка
// того же или более высокого уровня. Лучше, чем в Word: панель сразу показывает ошибки структуры —
// пропущенный уровень заголовка, пустой заголовок, раздел без текста.

import type { Node as PMNode } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { el } from './ui';

export interface Section {
  /** Номер верхнего блока заголовка и конец главы (не включая) — индексы детей документа. */
  index: number; end: number;
  pos: number; level: number; text: string;
  problems: string[];
}

/** Главы по заголовкам верхнего уровня документа. */
export function sections(doc: PMNode): Section[] {
  const tops: { node: PMNode; pos: number }[] = [];
  doc.forEach((node, pos) => tops.push({ node, pos }));
  const heads = tops.map((t, i) => ({ ...t, i })).filter((t) => t.node.type === schema.nodes.heading);
  return heads.map((h, k) => {
    const level = h.node.attrs.level as number;
    const next = heads.slice(k + 1).find((x) => (x.node.attrs.level as number) <= level);
    const end = next ? next.i : tops.length;
    const text = h.node.textContent.replace(/\s+/g, ' ').trim();
    const problems: string[] = [];
    const prevLevel = k ? heads[k - 1].node.attrs.level as number : 0;
    if (level > prevLevel + 1) problems.push(`пропущен уровень: после ${prevLevel ? `«Заголовка ${prevLevel}»` : 'начала документа'} идёт «Заголовок ${level}»`);
    if (!text) problems.push('пустой заголовок');
    // Текст раздела — блоки между заголовком и следующим заголовком (любого уровня).
    const nextAny = heads[k + 1]?.i ?? tops.length;
    const body = tops.slice(h.i + 1, nextAny).some((t) => t.node.textContent.trim() || t.node.type !== schema.nodes.paragraph);
    const hasSub = nextAny < end;
    if (!body && !hasSub) problems.push('раздел без текста');
    return { index: h.i, end, pos: h.pos, level, text, problems };
  });
}

/** Перенести главу (заголовок со всем содержимым) так, чтобы она встала перед блоком with index `before`. */
export function moveSection(state: EditorState, s: Section, before: number): Transaction | null {
  if (before >= s.index && before <= s.end) return null; // внутрь самой себя или на то же место
  const offs: number[] = [];
  state.doc.forEach((_n, off) => offs.push(off));
  const at = (i: number) => (i >= offs.length ? state.doc.content.size : offs[i]);
  const from = at(s.index), to = at(s.end);
  const slice = state.doc.slice(from, to);
  const tr = state.tr;
  const target = at(before);
  if (target > to) { tr.insert(target, slice.content); tr.delete(from, to); }
  else { tr.delete(from, to); tr.insert(target, slice.content); }
  const newPos = target > to ? target - (to - from) : target;
  return tr.setSelection(TextSelection.create(tr.doc, Math.min(newPos + 1, tr.doc.content.size))).scrollIntoView();
}

/** Соседняя глава того же уровня (для кнопок ↑ ↓). */
function sibling(all: Section[], s: Section, dir: 1 | -1): Section | null {
  const i = all.indexOf(s);
  if (dir < 0) {
    for (let k = i - 1; k >= 0; k--) { if (all[k].level < s.level) return null; if (all[k].level === s.level) return all[k]; }
    return null;
  }
  for (let k = i + 1; k < all.length; k++) { if (all[k].level < s.level) return null; if (all[k].level === s.level && all[k].index >= s.end) return all[k]; }
  return null;
}

let pane: HTMLElement | null = null;
let paneView: EditorView | null = null;
let sig = '';
let dragged: Section | null = null;

export const outlineOpen = () => !!pane;

export function toggleOutline(view: EditorView): void {
  if (pane) { pane.remove(); pane = null; document.body.classList.remove('with-outline'); return; }
  paneView = view;
  pane = el('aside', { class: 'outline-pane', ariaLabel: 'Структура документа' });
  document.body.append(pane);
  document.body.classList.add('with-outline');
  sig = '';
  renderOutline();
}

export function renderOutline(): void {
  const view = paneView;
  if (!pane || !view) return;
  const all = sections(view.state.doc);
  const at = view.state.selection.from;
  const cur = [...all].reverse().find((s) => s.pos <= at);
  const nextSig = JSON.stringify([all.map((s) => [s.index, s.level, s.text, s.problems]), cur?.index]);
  if (nextSig === sig) return;
  sig = nextSig;
  const go = (s: Section) => {
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, Math.min(s.pos + 1, view.state.doc.content.size))).scrollIntoView());
    view.focus();
  };
  const move = (s: Section, before: number) => { const tr = moveSection(view.state, s, before); if (tr) { view.dispatch(tr); view.focus(); } };
  const problems = all.reduce((n, s) => n + s.problems.length, 0);
  const rows = all.map((s) => {
    const up = sibling(all, s, -1), down = sibling(all, s, 1);
    const row = el('div', { class: `ol-row${s === cur ? ' current' : ''}${s.problems.length ? ' bad' : ''}`, draggable: true, title: s.problems.join('; ') || 'Щелчок — перейти; перетащите, чтобы перенести главу целиком' });
    row.dataset.level = String(s.level);
    row.dataset.index = String(s.index);
    row.append(
      el('span', { class: 'ol-text', textContent: s.text || '(пустой заголовок)' }),
      ...(s.problems.length ? [el('span', { class: 'ol-warn', textContent: '⚠', ariaLabel: s.problems.join('; ') })] : []),
      el('button', { class: 'ol-move', textContent: '↑', title: 'Переместить главу выше', disabled: !up, onclick: (e: Event) => { e.stopPropagation(); if (up) move(s, up.index); } }),
      el('button', { class: 'ol-move', textContent: '↓', title: 'Переместить главу ниже', disabled: !down, onclick: (e: Event) => { e.stopPropagation(); if (down) move(s, down.end); } }),
    );
    row.addEventListener('click', () => go(s));
    row.addEventListener('dragstart', (e) => { dragged = s; e.dataTransfer?.setData('text/plain', s.text); row.classList.add('dragging'); });
    row.addEventListener('dragend', () => { dragged = null; row.classList.remove('dragging'); });
    row.addEventListener('dragover', (e) => { if (dragged && dragged !== s) { e.preventDefault(); row.classList.add('drop'); } });
    row.addEventListener('dragleave', () => row.classList.remove('drop'));
    row.addEventListener('drop', (e) => { e.preventDefault(); row.classList.remove('drop'); if (dragged) move(dragged, s.index); dragged = null; });
    return row;
  });
  pane.replaceChildren(
    el('div', { class: 'ol-head' }, el('b', { textContent: 'Структура' }), el('button', { textContent: '✕', title: 'Закрыть', onclick: () => toggleOutline(view) })),
    ...(problems ? [el('p', { class: 'ol-summary', textContent: `Замечаний к структуре: ${problems}. Наведите на ⚠, чтобы узнать подробности.` })] : []),
    ...(rows.length ? rows : [el('p', { class: 'ol-empty', textContent: 'Заголовков нет. Оформите названия разделов стилями «Заголовок 1–6» — они появятся здесь.' })]),
  );
}
