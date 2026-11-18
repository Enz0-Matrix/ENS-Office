// SPDX-License-Identifier: AGPL-3.0-or-later
// Примечания (комментарии рецензентов): Проверка → Примечания, Ctrl+Alt+M — создать, как в Word.
//
// Ветка примечания (текст, ответы, «решено») хранится в метке `comment` на тексте, поэтому отмена, копирование
// и вырезание работают с ней как с текстом. Панель справа показывает все ветки по порядку; щелчок по карточке
// выделяет текст, щелчок по тексту с примечанием подсвечивает карточку. В docx — родные примечания Word
// (comments.xml, ответы и «решено» — commentsExtended.xml).

import type { Mark, Node as PMNode } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { el } from './ui';

export interface Reply { author: string; date: string; text: string }
export interface Thread { id: number; author: string; date: string; text: string; replies: Reply[]; done: boolean }
export interface ThreadInfo extends Thread { from: number; to: number; quote: string }

const nowIso = () => new Date().toISOString().replace(/\.\d+Z$/, 'Z');
export const fmtDate = (iso: string) => {
  const d = iso ? new Date(iso) : null;
  return d && !isNaN(+d) ? d.toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '';
};

const attrsOf = (m: Mark): Thread => m.attrs as Thread;

/** Все ветки по порядку в тексте: где лежат и какой текст отмечен. */
export function listThreads(doc: PMNode): ThreadInfo[] {
  const map = new Map<number, ThreadInfo>();
  doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    for (const m of node.marks) {
      if (m.type !== schema.marks.comment) continue;
      const a = attrsOf(m);
      const t = map.get(a.id);
      const piece = node.isText ? node.text ?? '' : '';
      if (t) { t.to = pos + node.nodeSize; t.quote += piece; } else map.set(a.id, { ...a, replies: [...(a.replies ?? [])], from: pos, to: pos + node.nodeSize, quote: piece });
    }
    return false;
  });
  return [...map.values()].sort((a, b) => a.from - b.from);
}

const nextId = (doc: PMNode) => listThreads(doc).reduce((m, t) => Math.max(m, t.id), 0) + 1;

/** Создать примечание на выделении; без выделения — на слове у курсора (как Word). */
export function addComment(state: EditorState, author: string, text = ''): Transaction | null {
  let { from, to } = state.selection;
  if (from === to) {
    const $p = state.doc.resolve(from);
    const parent = $p.parent, off = $p.parentOffset;
    const s = parent.textBetween(0, parent.content.size, '\u0000', '\u0000');
    let a = off, b = off;
    while (a > 0 && /[\p{L}\p{N}_-]/u.test(s[a - 1])) a--;
    while (b < s.length && /[\p{L}\p{N}_-]/u.test(s[b])) b++;
    if (a === b) return null;
    from = $p.start() + a; to = $p.start() + b;
  }
  const mark = schema.marks.comment.create({ id: nextId(state.doc), author, date: nowIso(), text, replies: [], done: false });
  const tr = state.tr.addMark(from, to, mark);
  return tr.setSelection(TextSelection.create(tr.doc, to));
}

/** Заменить данные ветки на всех её кусках. */
export function updateThread(state: EditorState, id: number, change: (t: Thread) => Thread): Transaction | null {
  const t = listThreads(state.doc).find((x) => x.id === id);
  if (!t) return null;
  const tr = state.tr;
  const next = change({ id: t.id, author: t.author, date: t.date, text: t.text, replies: t.replies, done: t.done });
  state.doc.nodesBetween(t.from, t.to, (node, pos) => {
    if (!node.isInline) return true;
    const old = node.marks.find((m) => m.type === schema.marks.comment && attrsOf(m).id === id);
    if (!old) return false;
    const from = Math.max(pos, t.from), to = Math.min(pos + node.nodeSize, t.to);
    tr.removeMark(from, to, old);
    tr.addMark(from, to, schema.marks.comment.create(next));
    return false;
  });
  return tr;
}

export function deleteThread(state: EditorState, id: number | 'all'): Transaction | null {
  const tr = state.tr;
  let any = false;
  state.doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    for (const m of node.marks) if (m.type === schema.marks.comment && (id === 'all' || attrsOf(m).id === id)) { tr.removeMark(pos, pos + node.nodeSize, m); any = true; }
    return false;
  });
  return any ? tr : null;
}

/** Ветка под курсором (если их несколько — самая короткая, «внутренняя»). */
export function threadAt(state: EditorState): ThreadInfo | null {
  const { from } = state.selection;
  const all = listThreads(state.doc).filter((t) => t.from <= from && t.to >= from);
  return all.sort((a, b) => (a.to - a.from) - (b.to - b.from))[0] ?? null;
}

// --- Панель примечаний

let pane: HTMLElement | null = null;
let paneView: EditorView | null = null;
let me: () => string = () => 'Пользователь';
let lastSig = '';
let focusId: number | null = null;

export const panelOpen = () => !!pane;

export function openComments(view: EditorView, user: () => string, focus?: number): void {
  me = user;
  paneView = view;
  if (!pane) {
    pane = el('aside', { class: 'comments-pane', ariaLabel: 'Примечания' });
    document.body.append(pane);
    document.body.classList.add('with-comments');
  }
  focusId = focus ?? null;
  lastSig = '';
  renderComments();
}

export function closeComments(): void {
  pane?.remove();
  pane = null;
  document.body.classList.remove('with-comments');
}

function select(view: EditorView, t: ThreadInfo): void {
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, t.from, t.to)).scrollIntoView());
}

/** Перерисовать панель (после правок). Без изменений веток и текущей ветки DOM не трогается. */
export function renderComments(): void {
  const view = paneView;
  if (!pane || !view) return;
  const threads = listThreads(view.state.doc);
  const cur = threadAt(view.state)?.id ?? null;
  const sig = JSON.stringify([threads.map((t) => [t.id, t.text, t.replies, t.done, t.quote]), cur, focusId]);
  // Не перерисовываем, пока в карточке набирают текст: иначе пропадёт ввод.
  if (sig === lastSig || pane.contains(document.activeElement) && document.activeElement?.tagName === 'TEXTAREA') return;
  lastSig = sig;
  document.querySelectorAll('.ProseMirror span.cmt.cmt-active').forEach((e) => e.classList.remove('cmt-active'));
  if (cur != null) document.querySelectorAll(`.ProseMirror span.cmt[data-cmt-id="${cur}"]`).forEach((e) => e.classList.add('cmt-active'));
  const apply = (tr: Transaction | null) => { if (tr) view.dispatch(tr); };
  const cards = threads.map((t) => {
    const editor = (initial: string, placeholder: string, save: (text: string) => void, cancel?: () => void) => {
      const area = el('textarea', { value: initial, placeholder, rows: 2 });
      area.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); area.blur(); save(area.value.trim()); }
        if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); area.blur(); cancel?.(); }
      });
      return area;
    };
    const card = el('section', { class: `cmt-card${t.done ? ' done' : ''}${t.id === cur ? ' active' : ''}` });
    card.dataset.id = String(t.id);
    const head = el('div', { class: 'cmt-head' }, el('b', { textContent: t.author || 'Без имени' }), el('span', { class: 'cmt-date', textContent: fmtDate(t.date) }));
    const quote = el('div', { class: 'cmt-quote', textContent: t.quote.length > 80 ? t.quote.slice(0, 80) + '…' : t.quote });
    card.append(head, quote);
    if (!t.text || focusId === t.id && !t.text) {
      const area = editor(t.text, 'Текст примечания (Ctrl+Enter — сохранить)', (text) => { focusId = null; apply(updateThread(view.state, t.id, (x) => ({ ...x, text }))); });
      card.append(area, el('div', { class: 'cmt-actions' },
        el('button', { class: 'primary', textContent: 'Сохранить', onclick: () => { focusId = null; apply(updateThread(view.state, t.id, (x) => ({ ...x, text: area.value.trim() }))); } }),
        el('button', { textContent: 'Удалить', onclick: () => apply(deleteThread(view.state, t.id)) })));
      if (focusId === t.id) setTimeout(() => area.focus(), 0);
    } else {
      card.append(el('div', { class: 'cmt-text', textContent: t.text }));
      for (const r of t.replies) card.append(el('div', { class: 'cmt-reply' }, el('div', { class: 'cmt-head' }, el('b', { textContent: r.author }), el('span', { class: 'cmt-date', textContent: fmtDate(r.date) })), el('div', { class: 'cmt-text', textContent: r.text })));
      const reply = editor('', 'Ответить… (Ctrl+Enter)', (text) => {
        if (text) apply(updateThread(view.state, t.id, (x) => ({ ...x, replies: [...x.replies, { author: me(), date: nowIso(), text }] })));
      });
      card.append(reply, el('div', { class: 'cmt-actions' },
        el('button', { textContent: 'Ответить', onclick: () => { const text = reply.value.trim(); if (text) apply(updateThread(view.state, t.id, (x) => ({ ...x, replies: [...x.replies, { author: me(), date: nowIso(), text }] }))); } }),
        el('button', { textContent: t.done ? 'Открыть снова' : 'Решено', onclick: () => apply(updateThread(view.state, t.id, (x) => ({ ...x, done: !x.done }))) }),
        el('button', { textContent: 'Удалить', onclick: () => apply(deleteThread(view.state, t.id)) })));
    }
    card.addEventListener('mousedown', (e) => {
      if ((e.target as HTMLElement).closest('textarea,button')) return;
      e.preventDefault();
      select(view, t);
    });
    return card;
  });
  const open = threads.filter((t) => !t.done).length;
  pane.replaceChildren(
    el('div', { class: 'cmt-pane-head' }, el('b', { textContent: `Примечания (${threads.length}${threads.length !== open ? `, открытых ${open}` : ''})` }),
      el('button', { textContent: '✕', title: 'Закрыть панель', onclick: closeComments })),
    ...(cards.length ? cards : [el('p', { class: 'cmt-empty', textContent: 'Примечаний нет. Выделите текст и нажмите «Создать примечание» (Ctrl+Alt+M).' })]),
  );
  pane.querySelector('.cmt-card.active')?.scrollIntoView({ block: 'nearest' });
}

/** К следующей или предыдущей ветке. */
export function gotoThread(view: EditorView, dir: 1 | -1): boolean {
  const all = listThreads(view.state.doc);
  if (!all.length) return false;
  const at = view.state.selection.from;
  const t = dir > 0 ? all.find((x) => x.from > at) ?? all[0] : [...all].reverse().find((x) => x.to < at) ?? all[all.length - 1];
  select(view, t);
  return true;
}
