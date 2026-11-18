// SPDX-License-Identifier: AGPL-3.0-or-later
// Закладки: невидимые именованные точки в тексте (как «Вставка → Закладка» в Word) и ссылки на них «#имя».
// Лучше, чем в Word: при переименовании закладки ссылки на неё обновляются сами, а при удалении программа
// показывает, сколько ссылок станет «битыми», и список таких ссылок есть в окне.

import type { Node as PMNode } from 'prosemirror-model';
import { closeHistory } from 'prosemirror-history';
import { Plugin, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { schema } from './schema';

export interface BookmarkInfo { name: string; pos: number; context: string }

/** Имя как в Word: начинается с буквы, дальше буквы, цифры и «_», до 40 знаков. Скрытые (служебные) начинаются с «_». */
export const BOOKMARK_RE = /^[\p{L}][\p{L}\p{N}_]{0,39}$/u;
export const isHidden = (name: string) => name.startsWith('_');

export function listBookmarks(doc: PMNode): BookmarkInfo[] {
  const out: BookmarkInfo[] = [];
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.bookmark) return true;
    const $p = doc.resolve(pos);
    const block = $p.parent;
    const off = $p.parentOffset;
    const text = block.textBetween(0, block.content.size, ' ', ' ');
    // Короткий кусок текста после закладки — чтобы в списке было видно, где она стоит.
    const after = block.textBetween(Math.min(off, block.content.size), block.content.size, ' ', ' ');
    out.push({ name: node.attrs.name as string, pos, context: (after || text).trim().slice(0, 60) });
    return false;
  });
  return out;
}

export const findBookmark = (doc: PMNode, name: string) => listBookmarks(doc).find((b) => b.name.toLowerCase() === name.toLowerCase()) ?? null;

/** Ссылки «#имя» в документе: позиции и текст. */
export function linksTo(doc: PMNode, name: string): { from: number; to: number; text: string }[] {
  const out: { from: number; to: number; text: string }[] = [];
  const href = '#' + name.toLowerCase();
  doc.descendants((node, pos) => {
    if (!node.isText) return true;
    const link = node.marks.find((m) => m.type === schema.marks.link);
    if (!link || (link.attrs.href as string).toLowerCase() !== href) return false;
    const last = out[out.length - 1];
    if (last && last.to === pos) { last.to = pos + node.nodeSize; last.text += node.text; } else out.push({ from: pos, to: pos + node.nodeSize, text: node.text ?? '' });
    return false;
  });
  return out;
}

/** Ссылки на закладки, которых нет в документе. */
export function brokenLinks(doc: PMNode): { href: string; from: number; to: number; text: string }[] {
  const names = new Set(listBookmarks(doc).map((b) => b.name.toLowerCase()));
  const out: { href: string; from: number; to: number; text: string }[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText) return true;
    const link = node.marks.find((m) => m.type === schema.marks.link);
    const href = link?.attrs.href as string | undefined;
    if (!href?.startsWith('#') || names.has(href.slice(1).toLowerCase())) return false;
    const last = out[out.length - 1];
    if (last && last.to === pos && last.href === href) { last.to += node.nodeSize; last.text += node.text; } else out.push({ href, from: pos, to: pos + node.nodeSize, text: node.text ?? '' });
    return false;
  });
  return out;
}

export function validateName(doc: PMNode, name: string, except?: string): string | null {
  if (!name) return 'Введите имя закладки.';
  if (!BOOKMARK_RE.test(name)) return 'Имя начинается с буквы и состоит из букв, цифр и «_» (без пробелов), не длиннее 40 знаков.';
  const other = findBookmark(doc, name);
  if (other && other.name.toLowerCase() !== except?.toLowerCase()) return `Закладка «${other.name}» уже есть. Перенести её сюда можно кнопкой «Добавить» с тем же именем — или выберите другое имя.`;
  return null;
}

/** Поставить закладку в начало выделения. Закладка с таким именем переносится (как в Word). */
export function addBookmark(state: EditorState, name: string): Transaction {
  const tr = state.tr;
  const old = findBookmark(state.doc, name);
  if (old) tr.delete(old.pos, old.pos + 1);
  const at = tr.mapping.map(state.selection.from);
  tr.insert(at, schema.nodes.bookmark.create({ name }));
  return closeHistory(tr);
}

export function removeBookmark(state: EditorState, name: string): Transaction | null {
  const b = findBookmark(state.doc, name);
  return b ? closeHistory(state.tr.delete(b.pos, b.pos + 1)) : null;
}

/** Переименовать закладку и все ссылки на неё. */
export function renameBookmark(state: EditorState, from: string, to: string): Transaction | null {
  const b = findBookmark(state.doc, from);
  if (!b) return null;
  const tr = state.tr.setNodeMarkup(b.pos, undefined, { name: to });
  for (const l of linksTo(state.doc, from)) {
    tr.removeMark(l.from, l.to, schema.marks.link);
    tr.addMark(l.from, l.to, schema.marks.link.create({ href: '#' + to }));
  }
  return closeHistory(tr);
}

export function gotoBookmark(view: EditorView, name: string): boolean {
  const b = findBookmark(view.state.doc, name);
  if (!b) return false;
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, b.pos + 1)).scrollIntoView());
  // Мигающая отметка: видно, куда перешли.
  const dom = view.nodeDOM(b.pos) as HTMLElement | null;
  if (dom?.classList) { dom.classList.remove('bm-flash'); void dom.offsetWidth; dom.classList.add('bm-flash'); }
  view.focus();
  return true;
}

/** Уникальность имён: при вставке копии текста с закладкой повтор убирается (остаётся первая по тексту). */
export const uniqueBookmarks = new Plugin({
  appendTransaction(trs, _old, state) {
    // Проверяем документ, только если вставлен кусок с закладкой: на обычном наборе цена нулевая.
    const bringsBookmark = trs.some((t) => t.steps.some((step) => {
      const slice = (step as unknown as { slice?: { content: { descendants(f: (n: PMNode) => boolean | void): void } } }).slice;
      let found = false;
      slice?.content.descendants((n) => { if (n.type === schema.nodes.bookmark) found = true; return !found; });
      return found;
    }));
    if (!bringsBookmark) return null;
    const seen = new Set<string>();
    const dups: number[] = [];
    state.doc.descendants((node, pos) => {
      if (node.type !== schema.nodes.bookmark) return true;
      const k = (node.attrs.name as string).toLowerCase();
      if (seen.has(k)) dups.push(pos); else seen.add(k);
      return false;
    });
    if (!dups.length) return null;
    const tr = state.tr;
    for (const pos of dups.reverse()) tr.delete(pos, pos + 1);
    return tr.setMeta('addToHistory', false);
  },
});
