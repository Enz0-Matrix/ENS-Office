// SPDX-License-Identifier: AGPL-3.0-or-later
// Перекрёстные ссылки (Ссылки → Перекрёстная ссылка): «см. рисунок 3», «таблица 2», «раздел „Введение“», «с. 5».
//
// Ссылка — узел `xref` с именем цели (закладки) и видом текста. На подписи и заголовки ссылаемся, как Word, через
// скрытую закладку «_Ref…» в начале абзаца цели: номер и текст берутся из документа при каждой правке, поэтому
// «Обновить поле» не нужно. Лучше, чем в Word: если цель удалили, ссылка сразу показывает «‹нет цели›», программа
// предупреждает, а окно перечисляет все такие ссылки с переходом к каждой.

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import { findBookmark, isHidden, listBookmarks } from './bookmarks';
import { listCaptions } from './captions';
import { schema } from './schema';
import { dialog, el, toast } from './ui';

export type XrefShow = 'label' | 'text' | 'page' | 'above';
export const SHOW_LABEL: Record<XrefShow, string> = { label: 'Номер («рисунок 3»)', text: 'Текст целиком', page: 'Номер страницы', above: '«выше» / «ниже»' };

export interface Target { kind: 'caption' | 'heading' | 'bookmark'; title: string; block: number; label?: string; bookmark?: string }

/** Куда можно сослаться: подписи, заголовки, видимые закладки. block — позиция абзаца цели. */
export function listTargets(doc: PMNode): Target[] {
  const out: Target[] = [];
  for (const c of listCaptions(doc)) out.push({ kind: 'caption', title: `${c.name} ${c.n}${c.text ? ' – ' + c.text : ''}`, block: c.block, label: `${c.name.toLowerCase()} ${c.n}` });
  doc.descendants((n, pos) => {
    if (n.type === schema.nodes.heading && n.textContent.trim()) out.push({ kind: 'heading', title: n.textContent.trim(), block: pos });
    return n.type !== schema.nodes.heading;
  });
  for (const b of listBookmarks(doc)) if (!isHidden(b.name)) out.push({ kind: 'bookmark', title: b.name, block: doc.resolve(b.pos).before(), bookmark: b.name });
  return out;
}

/** Имя скрытой закладки в начале абзаца или null. */
function refBookmarkIn(doc: PMNode, block: number): string | null {
  const node = doc.nodeAt(block);
  let name: string | null = null;
  node?.forEach((c) => { if (!name && c.type === schema.nodes.bookmark && /^_Ref/.test(c.attrs.name as string)) name = c.attrs.name as string; });
  return name;
}

/** Вставить ссылку на цель: при необходимости ставит скрытую закладку «_Ref…» в начало абзаца цели. */
export function insertXref(state: EditorState, t: Target, show: XrefShow): Transaction {
  const tr = state.tr;
  let name = t.bookmark ?? refBookmarkIn(state.doc, t.block);
  if (!name) {
    name = `_Ref${Date.now().toString(36)}${Math.floor(Math.random() * 1296).toString(36)}`;
    tr.insert(t.block + 1, schema.nodes.bookmark.create({ name }));
  }
  const at = tr.mapping.map(state.selection.from);
  tr.insert(at, schema.nodes.xref.create({ target: name, show }));
  return tr.setSelection(TextSelection.create(tr.doc, at + 1)).scrollIntoView();
}

let pageOf: (pos: number) => number | undefined = () => undefined;
export function setXrefPages(f: (pos: number) => number | undefined): void { pageOf = f; }

/** Текст ссылки по цели (или null — цели нет). */
export function xrefText(doc: PMNode, node: PMNode, pos: number): string | null {
  const b = findBookmark(doc, node.attrs.target as string);
  if (!b) return null;
  const $b = doc.resolve(b.pos);
  const block = $b.before();
  const para = $b.parent;
  const show = node.attrs.show as XrefShow;
  if (show === 'page') return String(pageOf(b.pos) ?? '?');
  if (show === 'above') return b.pos < pos ? 'выше' : 'ниже';
  const cap = listCaptions(doc).find((c) => c.block === block);
  if (show === 'label' && cap) return `${cap.name.toLowerCase()} ${cap.n}`;
  const text = para.textContent.replace(/\s+/g, ' ').trim();
  if (cap) return `${cap.name} ${cap.n}${cap.text ? ' – ' + cap.text : ''}`;
  return text || (node.attrs.target as string);
}

/** Ссылки без цели (закладку удалили вместе с абзацем). */
export function brokenXrefs(doc: PMNode): { pos: number; target: string }[] {
  const names = new Set(listBookmarks(doc).map((b) => b.name.toLowerCase()));
  const out: { pos: number; target: string }[] = [];
  doc.descendants((n, pos) => {
    if (n.type === schema.nodes.xref && !names.has((n.attrs.target as string).toLowerCase())) out.push({ pos, target: n.attrs.target as string });
    return true;
  });
  return out;
}

function build(doc: PMNode): DecorationSet {
  const decos: Decoration[] = [];
  doc.descendants((n, pos) => {
    if (n.type !== schema.nodes.xref) return true;
    const t = xrefText(doc, n, pos);
    decos.push(Decoration.node(pos, pos + 1, t == null ? { 'data-text': '‹нет цели›', class: 'xref-broken' } : { 'data-text': t }));
    return false;
  });
  return decos.length ? DecorationSet.create(doc, decos) : DecorationSet.empty;
}

let onBroken: (n: number) => void = () => {};
export function setXrefWarning(f: (n: number) => void): void { onBroken = f; }

export const xrefPlugin: Plugin<DecorationSet> = new Plugin<DecorationSet>({
  state: {
    init: (_c, state) => build(state.doc),
    apply: (tr, prev, old, state) => {
      if (tr.getMeta(xrefPlugin) === 'refresh') return build(state.doc);
      if (!tr.docChanged) return prev;
      // Удалили цель — предупреждаем (один раз на правку, только если «битых» стало больше).
      const before = brokenXrefs(old.doc).length, after = brokenXrefs(state.doc).length;
      if (after > before) setTimeout(() => onBroken(after - before), 0);
      return build(state.doc);
    },
  },
  props: {
    decorations(state) { return xrefPlugin.getState(state); },
    // Ctrl+щелчок по ссылке — переход к цели, как у гиперссылок.
    handleClickOn(view, _pos, node, nodePos, event) {
      if (node.type !== schema.nodes.xref || !(event.ctrlKey || event.metaKey)) return false;
      const b = findBookmark(view.state.doc, node.attrs.target as string);
      if (!b) { toast('Цель ссылки удалена.'); return true; }
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, b.pos + 1)).scrollIntoView());
      void nodePos;
      return true;
    },
  },
});

/** После перестройки страниц: перерисовать ссылки «с. N» (служебная транзакция, в историю не попадает). */
export function refreshXrefPages(view: EditorView): void {
  let has = false;
  view.state.doc.descendants((n) => { if (n.type === schema.nodes.xref && n.attrs.show === 'page') has = true; return !has; });
  if (has) view.dispatch(view.state.tr.setMeta(xrefPlugin, 'refresh').setMeta('addToHistory', false));
}

export function xrefDialog(view: EditorView): void {
  dialog((box, close) => {
    box.classList.add('xref-dialog');
    const kind = el('select', { ariaLabel: 'Тип ссылки' }, new Option('Подписи (рисунки, таблицы…)', 'caption'), new Option('Заголовки', 'heading'), new Option('Закладки', 'bookmark'));
    const show = el('select', { ariaLabel: 'Вставить ссылку на' }, ...(Object.keys(SHOW_LABEL) as XrefShow[]).map((k) => new Option(SHOW_LABEL[k], k)));
    const list = el('select', { size: 8, ariaLabel: 'Цель', class: 'xref-list' });
    const broken = el('div', { class: 'xref-broken-list' });
    let targets: Target[] = [];
    const draw = () => {
      targets = listTargets(view.state.doc).filter((t) => t.kind === kind.value);
      list.replaceChildren(...targets.map((t, i) => new Option(t.title, String(i))));
      if (targets.length) list.selectedIndex = 0;
      show.options[0].disabled = kind.value !== 'caption';
      if (kind.value !== 'caption' && show.value === 'label') show.value = 'text';
      const bad = brokenXrefs(view.state.doc);
      broken.replaceChildren(...(bad.length ? [el('b', { textContent: `Ссылки без цели: ${bad.length}` }), ...bad.map((x) => el('button', { class: 'xref-broken-item', textContent: `перейти к ссылке (${x.target})`, onclick: () => {
        close(); view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, x.pos + 1)).scrollIntoView()); view.focus();
      } }))] : []));
    };
    kind.addEventListener('change', draw);
    const ok = () => {
      const t = targets[Number(list.value)];
      if (!t) return;
      close();
      view.dispatch(insertXref(view.state, t, show.value as XrefShow));
      view.focus();
    };
    list.addEventListener('dblclick', ok);
    box.append(el('h2', { textContent: 'Перекрёстная ссылка' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Тип ссылки' }), kind),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Вставить' }), show), list, broken,
      el('p', { class: 'hint', textContent: 'Номера и текст ссылок обновляются сами. Ctrl+щелчок по ссылке — переход к цели.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Вставить', onclick: ok })));
    draw();
  }, false);
}
