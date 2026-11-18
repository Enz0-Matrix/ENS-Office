// SPDX-License-Identifier: AGPL-3.0-or-later
// Подписи рисунков, таблиц и формул: «Рисунок 1 – Схема», «Таблица 2 – Итоги» (ГОСТ 7.32, ГОСТ 2.105).
//
// Номер — узел `seq` (поле SEQ в docx). Числа считает этот плагин по порядку в документе при каждой правке,
// поэтому «Обновить поле», как в Word, не нужно: вставили рисунок в середину — номера ниже сдвинулись сами.

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { schema } from './schema';

export const CAPTION_LABELS = ['Рисунок', 'Таблица', 'Формула'];
/** Разделитель по ГОСТ: тире с пробелами. */
export const CAPTION_SEP = ' – ';

export interface CaptionInfo { name: string; n: number; pos: number; block: number; text: string }

/** Все номера по порядку: имя, число, позиция узла seq, позиция абзаца подписи, текст подписи после номера. */
export function listCaptions(doc: PMNode): CaptionInfo[] {
  const count = new Map<string, number>();
  const out: CaptionInfo[] = [];
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    node.forEach((child, off) => {
      if (child.type !== schema.nodes.seq) return;
      const name = child.attrs.name as string;
      const n = (count.get(name) ?? 0) + 1;
      count.set(name, n);
      const after = node.textBetween(off + 1, node.content.size, ' ', ' ').replace(/^\s*[–—-]\s*/, '').trim();
      out.push({ name, n, pos: pos + 1 + off, block: pos, text: after });
    });
    return false;
  });
  return out;
}

function build(doc: PMNode): DecorationSet {
  const decos = listCaptions(doc).map((c) => Decoration.node(c.pos, c.pos + 1, { 'data-n': String(c.n) }));
  return decos.length ? DecorationSet.create(doc, decos) : DecorationSet.empty;
}

export const captionsPlugin: Plugin<DecorationSet> = new Plugin<DecorationSet>({
  state: {
    init: (_c, state) => build(state.doc),
    apply: (tr, prev, _o, state) => (tr.docChanged ? build(state.doc) : prev),
  },
  props: { decorations(state) { return captionsPlugin.getState(state); } },
});

/** Что подписывать: таблица под курсором — над ней; рисунок — под абзацем с ним; иначе — новая строка после абзаца. */
export function captionTarget(state: EditorState): { kind: 'table' | 'image' | 'text'; before: number; after: number } {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    if ($from.node(d).type === schema.nodes.table) return { kind: 'table', before: $from.before(d), after: $from.after(d) };
  }
  // Курсор в абзаце: какой блок верхнего уровня — от него и считаем.
  const depth = Math.min(1, $from.depth);
  const before = depth ? $from.before(depth) : $from.pos, after = depth ? $from.after(depth) : $from.pos;
  const sel = state.selection as { node?: PMNode };
  let image = sel.node?.type === schema.nodes.image;
  if (!image && depth) $from.node(depth).descendants((n) => { if (n.type === schema.nodes.image) image = true; return !image; });
  return { kind: image ? 'image' : 'text', before, after };
}

export interface CaptionOptions { name: string; text: string; above: boolean }

/** Вставить подпись: абзац «Имя N – текст» над или под объектом; рисунок — по центру, таблица — слева (ГОСТ). */
export function insertCaption(state: EditorState, o: CaptionOptions): Transaction {
  const t = captionTarget(state);
  const content: PMNode[] = [schema.text(o.name + ' '), schema.nodes.seq.create({ name: o.name })];
  if (o.text.trim()) content.push(schema.text(CAPTION_SEP + o.text.trim()));
  const para = schema.nodes.paragraph.create({ align: o.above ? null : 'center', after: o.above ? 6 : 12, before: o.above ? 12 : 6 }, content);
  const at = o.above ? t.before : t.after;
  const tr = state.tr.insert(at, para);
  return tr.setSelection(TextSelection.create(tr.doc, at + para.nodeSize - 1)).scrollIntoView();
}

/** Число для подписи: то, которое получит новая подпись в этом месте (для предпросмотра в окне). */
export function nextNumber(state: EditorState, name: string, above: boolean): number {
  const t = captionTarget(state);
  const at = above ? t.before : t.after;
  return listCaptions(state.doc).filter((c) => c.name === name && c.pos < at).length + 1;
}
