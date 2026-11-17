// SPDX-License-Identifier: AGPL-3.0-or-later
// Скрытые символы, как «Отобразить все знаки» в Word (Ctrl+*): ¶ в конце абзаца, · на месте пробела,
// ° на неразрывном пробеле, → на табуляции, ↵ на разрыве строки. Это только декорации: документ, вёрстка и печать
// не меняются (знаки не занимают места и не печатаются).

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';

const key = new PluginKey<{ on: boolean; deco: DecorationSet }>('ens-hidden-marks');
const STORE = 'ens-text.marks';

let wanted = (() => { try { return localStorage.getItem(STORE) === '1'; } catch { return false; } })();

/** Знак-виджет нулевой ширины: сам символ рисует CSS (::after), поэтому строки не сдвигаются. */
const mark = (cls: string) => (): HTMLElement => {
  const s = document.createElement('span');
  s.className = `hm-w ${cls}`;
  s.contentEditable = 'false';
  return s;
};
const pilcrow = mark('hm-para');

/** Знаки для куска документа [from, to) — целые блоки верхнего уровня. */
function collect(doc: PMNode, from: number, to: number): Decoration[] {
  const out: Decoration[] = [];
  doc.nodesBetween(from, to, (n, pos) => {
    if (n.isTextblock) {
      out.push(Decoration.widget(pos + n.nodeSize - 1, pilcrow, { side: 1, key: 'hm-p', ignoreSelection: true }));
      return true;
    }
    if (n.type.name === 'hard_break') { out.push(Decoration.widget(pos, mark('hm-br'), { side: -1, key: 'hm-br', ignoreSelection: true })); return false; }
    if (!n.isText) return true;
    const t = n.text!;
    for (let i = 0; i < t.length; i++) {
      const c = t[i];
      if (c === ' ') out.push(Decoration.inline(pos + i, pos + i + 1, { class: 'hm-sp' }));
      else if (c === '\u00a0') out.push(Decoration.inline(pos + i, pos + i + 1, { class: 'hm-nb' }));
    }
    return false;
  });
  return out;
}

const build = (doc: PMNode): DecorationSet => DecorationSet.create(doc, collect(doc, 0, doc.content.size));

/** После правки пересобрать знаки только в затронутых блоках верхнего уровня: в большом документе это важно. */
function update(prev: DecorationSet, tr: Transaction): DecorationSet {
  const doc = tr.doc;
  let deco = prev.map(tr.mapping, doc);
  let lo = Infinity, hi = -Infinity;
  tr.mapping.maps.forEach((map, i) => {
    const rest = tr.mapping.slice(i + 1);
    map.forEach((_a, _b, start, end) => { lo = Math.min(lo, rest.map(start, -1)); hi = Math.max(hi, rest.map(end, 1)); });
  });
  if (lo > hi) return deco;
  // Расширяем до границ блоков верхнего уровня.
  const a = doc.resolve(Math.max(0, Math.min(lo, doc.content.size)));
  const b = doc.resolve(Math.max(0, Math.min(hi, doc.content.size)));
  const from = a.depth ? a.before(1) : a.pos;
  const to = b.depth ? b.after(1) : b.pos;
  deco = deco.remove(deco.find(from, to));
  return deco.add(doc, collect(doc, from, to));
}

export const hiddenMarksPlugin = new Plugin<{ on: boolean; deco: DecorationSet }>({
  key,
  state: {
    init: (_, s: EditorState) => ({ on: wanted, deco: wanted ? build(s.doc) : DecorationSet.empty }),
    apply: (tr, prev) => {
      const m = tr.getMeta(key) as boolean | undefined;
      const on = m ?? prev.on;
      if (!on) return prev.on ? { on, deco: DecorationSet.empty } : prev;
      if (m !== undefined && !prev.on) return { on, deco: build(tr.doc) };
      return tr.docChanged ? { on, deco: update(prev.deco, tr) } : prev;
    },
  },
  props: {
    decorations: (s) => key.getState(s)?.deco,
    // Табуляция уже обёрнута (tabs.ts) — ей стрелку рисует CSS по этому классу.
    attributes: (s): Record<string, string> => (key.getState(s)?.on ? { class: 'show-marks' } : {}),
  },
});

export const hiddenMarksOn = (s: EditorState): boolean => !!key.getState(s)?.on;

/** Включить или выключить знаки; выбор запоминается для следующих запусков. */
export function setHiddenMarks(view: EditorView, on: boolean): void {
  wanted = on;
  try { localStorage.setItem(STORE, on ? '1' : '0'); } catch { /* без хранилища — до закрытия окна */ }
  view.dispatch(view.state.tr.setMeta(key, on).setMeta('addToHistory', false));
}
