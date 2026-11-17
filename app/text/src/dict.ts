// SPDX-License-Identifier: AGPL-3.0-or-later
// Словарь ENS-Text: слова, которые проверка орфографии не должна подчёркивать (фамилии, термины, названия).
// Проверку делает сам браузер (WebView, Chromium), и добавить слово в его словарь со страницы нельзя. Поэтому слова
// из нашего словаря оборачиваются декорацией с атрибутом spellcheck="false": браузер их не проверяет и не
// подчёркивает. Документ не меняется. Встроенные слова — фамилия автора во всех формах; свои слова хранятся
// у пользователя (localStorage) и правятся в окне «Словарь».

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import { declineFullName } from './russian';

const STORE = 'ens-text.dict';

/** Формы фамилии на -ов/-ев/-ёв/-ин/-ын: мужская и женская по падежам и множественное число. */
export function surnameForms(word: string): string[] {
  const w = word.trim();
  if (!/^[А-ЯЁ][а-яё-]+(ов|ев|ёв|ин|ын)$/.test(w)) return [w];
  const male = Object.values(declineFullName(w, { gender: 'm' }).forms);
  const female = Object.values(declineFullName(`${w}а`, { gender: 'f' }).forms);
  const plural = ['ы', 'ых', 'ым', 'ыми'].map((e) => w + e);
  return [...new Set([...male, ...female, ...plural, `${w}ою`])];
}

/** Встроенные слова: фамилия автора программы во всех формах. */
export const BUILTIN_WORDS: readonly string[] = surnameForms('Едренников');

let user: string[] = (() => {
  try {
    const v = JSON.parse(localStorage.getItem(STORE) ?? '[]');
    return Array.isArray(v) ? v.filter((x): x is string => typeof x === 'string') : [];
  } catch { return []; }
})();
let known = new Set<string>();
const rebuild = () => { known = new Set([...BUILTIN_WORDS, ...user].map((w) => w.toLowerCase())); };
rebuild();

export const userWords = (): string[] => [...user];
export const isKnownWord = (w: string): boolean => known.has(w.toLowerCase());

const key = new PluginKey<DecorationSet>('ens-dict');
const WORD = /[\p{L}][\p{L}\p{M}'’-]*/gu;

function collect(doc: PMNode, from: number, to: number): Decoration[] {
  const out: Decoration[] = [];
  doc.nodesBetween(from, to, (n, pos) => {
    if (!n.isText) return true;
    for (const m of n.text!.matchAll(WORD)) {
      if (!known.has(m[0].toLowerCase())) continue;
      const a = pos + m.index!;
      out.push(Decoration.inline(a, a + m[0].length, { spellcheck: 'false', class: 'dict-word' }));
    }
    return false;
  });
  return out;
}

const build = (doc: PMNode) => DecorationSet.create(doc, collect(doc, 0, doc.content.size));

/** Пересобрать только затронутые блоки верхнего уровня (как у скрытых символов). */
function update(prev: DecorationSet, tr: Transaction): DecorationSet {
  const doc = tr.doc;
  let deco = prev.map(tr.mapping, doc);
  let lo = Infinity, hi = -Infinity;
  tr.mapping.maps.forEach((map, i) => {
    const rest = tr.mapping.slice(i + 1);
    map.forEach((_a, _b, start, end) => { lo = Math.min(lo, rest.map(start, -1)); hi = Math.max(hi, rest.map(end, 1)); });
  });
  if (lo > hi) return deco;
  const a = doc.resolve(Math.max(0, Math.min(lo, doc.content.size)));
  const b = doc.resolve(Math.max(0, Math.min(hi, doc.content.size)));
  const from = a.depth ? a.before(1) : a.pos;
  const to = b.depth ? b.after(1) : b.pos;
  deco = deco.remove(deco.find(from, to));
  return deco.add(doc, collect(doc, from, to));
}

export const dictPlugin = new Plugin<DecorationSet>({
  key,
  state: {
    init: (_, s: EditorState) => build(s.doc),
    apply: (tr, prev) => (tr.getMeta(key) ? build(tr.doc) : tr.docChanged ? update(prev, tr) : prev),
  },
  props: { decorations: (s) => key.getState(s) },
});

function save(view: EditorView | null): void {
  user = [...new Set(user)].sort((a, b) => a.localeCompare(b, 'ru'));
  try { localStorage.setItem(STORE, JSON.stringify(user)); } catch { /* без хранилища — до закрытия окна */ }
  rebuild();
  if (view) view.dispatch(view.state.tr.setMeta(key, true).setMeta('addToHistory', false));
}

/** Добавить слово (и, для фамилии на -ов/-ин, все его формы). Возвращает добавленные слова. */
export function addWords(view: EditorView | null, word: string, withForms = true): string[] {
  const w = word.trim();
  if (!w || /\s/.test(w)) return [];
  const add = (withForms ? surnameForms(w) : [w]).filter((x) => !known.has(x.toLowerCase()));
  user.push(...add);
  save(view);
  return add;
}

export function removeWord(view: EditorView | null, word: string): void {
  user = user.filter((x) => x.toLowerCase() !== word.toLowerCase());
  save(view);
}

/** Слово под курсором (для пункта «Добавить в словарь» контекстного меню). */
export function wordAt(state: EditorState): string | null {
  const $p = state.selection.$from;
  if (!$p.parent.isTextblock) return null;
  const text = $p.parent.textBetween(0, $p.parent.content.size, undefined, '￼');
  const off = $p.parentOffset;
  for (const m of text.matchAll(WORD)) if (m.index! <= off && off <= m.index! + m[0].length) return m[0];
  return null;
}
