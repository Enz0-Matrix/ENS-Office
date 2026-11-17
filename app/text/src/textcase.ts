// SPDX-License-Identifier: AGPL-3.0-or-later
// «Регистр» как в Word (кнопка Aa, Shift+F3): ВСЕ ПРОПИСНЫЕ, все строчные, Как в предложениях, Каждое Слово С Прописной,
// иЗМЕНИТЬ рЕГИСТР (случайно нажатый Caps Lock: «пРИВЕТ» → «Привет»). Меняются только буквы, оформление и длина текста
// сохраняются (буква, которая при смене регистра стала бы двумя, например ß, остаётся как есть). Без выделения — слово
// под курсором.

import { closeHistory } from 'prosemirror-history';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';

export type CaseMode = 'upper' | 'lower' | 'sentence' | 'title' | 'toggle' | 'cycle';

const up = (c: string) => { const u = c.toUpperCase(); return u.length === 1 ? u : c; };
const low = (c: string) => { const l = c.toLowerCase(); return l.length === 1 ? l : c; };
const isLetter = (c: string) => /\p{L}/u.test(c);

/** Преобразовать строку; s — текст выделения, абзацы разделены '\n'; before — символ перед выделением. */
export function convertCase(s: string, mode: Exclude<CaseMode, 'cycle'>, before = ''): string {
  const chars = [...s];
  switch (mode) {
    case 'upper': return chars.map(up).join('');
    case 'lower': return chars.map(low).join('');
    case 'toggle': return chars.map((c) => (c === up(c) ? low(c) : up(c))).join('');
    case 'title': {
      let prev = before.slice(-1);
      return chars.map((c) => { const r = isLetter(c) && !/[\p{L}\p{N}'’]/u.test(prev) ? up(c) : low(c); prev = c; return r; }).join('');
    }
    case 'sentence': {
      // Начало предложения: начало абзаца или после . ! ? … (пробелы, кавычки и скобки между ними не мешают).
      let start = /^[\s«"„“(]*$/.test(before) || /[.!?…][\s«»"„“”)]*$/.test(before);
      return chars.map((c) => {
        if (isLetter(c)) { const r = start ? up(c) : low(c); start = false; return r; }
        if (/[.!?…\n]/.test(c)) start = true;
        else if (/\p{N}/u.test(c)) start = false;
        return c;
      }).join('');
    }
  }
}

/** Следующий режим для Shift+F3: смешанный или строчный с прописной → ВСЕ ПРОПИСНЫЕ → все строчные → Как в предложениях. */
export function nextCase(s: string): Exclude<CaseMode, 'cycle'> {
  const letters = [...s].filter(isLetter);
  const hasLow = letters.some((c) => c !== up(c)), hasUp = letters.some((c) => c !== low(c));
  if (!hasLow && hasUp) return 'lower';
  if (hasLow && !hasUp) return 'sentence';
  return 'upper';
}

/** Сменить регистр выделения (или слова под курсором). null — менять нечего. */
export function changeCase(state: EditorState, mode: CaseMode): Transaction | null {
  let { from, to } = state.selection;
  const word = from === to;
  const tr = state.tr;
  if (from === to) {
    // Слово под курсором.
    const $p = state.doc.resolve(from);
    const text = $p.parent.textContent, off = $p.parentOffset;
    let a = off, b = off;
    while (a > 0 && /[\p{L}\p{N}'’-]/u.test(text[a - 1])) a--;
    while (b < text.length && /[\p{L}\p{N}'’-]/u.test(text[b])) b++;
    if (a === b) return null;
    from = $p.start() + a; to = $p.start() + b;
    // Позиции в тексте абзаца совпадают с позициями в документе, только если в абзаце нет вставок (рисунков);
    // иначе берём слово по тексту ProseMirror.
    if (state.doc.textBetween(from, to) !== text.slice(a, b)) return null;
  }
  const segs: { pos: number; text: string }[] = [];
  let all = '';
  let lastBlock = -1;
  state.doc.nodesBetween(from, to, (n, pos) => {
    if (!n.isText) return true;
    const s = Math.max(from, pos), e = Math.min(to, pos + n.nodeSize);
    if (e <= s) return false;
    const blockStart = state.doc.resolve(s).start();
    if (lastBlock >= 0 && blockStart !== lastBlock) all += '\n';
    lastBlock = blockStart;
    const t = n.text!.slice(s - pos, e - pos);
    segs.push({ pos: s, text: t });
    all += t;
    return false;
  });
  if (!all) return null;
  const $f = state.doc.resolve(from);
  const before = $f.parent.textBetween(0, $f.parentOffset, '', '￼').slice(-3);
  let m = mode === 'cycle' ? nextCase(all) : mode;
  if (word && mode === 'cycle' && m === 'sentence') m = 'title'; // слово: хорошо → Хорошо → ХОРОШО → хорошо, как в Word
  const out = convertCase(all, m, before);
  // Раскладываем результат обратно по кускам (длина в UTF-16 та же, что у исходного).
  let i = 0;
  let changed = false;
  for (const seg of segs) {
    while (out[i] === '\n' && all[i] === '\n') i++;
    const t = out.slice(i, i + seg.text.length);
    i += seg.text.length;
    if (t === seg.text) continue;
    changed = true;
    const node = state.doc.nodeAt(seg.pos)!;
    tr.replaceWith(seg.pos, seg.pos + t.length, state.schema.text(t, node.marks));
  }
  if (!changed) return null;
  tr.setSelection(TextSelection.create(tr.doc, from, to));
  return closeHistory(tr); // каждая смена регистра — отдельный шаг отмены, даже подряд
}
