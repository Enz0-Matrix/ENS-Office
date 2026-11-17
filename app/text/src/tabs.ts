// SPDX-License-Identifier: AGPL-3.0-or-later
// Табуляция как в Word. Браузер отсчитывает шаг табуляции (tab-size) от начала блока, то есть от отступа абзаца,
// а Word — от поля страницы (шаг 1,27 см), и при выступе ставит ещё неявный стоп на отступе абзаца.
// Каждый символ табуляции оборачивается декорацией; после отрисовки замеряем, где он кончается, и доводим конец до
// стопа Word отступом справа (margin-right может быть и отрицательным). Документ не меняется: это только вид.

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';

/** Шаг табуляции по умолчанию, pt (Word: 720 twips). */
export const TAB_STEP_PT = 36;
const PX_PT = 4 / 3;

interface TabState { /** Поправка конца табуляции, px листа (без масштаба), по позиции символа. */ shift: Map<number, number>; deco: DecorationSet }

const key = new PluginKey<TabState>('ens-tabs');

function build(doc: PMNode, shift: Map<number, number>): DecorationSet {
  const out: Decoration[] = [];
  doc.descendants((n, pos) => {
    if (!n.isText) return true;
    const t = n.text!;
    for (let i = t.indexOf('\t'); i >= 0; i = t.indexOf('\t', i + 1)) {
      const p = pos + i;
      const m = shift.get(p);
      out.push(Decoration.inline(p, p + 1, { class: 'tab', ...(m ? { style: `margin-right:${m.toFixed(2)}px` } : {}) }));
    }
    return false;
  });
  return DecorationSet.create(doc, out);
}

/** Стоп Word после x (экранные px): неявный — на отступе при выступе, иначе ближайший шаг от поля. */
function wordStop(x: number, origin: number, hang: number | null, step: number): number {
  if (hang != null && x < hang - 1) return hang;
  // Табуляция, начатая на стопе (с точностью до пикселя), идёт к следующему — как в Word.
  return origin + step * (Math.floor((x - origin + 1) / step) + 1);
}

/** Блок текста, внутри которого стоит табуляция, и его отступы. */
function blockOf(view: EditorView, span: HTMLElement): { el: HTMLElement; node: PMNode } | null {
  const el = span.closest('p, h1, h2, h3, h4, h5, h6') as HTMLElement | null;
  if (!el || !view.dom.contains(el)) return null;
  let pos: number;
  try { pos = view.posAtDOM(span, 0); } catch { return null; }
  const node = view.state.doc.resolve(pos).parent;
  return node.isTextblock ? { el, node } : null;
}

/** Замер: новые поправки (или null, если всё уже на местах). */
function measure(view: EditorView): Map<number, number> | null {
  const st = key.getState(view.state);
  if (!st) return null;
  const spans = [...view.dom.querySelectorAll<HTMLElement>('span.tab')];
  if (!spans.length) return null;
  const k = view.dom.getBoundingClientRect().width / (view.dom.offsetWidth || 1) || 1;
  const step = TAB_STEP_PT * PX_PT * k;
  const next = new Map<number, number>();
  let changed = false;
  let block: HTMLElement | null = null;
  let origin = 0, hang: number | null = null;
  for (const span of spans) {
    const b = blockOf(view, span);
    if (!b) continue;
    let pos: number;
    try { pos = view.posAtDOM(span, 0); } catch { continue; }
    if (b.el !== block) {
      block = b.el;
      const left = block.getBoundingClientRect().left;
      const indent = ((b.node.attrs.indent as number | null) ?? 0) * PX_PT * k;
      const first = (b.node.attrs.first as number | null) ?? 0;
      origin = left - indent; // поле страницы (или край ячейки) — от него считает Word
      hang = first < 0 ? left : null;
    }
    // Свой конец браузер ставит на сетку tab-size, отсчитанную от отступа. Ставим конец на стоп Word:
    // поправка = стоп − конец без поправки. Табуляции правее на той же строке сдвинутся и
    // доведутся следующим проходом (их начало станет известно после этого).
    const r = span.getBoundingClientRect();
    const cur = (st.shift.get(pos) ?? 0) * k;
    const target = wordStop(r.left, origin, hang, step);
    const m = (target - r.right) / k;
    if (Math.abs(m) > 0.05) next.set(pos, Math.round(m * 100) / 100);
    if (Math.abs(target - (r.right + cur)) > 0.5) changed = true;
  }
  if (!changed) return null;
  return next;
}

export const tabsPlugin = new Plugin<TabState>({
  key,
  state: {
    init: (_, s: EditorState) => ({ shift: new Map(), deco: build(s.doc, new Map()) }),
    apply: (tr, prev) => {
      const meta = tr.getMeta(key) as Map<number, number> | undefined;
      if (meta) return { shift: meta, deco: build(tr.doc, meta) };
      if (!tr.docChanged) return prev;
      // Поправки едут за текстом до следующего замера.
      const shift = new Map<number, number>();
      for (const [p, m] of prev.shift) {
        const r = tr.mapping.mapResult(p, 1);
        if (!r.deleted) shift.set(r.pos, m);
      }
      return { shift, deco: build(tr.doc, shift) };
    },
  },
  props: { decorations: (s) => key.getState(s)?.deco },
  view: (view) => {
    let raf = 0;
    let passes = 0;
    const run = () => {
      raf = 0;
      const next = measure(view);
      if (!next || ++passes > 12) return; // защита от зацикливания при переносах строк
      view.dispatch(view.state.tr.setMeta(key, next).setMeta('addToHistory', false));
    };
    const schedule = () => { if (!raf) raf = requestAnimationFrame(run); };
    schedule();
    return {
      update: (v, prev) => {
        const own = key.getState(v.state) !== key.getState(prev) && v.state.doc === prev.doc;
        if (!own) passes = 0;
        schedule();
      },
      destroy: () => cancelAnimationFrame(raf),
    };
  },
});

/** Для тестов: поправки табуляции на сейчас. */
export const tabShifts = (s: EditorState): Map<number, number> => key.getState(s)?.shift ?? new Map();
