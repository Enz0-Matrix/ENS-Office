// Ядро редактирования: состояние ProseMirror, плагины и команды форматирования.

import { baseKeymap, chainCommands, exitCode, joinBackward, selectParentNode, setBlockType, toggleMark, lift } from 'prosemirror-commands';
import { dropCursor } from 'prosemirror-dropcursor';
import { gapCursor } from 'prosemirror-gapcursor';
import { history } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import type { MarkType, Node as PMNode, NodeType } from 'prosemirror-model';
import { liftListItem, sinkListItem, splitListItem, wrapInList } from 'prosemirror-schema-list';
import { EditorState, Plugin, PluginKey, Selection, TextSelection, type Command, type Transaction } from 'prosemirror-state';
import { columnResizing, goToNextCell, tableEditing, TableView } from 'prosemirror-tables';
import { Decoration, DecorationSet, EditorView } from 'prosemirror-view';
import { schema } from './schema';

// ---------------------------------------------------------------- поиск

export interface FindState { query: string; matchCase: boolean; matches: { from: number; to: number }[]; current: number }
export const findKey = new PluginKey<FindState>('find');

export function findMatches(doc: PMNode, query: string, matchCase: boolean): { from: number; to: number }[] {
  const out: { from: number; to: number }[] = [];
  if (!query) return out;
  const q = matchCase ? query : query.toLowerCase();
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    // Текст блока с позициями: узлы-не-тексты (картинки) заменяем символом, чтобы смещения совпали.
    let text = '';
    node.forEach((c) => { text += c.isText ? c.text : '￼'; });
    const hay = matchCase ? text : text.toLowerCase();
    for (let i = hay.indexOf(q); i >= 0; i = hay.indexOf(q, i + Math.max(1, q.length))) {
      out.push({ from: pos + 1 + i, to: pos + 1 + i + q.length });
    }
    return false;
  });
  return out;
}

const findPlugin = new Plugin<FindState>({
  key: findKey,
  state: {
    init: () => ({ query: '', matchCase: false, matches: [], current: -1 }),
    apply(tr, prev, _old, state) {
      const meta = tr.getMeta(findKey) as Partial<FindState> | undefined;
      if (!meta && !tr.docChanged) return prev;
      const next = { ...prev, ...meta };
      next.matches = findMatches(state.doc, next.query, next.matchCase);
      if (next.current >= next.matches.length) next.current = next.matches.length - 1;
      return next;
    },
  },
  props: {
    decorations(state) {
      const s = findKey.getState(state);
      if (!s || !s.matches.length) return null;
      return DecorationSet.create(state.doc, s.matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === s.current ? 'find-hit current' : 'find-hit' })));
    },
  },
});

// ---------------------------------------------------------------- команды

const { paragraph, heading, bullet_list, ordered_list, list_item, image, hard_break } = schema.nodes;

const inList = (state: EditorState) => {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type === list_item) return true;
  return false;
};

/** Enter: в списке — новый пункт (в пустом — выход из списка), иначе обычный разрыв абзаца с теми же свойствами. */
const enter: Command = chainCommands(
  (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty || !inList(state)) return false;
    // Пустой пункт: выходим из списка.
    if ($from.parent.content.size === 0 && $from.parent.type === paragraph) return liftListItem(list_item)(state, dispatch);
    return splitListItem(list_item)(state, dispatch);
  },
  (state, dispatch) => {
    const { $from, $to, empty } = state.selection;
    if (!(state.selection instanceof TextSelection)) return false;
    const type = $from.parent.type;
    if (type !== heading && type !== paragraph) return false;
    if (dispatch) {
      let tr = state.tr;
      if (!empty) tr = tr.deleteSelection();
      const pos = tr.mapping.map($from.pos);
      const atEnd = tr.doc.resolve(pos).parentOffset === tr.doc.resolve(pos).parent.content.size;
      // Заголовок, оборвавшийся в конце, продолжается обычным абзацем.
      const next = type === heading && atEnd ? { type: paragraph, attrs: { ...$from.parent.attrs, level: undefined, pageBreakBefore: false } } : { type, attrs: { ...$from.parent.attrs, pageBreakBefore: false } };
      tr = tr.split(pos, 1, [next]);
      dispatch(tr.scrollIntoView());
    }
    void $to;
    return true;
  },
);

/** Shift+Enter — перевод строки внутри абзаца. */
const softBreak: Command = chainCommands(exitCode, (state, dispatch) => {
  if (dispatch) dispatch(state.tr.replaceSelectionWith(hard_break.create()).scrollIntoView());
  return true;
});

/** Backspace в начале пункта списка сначала убирает маркер, а не склеивает абзацы. */
const backspace: Command = (state, dispatch) => {
  const { $from, empty } = state.selection;
  if (empty && $from.parentOffset === 0 && inList(state)) {
    // Только в самом первом абзаце пункта.
    let d = $from.depth;
    while (d > 0 && $from.node(d).type !== list_item) d--;
    if ($from.index(d) === 0) return liftListItem(list_item)(state, dispatch);
  }
  return false;
};

const tab: Command = chainCommands(
  (state, dispatch) => (inList(state) ? sinkListItem(list_item)(state, dispatch) || true : false),
  goToNextCell(1),
  (state, dispatch) => {
    if (dispatch) dispatch(state.tr.insertText('\t').scrollIntoView());
    return true;
  },
);
const shiftTab: Command = chainCommands(
  (state, dispatch) => (inList(state) ? liftListItem(list_item)(state, dispatch) || true : false),
  goToNextCell(-1),
  () => true,
);

// --- Форматирование знаков

export const toggles: Record<string, MarkType> = {
  bold: schema.marks.strong, italic: schema.marks.em, underline: schema.marks.underline, strike: schema.marks.strike,
  superscript: schema.marks.sup, subscript: schema.marks.sub,
};
export const toggle = (name: string): Command => toggleMark(toggles[name]);

export function markActive(state: EditorState, type: MarkType): boolean {
  const { from, $from, to, empty } = state.selection;
  if (empty) return !!type.isInSet(state.storedMarks || $from.marks());
  return state.selection.ranges.some((r) => state.doc.rangeHasMark(r.$from.pos, r.$to.pos, type)) && from !== to;
}

type TextStyle = { font: string | null; size: number | null; color: string | null; bg: string | null };

/** Шрифт, кегль, цвет, заливка: меняет только указанное, остальное в знаке сохраняется. */
export function setTextStyle(patch: Partial<TextStyle>): Command {
  return (state, dispatch) => {
    const type = schema.marks.textStyle;
    const merge = (cur: TextStyle | undefined): TextStyle => ({ font: null, size: null, color: null, bg: null, ...cur, ...patch });
    const isEmpty = (a: TextStyle) => !a.font && !a.size && !a.color && !a.bg;
    if (!dispatch) return true;
    let tr = state.tr;
    if (state.selection.empty) {
      const cur = type.isInSet(state.storedMarks || state.selection.$from.marks());
      const next = merge(cur?.attrs as TextStyle | undefined);
      tr = tr.removeStoredMark(type);
      if (!isEmpty(next)) tr = tr.addStoredMark(type.create(next));
    } else {
      for (const r of state.selection.ranges) {
        state.doc.nodesBetween(r.$from.pos, r.$to.pos, (node, pos) => {
          if (!node.isInline) return true;
          const from = Math.max(r.$from.pos, pos);
          const to = Math.min(r.$to.pos, pos + node.nodeSize);
          const next = merge(type.isInSet(node.marks)?.attrs as TextStyle | undefined);
          tr = tr.removeMark(from, to, type);
          if (!isEmpty(next)) tr = tr.addMark(from, to, type.create(next));
          return false;
        });
      }
    }
    dispatch(tr);
    return true;
  };
}

/** Очистить форматирование знаков (без ссылок). */
export const clearFormatting: Command = (state, dispatch) => {
  if (dispatch) {
    let tr = state.tr;
    if (state.selection.empty) {
      tr = tr.setStoredMarks([]);
    } else {
      for (const r of state.selection.ranges) {
        for (const t of Object.values(schema.marks)) if (t !== schema.marks.link) tr = tr.removeMark(r.$from.pos, r.$to.pos, t);
      }
    }
    dispatch(tr);
  }
  return true;
};

export function currentStyle(state: EditorState): { font: string | null; size: number | null } {
  const marks = state.selection.empty ? state.storedMarks || state.selection.$from.marks() : state.selection.$from.marksAcross(state.selection.$to) ?? [];
  const m = schema.marks.textStyle.isInSet(marks);
  return { font: (m?.attrs.font as string) ?? null, size: (m?.attrs.size as number) ?? null };
}

// --- Абзацы

function eachTextblock(state: EditorState, f: (node: PMNode, pos: number) => void): void {
  const seen = new Set<number>();
  for (const r of state.selection.ranges) {
    state.doc.nodesBetween(r.$from.pos, r.$to.pos, (node, pos) => {
      if (node.isTextblock && !seen.has(pos)) { seen.add(pos); f(node, pos); }
      return true;
    });
  }
}

export function setParaAttrs(patch: Record<string, unknown>): Command {
  return (state, dispatch) => {
    if (dispatch) {
      const tr = state.tr;
      eachTextblock(state, (node, pos) => tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...patch }));
      dispatch(tr);
    }
    return true;
  };
}

export function paraAttr<T>(state: EditorState, name: string): T | null {
  const node = state.selection.$from.parent;
  return node.isTextblock ? ((node.attrs[name] as T) ?? null) : null;
}

/** Отступ слева: в списке — вложенность, вне списка — на 36 pt (1,27 см) на шаг. */
export function indent(dir: 1 | -1): Command {
  return (state, dispatch) => {
    if (inList(state)) return (dir > 0 ? sinkListItem(list_item) : liftListItem(list_item))(state, dispatch) || true;
    if (dispatch) {
      const tr = state.tr;
      eachTextblock(state, (node, pos) => {
        const cur = (node.attrs.indent as number | null) ?? 0;
        tr.setNodeMarkup(pos, undefined, { ...node.attrs, indent: Math.max(0, cur + dir * 36) || null });
      });
      dispatch(tr);
    }
    return true;
  };
}

/** Стиль абзаца: обычный или заголовок 1–6 (свойства абзаца — выравнивание, отступы — сохраняются). */
export function setStyle(level: number): Command {
  return (state, dispatch) => {
    if (dispatch) {
      const tr = state.tr;
      eachTextblock(state, (node, pos) => {
        const attrs = { ...node.attrs, level: level || undefined };
        tr.setNodeMarkup(pos, level ? heading : paragraph, attrs);
      });
      dispatch(tr);
    }
    return true;
  };
}

export function currentLevel(state: EditorState): number {
  const n = state.selection.$from.parent;
  return n.type === heading ? (n.attrs.level as number) : 0;
}

/** Маркированный / нумерованный список: включить, переключить вид или убрать. */
export function toggleList(type: NodeType): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;
    const range = $from.blockRange($to);
    if (!range) return false;
    for (let d = $from.depth; d > 0; d--) {
      const n = $from.node(d);
      if (n.type === bullet_list || n.type === ordered_list) {
        if (n.type === type) {
          // Тот же вид — убираем список: поднимаем пункты, пока не выйдем из всех вложенных списков.
          const steps: Transaction[] = [];
          let cur = state;
          for (let i = 0; i < 12 && inList(cur); i++) {
            let t: Transaction | null = null;
            if (!liftListItem(list_item)(cur, (x) => { t = x; })) break;
            steps.push(t!);
            cur = cur.apply(t!);
          }
          if (!steps.length) return false;
          if (dispatch) for (const t of steps) dispatch(t);
          return true;
        }
        if (dispatch) {
          const pos = $from.before(d);
          dispatch(state.tr.setNodeMarkup(pos, type, type === ordered_list ? { order: 1, fmt: 'decimal' } : null));
        }
        return true;
      }
    }
    return wrapInList(type, type === ordered_list ? { order: 1, fmt: 'decimal' } : null)(state, dispatch);
  };
}

export function listKind(state: EditorState): 'bullet' | 'ordered' | null {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) {
    const t = $from.node(d).type;
    if (t === bullet_list) return 'bullet';
    if (t === ordered_list) return 'ordered';
  }
  return null;
}

/** Разрыв страницы: абзац после курсора начинается с новой страницы. */
export const insertPageBreak: Command = (state, dispatch) => {
  const { $from } = state.selection;
  if (!$from.parent.isTextblock) return false;
  if (dispatch) {
    let tr = state.tr;
    if (!state.selection.empty) tr = tr.deleteSelection();
    const pos = tr.mapping.map(state.selection.from);
    tr = tr.split(pos, 1, [{ type: paragraph, attrs: { ...$from.parent.attrs, level: undefined, pageBreakBefore: true } }]);
    dispatch(tr.scrollIntoView());
  }
  return true;
};

export function insertTable(rows: number, cols: number, borders = true): Command {
  return (state, dispatch) => {
    const cell = () => schema.nodes.table_cell.createAndFill()!;
    const table = schema.nodes.table.create({ borders }, Array.from({ length: rows }, () => schema.nodes.table_row.create(null, Array.from({ length: cols }, cell))));
    if (dispatch) {
      const at = state.selection.$from;
      // Таблица встаёт после текущего абзаца (в пустом абзаце — на его место), курсор — в первую ячейку.
      const empty = at.parent.isTextblock && at.parent.content.size === 0 && at.depth === 1;
      const pos = empty ? at.before() : at.after(at.depth > 0 ? 1 : 0);
      let tr = state.tr;
      if (empty) tr = tr.replaceWith(pos, at.after(), table);
      else tr = tr.insert(pos, table);
      // После таблицы всегда есть абзац, куда можно встать курсором.
      const end = pos + table.nodeSize;
      if (end >= tr.doc.content.size) tr = tr.insert(end, schema.nodes.paragraph.create());
      tr = tr.setSelection(Selection.near(tr.doc.resolve(pos + 1)));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

export function insertImage(src: string, width: number, height: number, alt = ''): Command {
  return (state, dispatch) => {
    if (dispatch) dispatch(state.tr.replaceSelectionWith(image.create({ src, width, height, alt })).scrollIntoView());
    return true;
  };
}

export function setLink(href: string): Command {
  return (state, dispatch) => {
    const type = schema.marks.link;
    if (!dispatch) return true;
    let tr = state.tr;
    if (state.selection.empty) {
      if (href) {
        const text = schema.text(href, [type.create({ href })]);
        tr = tr.replaceSelectionWith(text, false);
      }
    } else {
      for (const r of state.selection.ranges) {
        tr = tr.removeMark(r.$from.pos, r.$to.pos, type);
        if (href) tr = tr.addMark(r.$from.pos, r.$to.pos, type.create({ href }));
      }
    }
    dispatch(tr);
    return true;
  };
}

export function linkAt(state: EditorState): string | null {
  const m = schema.marks.link.isInSet(state.selection.$from.marks());
  return m ? (m.attrs.href as string) : null;
}

export function replaceMatch(view: EditorView, index: number, text: string): void {
  const s = findKey.getState(view.state);
  const m = s?.matches[index];
  if (!m) return;
  const tr = view.state.tr.insertText(text, m.from, m.to);
  view.dispatch(tr);
}

export function replaceAll(view: EditorView, text: string): number {
  const s = findKey.getState(view.state);
  if (!s?.matches.length) return 0;
  let tr = view.state.tr;
  for (let i = s.matches.length - 1; i >= 0; i--) tr = tr.insertText(text, s.matches[i].from, s.matches[i].to);
  view.dispatch(tr);
  return s.matches.length;
}

export function findStep(view: EditorView, dir: 1 | -1): void {
  const s = findKey.getState(view.state);
  if (!s?.matches.length) return;
  const sel = view.state.selection.from;
  let idx: number;
  if (dir > 0) {
    idx = s.matches.findIndex((m) => m.from >= sel + (s.current >= 0 && s.matches[s.current]?.from === sel ? 1 : 0));
    if (idx < 0) idx = 0;
  } else {
    idx = -1;
    for (let i = s.matches.length - 1; i >= 0; i--) if (s.matches[i].from < sel) { idx = i; break; }
    if (idx < 0) idx = s.matches.length - 1;
  }
  const m = s.matches[idx];
  const tr = view.state.tr.setSelection(TextSelection.create(view.state.doc, m.from, m.to)).setMeta(findKey, { current: idx }).scrollIntoView();
  view.dispatch(tr);
}

// ---------------------------------------------------------------- сборка

/** Вид таблицы: признак «с границами» хранится в узле, а столбцы рисует плагин изменения ширины. */
class BorderTableView extends TableView {
  constructor(node: PMNode, cellMinWidth: number) {
    super(node, cellMinWidth);
    this.table.className = node.attrs.borders ? 'bordered' : 'plain';
  }

  update(node: PMNode): boolean {
    const ok = super.update(node);
    if (ok) this.table.className = node.attrs.borders ? 'bordered' : 'plain';
    return ok;
  }
}

export function createState(doc: PMNode, extra: Plugin[] = []): EditorState {
  return EditorState.create({
    doc,
    plugins: [
      ...extra,
      history({ newGroupDelay: 600 }),
      keymap({
        Enter: enter,
        'Shift-Enter': softBreak,
        Backspace: chainCommands(backspace, joinBackward),
        Tab: tab,
        'Shift-Tab': shiftTab,
        Escape: selectParentNode,
        'Mod-Enter': insertPageBreak,
        'Mod-[': lift,
      }),
      keymap(baseKeymap),
      columnResizing({ cellMinWidth: 30, View: BorderTableView }),
      tableEditing(),
      dropCursor({ color: '#2b579a' }),
      gapCursor(),
      findPlugin,
    ],
  });
}

export { setBlockType };
