// SPDX-License-Identifier: AGPL-3.0-or-later
// Сноски и концевые сноски (Ссылки → Вставить сноску, Ctrl+Alt+F; концевая — Ctrl+Alt+D, как в Word).
//
// Сноска — узел в тексте; её текст (абзацы) хранится в самом узле, поэтому сноска переезжает вместе с текстом
// при копировании, вырезании и отмене. Номера считаются по порядку при каждой правке. Лучше, чем в Word:
// текст сноски виден всплывающей подсказкой при наведении на номер, без прокрутки к низу страницы.

import { baseKeymap } from 'prosemirror-commands';
import { history, redo, undo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { DOMSerializer, type Node as PMNode } from 'prosemirror-model';
import { EditorState, Plugin, TextSelection, type Transaction } from 'prosemirror-state';
import { Decoration, DecorationSet, EditorView } from 'prosemirror-view';
import { schema } from './schema';
import { dialog, el } from './ui';

export interface NoteInfo { pos: number; n: number; label: string; end: boolean; text: string }

const ROMAN: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
const roman = (n: number) => { let s = ''; for (const [v, r] of ROMAN) while (n >= v) { s += r; n -= v; } return s; };

/** Текст сноски одной строкой (для подсказки и списка). */
export function noteText(content: unknown): string {
  try {
    const nodes = (content as Record<string, unknown>[]).map((j) => schema.nodeFromJSON(j));
    return nodes.map((n) => n.textBetween(0, n.content.size, ' ', ' ')).join('\n').trim();
  } catch { return ''; }
}

export function listNotes(doc: PMNode): NoteInfo[] {
  let foot = 0, end = 0;
  const out: NoteInfo[] = [];
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.footnote) return true;
    const isEnd = !!node.attrs.end;
    const n = isEnd ? ++end : ++foot;
    out.push({ pos, n, end: isEnd, label: isEnd ? roman(n) : String(n), text: noteText(node.attrs.content) });
    return false;
  });
  return out;
}

/** Сноска внизу страницы: номер надстрочным и текст с форматированием (тот же элемент и для измерения высоты). */
export function renderNote(label: string, content: unknown): HTMLElement {
  const box = document.createElement('div');
  box.className = 'fn-note';
  let nodes: PMNode[] = [];
  try { nodes = (content as Record<string, unknown>[]).map((j) => schema.nodeFromJSON(j)); } catch { nodes = []; }
  if (!nodes.length) nodes = [schema.nodes.paragraph.create()];
  const frag = DOMSerializer.fromSchema(schema).serializeFragment(schema.nodes.doc.create(null, nodes).content);
  box.append(frag);
  const sup = document.createElement('sup');
  sup.className = 'fn-label';
  sup.textContent = label;
  const first = box.firstElementChild;
  if (first) first.prepend(sup, ' '); else box.append(sup);
  return box;
}

function build(doc: PMNode): DecorationSet {
  const notes = listNotes(doc);
  if (!notes.length) return DecorationSet.empty;
  return DecorationSet.create(doc, notes.map((f) => Decoration.node(f.pos, f.pos + 1, { 'data-n': f.label, 'data-tip': f.text || '(пустая сноска — дважды щёлкните, чтобы ввести текст)' })));
}

let editNote: (view: EditorView, pos: number) => void = () => {};

export const footnotesPlugin: Plugin<DecorationSet> = new Plugin<DecorationSet>({
  state: {
    init: (_c, state) => build(state.doc),
    apply: (tr, prev, _o, state) => (tr.docChanged ? build(state.doc) : prev),
  },
  props: {
    decorations(state) { return footnotesPlugin.getState(state); },
    handleDoubleClickOn(view, _pos, node, nodePos) {
      if (node.type !== schema.nodes.footnote) return false;
      editNote(view, nodePos);
      return true;
    },
  },
});

/** Вставить сноску у курсора и сразу открыть окно для её текста. */
export function insertNote(view: EditorView, end = false): void {
  const { state } = view;
  const at = state.selection.to;
  const tr: Transaction = state.tr.insert(at, schema.nodes.footnote.create({ content: [], end }));
  tr.setSelection(TextSelection.create(tr.doc, at + 1));
  view.dispatch(tr);
  noteDialog(view, at, true);
}

/** Окно текста сноски: маленький редактор с форматированием (полужирный, курсив — Ctrl+B / Ctrl+I). */
export function noteDialog(view: EditorView, pos: number, fresh = false): void {
  const node = view.state.doc.nodeAt(pos);
  if (!node || node.type !== schema.nodes.footnote) return;
  dialog((box, close) => {
    box.classList.add('note-dialog');
    const blocks = (node.attrs.content as Record<string, unknown>[]);
    let doc: PMNode;
    try { doc = schema.nodes.doc.create(null, blocks.length ? blocks.map((j) => schema.nodeFromJSON(j)) : [schema.nodes.paragraph.create()]); }
    catch { doc = schema.nodes.doc.create(null, [schema.nodes.paragraph.create()]); }
    const toggle = (mark: 'strong' | 'em') => (s: EditorState, d?: (tr: Transaction) => void) => {
      const { from, to, empty } = s.selection;
      if (empty) return false;
      const t = schema.marks[mark];
      d?.(s.doc.rangeHasMark(from, to, t) ? s.tr.removeMark(from, to, t) : s.tr.addMark(from, to, t.create()));
      return true;
    };
    const host = el('div', { class: 'note-editor' });
    const mini = new EditorView(host, {
      state: EditorState.create({ doc, plugins: [history(), keymap({ 'Mod-z': undo, 'Mod-y': redo, 'Mod-Shift-z': redo, 'Mod-b': toggle('strong'), 'Mod-i': toggle('em') }), keymap(baseKeymap)] }),
      attributes: { 'aria-label': 'Текст сноски', spellcheck: 'true', lang: 'ru' },
    });
    const label = listNotes(view.state.doc).find((n) => n.pos === pos);
    const save = () => {
      const cur = view.state.doc.nodeAt(pos);
      if (cur?.type === schema.nodes.footnote) {
        const content: Record<string, unknown>[] = [];
        mini.state.doc.forEach((b) => content.push(b.toJSON() as Record<string, unknown>));
        // Курсор — сразу за номером сноски: дальше печатается основной текст.
        const tr = view.state.tr.setNodeMarkup(pos, undefined, { ...cur.attrs, content });
        view.dispatch(tr.setSelection(TextSelection.create(tr.doc, pos + 1)));
      }
      mini.destroy();
      close();
      view.focus();
    };
    const remove = el('button', { class: 'danger', textContent: 'Удалить сноску', onclick: () => {
      mini.destroy();
      close();
      const cur = view.state.doc.nodeAt(pos);
      if (cur?.type === schema.nodes.footnote) view.dispatch(view.state.tr.delete(pos, pos + 1));
      view.focus();
    } });
    box.append(el('h2', { textContent: `${node.attrs.end ? 'Концевая сноска' : 'Сноска'} ${label?.label ?? ''}` }), host,
      el('p', { class: 'hint', textContent: 'Номер ставится сам. Ctrl+Enter — готово.' }),
      el('div', { class: 'dlg-actions' }, remove, el('button', { textContent: fresh ? 'Без текста' : 'Отмена', onclick: () => { mini.destroy(); close(); view.focus(); } }),
        el('button', { class: 'primary', textContent: 'Готово', onclick: save })));
    host.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); save(); } });
    // Фокус — в редактор сноски, курсор в конец текста (окно не переводит фокус на кнопку: data-autofocus).
    mini.dispatch(mini.state.tr.setSelection(TextSelection.atEnd(mini.state.doc)));
    (mini.dom as HTMLElement).dataset.autofocus = '1';
    (mini.dom as HTMLElement & { ensFocus?: () => void }).ensFocus = () => mini.focus();
  }, false);
}
editNote = (view, pos) => noteDialog(view, pos);

/** К следующей (или предыдущей) сноске. */
export function gotoNote(view: EditorView, dir: 1 | -1): boolean {
  const notes = listNotes(view.state.doc);
  if (!notes.length) return false;
  const at = view.state.selection.from;
  const t = dir > 0 ? notes.find((n) => n.pos >= at) ?? notes[0] : [...notes].reverse().find((n) => n.pos + 1 < at) ?? notes[notes.length - 1];
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, t.pos + 1)).scrollIntoView());
  view.focus();
  return true;
}
