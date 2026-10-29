// Команды работы с нумерацией: схема многоуровневого списка, «начать с…», «продолжить», преобразование номеров,
// набранных вручную («1.2.14.Б Текст»), в настоящий список.

import { wrapInList } from 'prosemirror-schema-list';
import type { Node as PMNode } from 'prosemirror-model';
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import { type Level, type Scheme } from './numbering';
import { schema } from './schema';

const { ordered_list, bullet_list } = schema.nodes;
const isList = (n: PMNode) => n.type === ordered_list || n.type === bullet_list;

/** Схема всего списка хранится у внешнего списка: он — первый список среди предков выделения. */
export function applyScheme(scheme: Scheme | null): Command {
  return (state, dispatch) => {
    const { $from } = state.selection;
    for (let d = 1; d <= $from.depth; d++) {
      const n = $from.node(d);
      if (!isList(n)) continue;
      if (dispatch) {
        dispatch(state.tr.setNodeMarkup($from.before(d), ordered_list, {
          order: n.type === ordered_list ? n.attrs.order : 1, fmt: n.type === ordered_list ? n.attrs.fmt : 'decimal', scheme,
        }));
      }
      return true;
    }
    return wrapInList(ordered_list, { order: 1, fmt: 'decimal', scheme })(state, dispatch);
  };
}

/** Ближайший нумерованный список вокруг выделения (самый вложенный). */
function innerOrdered(state: EditorState): { depth: number; node: PMNode } | null {
  const { $from } = state.selection;
  for (let d = $from.depth; d > 0; d--) if ($from.node(d).type === ordered_list) return { depth: d, node: $from.node(d) };
  return null;
}

export function setListStart(n: number): Command {
  return (state, dispatch) => {
    const l = innerOrdered(state);
    if (!l) return false;
    if (dispatch) dispatch(state.tr.setNodeMarkup(state.selection.$from.before(l.depth), undefined, { ...l.node.attrs, order: Math.max(1, Math.floor(n)) }));
    return true;
  };
}

export function currentStart(state: EditorState): number | null {
  const l = innerOrdered(state);
  return l ? (l.node.attrs.order as number) : null;
}

/** «Продолжить нумерацию»: начать с номера, следующего за ближайшим предыдущим списком (и взять его схему). */
export const continueList: Command = (state, dispatch) => {
  const l = innerOrdered(state);
  if (!l) return false;
  const { $from } = state.selection;
  const parent = $from.node(l.depth - 1);
  const idx = $from.index(l.depth - 1);
  let prev: PMNode | null = null;
  for (let j = idx - 1; j >= 0; j--) if (parent.child(j).type === ordered_list) { prev = parent.child(j); break; }
  if (!prev) return false;
  const prevStart = prev.attrs.order !== 1 ? (prev.attrs.order as number) : ((prev.attrs.scheme as Scheme | null)?.[0]?.start ?? 1);
  if (dispatch) {
    dispatch(state.tr.setNodeMarkup($from.before(l.depth), undefined, {
      ...l.node.attrs, order: prevStart + prev.childCount, scheme: l.node.attrs.scheme ?? prev.attrs.scheme, fmt: prev.attrs.fmt,
    }));
  }
  return true;
};

/** Номер, с которого начинается список, по первому пункту: «14» из «1.2.14», «Б» → 2. */
function typedStart(p: Parsed): number {
  if (p.kind === 'N' || p.kind === 'P') return Math.max(1, p.parts[p.parts.length - 1]);
  const ch = (p.last ?? '').toLowerCase();
  const ru = 'абвгдежзиклмнопрстуфхцчшщэюя'.indexOf(ch);
  const la = 'abcdefghijklmnopqrstuvwxyz'.indexOf(ch);
  return Math.max(1, (ru >= 0 ? ru : la) + 1);
}

// --- Номера, набранные вручную

interface Parsed { sig: string; kind: 'N' | 'A' | 'L' | 'P'; parts: number[]; prefix: number; last?: string; /** После номера стоит точка («1.1.»). */ dot?: boolean }

/** Распознать начало абзаца: «1.», «1.2.», «1.2.14 », «1.2.14.Б », «а) », «1) ». */
export function parseManual(text: string): Parsed | null {
  let m = /^(\d+(?:\.\d+)*)\.([А-ЯЁA-Z])[.)]?\s+\S/.exec(text);
  if (m) return { sig: `A${m[1].split('.').length}`, kind: 'A', parts: m[1].split('.').map(Number), prefix: m[0].length - 1, last: m[2], dot: true };
  m = /^(\d+(?:\.\d+)+)\.?\s+\S/.exec(text);
  if (m) return { sig: `N${m[1].split('.').length}`, kind: 'N', parts: m[1].split('.').map(Number), prefix: m[0].length - 1, dot: /\.\s/.test(m[0].slice(m[1].length - 1, m[1].length + 1)) || m[0].slice(m[1].length).startsWith('.') };
  m = /^(\d+)\.\s+\S/.exec(text);
  if (m) return { sig: 'N1', kind: 'N', parts: [Number(m[1])], prefix: m[0].length - 1, dot: true };
  m = /^([а-яёa-z])\)\s+\S/.exec(text);
  if (m) return { sig: 'L', kind: 'L', parts: [], prefix: m[0].length - 1, last: m[1] };
  m = /^(\d+)\)\s+\S/.exec(text);
  if (m) return { sig: 'P', kind: 'P', parts: [Number(m[1])], prefix: m[0].length - 1 };
  return null;
}

/**
 * Находит в документе подряд идущие абзацы с ручными номерами и заменяет их настоящими вложенными списками.
 * Группа — не меньше двух пунктов (или один многоуровневый «1.1»), чтобы не тронуть «2026 год…».
 * Возвращает число преобразованных абзацев.
 */
export function convertManual(state: EditorState, dispatch?: (tr: Transaction) => void): number {
  const doc = state.doc;
  const { from, to, empty } = state.selection;
  const children: { node: PMNode; pos: number }[] = [];
  doc.forEach((node, pos) => children.push({ node, pos }));
  const inRange = (c: { node: PMNode; pos: number }) => empty || (c.pos + c.node.nodeSize > from && c.pos < to);

  const groups: { start: number; end: number; items: { node: PMNode; p: Parsed }[] }[] = [];
  let cur: { start: number; end: number; items: { node: PMNode; p: Parsed }[] } | null = null;
  for (const c of children) {
    const p = c.node.type === schema.nodes.paragraph && inRange(c) ? parseManual(c.node.textContent) : null;
    if (p) {
      cur ??= { start: c.pos, end: c.pos, items: [] };
      cur.items.push({ node: c.node, p });
      cur.end = c.pos + c.node.nodeSize;
    } else if (cur) {
      groups.push(cur);
      cur = null;
    }
  }
  if (cur) groups.push(cur);
  const valid = groups.filter((g) => g.items.length >= 2 || (g.items[0]?.p.parts.length ?? 0) >= 2);
  if (!valid.length) return 0;

  const tr = state.tr;
  let converted = 0;
  for (const g of [...valid].reverse()) {
    // Уровни: «1.1» — по числу частей, «1.1.1.Б» — на уровень глубже «1.1.1», «а)» и «1)» — внутри последнего «1.1».
    let baseLevel = 0;
    let lastKind = '';
    const items = g.items.map(({ node, p }) => {
      let level: number;
      if (p.kind === 'N') { level = p.parts.length - 1; baseLevel = level; } else if (p.kind === 'A') { level = p.parts.length; baseLevel = level; } else if (p.kind === 'L') level = baseLevel + 1;
      else level = baseLevel + (lastKind === 'L' ? 2 : 1);
      lastKind = p.kind;
      return { node, p, level };
    });
    // Схема — из того, что набрано: на каждом уровне шаблон первого встреченного номера.
    const spec: Level[] = [];
    const tokens = (n: number) => Array.from({ length: n }, (_, k) => `%${k + 1}`).join('.');
    for (const { p, level } of items) {
      if (spec[level]) continue;
      const hang = 28 + Math.min(level, 4) * 8;
      if (p.kind === 'N') spec[level] = { fmt: 'decimal', text: tokens(p.parts.length) + (p.dot ? '.' : ''), start: 1, hang };
      else if (p.kind === 'A') spec[level] = { fmt: /[А-ЯЁ]/.test(p.last ?? '') ? 'russianUpper' : 'upperLetter', text: tokens(level + 1), start: 1, hang };
      else if (p.kind === 'L') spec[level] = { fmt: /[а-яё]/.test(p.last ?? '') ? 'russianLower' : 'lowerLetter', text: `%${level + 1})`, start: 1, hang };
      else spec[level] = { fmt: 'decimal', text: `%${level + 1})`, start: 1, hang };
    }
    const simpleOnly = items.every((i) => i.p.kind === 'N' && i.p.parts.length === 1 && i.p.dot);
    const scheme: Scheme | null = simpleOnly && items.every((i) => i.level === 0) ? null
      : Array.from({ length: 9 }, (_, i) => spec[i] ?? { fmt: 'decimal', text: `%${i + 1}.`, start: 1, hang: 36 });

    // Сборка вложенных списков.
    interface Open { level: number; items: PMNode[][]; start: number }
    const open: Open[] = [];
    const out: PMNode[] = [];
    const build = (o: Open): PMNode => ordered_list.create({ order: o.start, fmt: 'decimal', scheme: null }, o.items.map((c) => schema.nodes.list_item.create(null, c)));
    const close = () => {
      const o = open.pop()!;
      const node = build(o);
      const parent = open[open.length - 1];
      if (parent) {
        if (!parent.items.length) parent.items.push([schema.nodes.paragraph.create()]);
        parent.items[parent.items.length - 1].push(node);
      } else out.push(node);
    };
    let prevLevel = -1;
    for (const it of items) {
      const level = Math.min(it.level, prevLevel + 1);
      while (open.length && open[open.length - 1].level > level) close();
      if (!open.length || open[open.length - 1].level < level) open.push({ level, items: [], start: typedStart(it.p) });
      open[open.length - 1].items.push([it.node.copy(it.node.content.cut(it.p.prefix))]);
      prevLevel = level;
    }
    while (open.length) close();
    // Схема и начальный номер — у внешнего списка.
    const root = out[0];
    out[0] = ordered_list.create({ ...root.attrs, scheme }, root.content);
    tr.replaceWith(g.start, g.end, out);
    converted += items.length;
  }
  if (dispatch) dispatch(tr);
  return converted;
}
