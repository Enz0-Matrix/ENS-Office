// SPDX-License-Identifier: AGPL-3.0-or-later
// Сравнение документов и исправления (рецензирование).
//
// Сравнение двух документов даёт третий: текст новой редакции, где вставленное помечено меткой `ins`,
// а удалённое оставлено на месте с меткой `del` (как «Сравнить» в Word). Блоки сопоставляются так же,
// как в слиянии (`diffBlocks` по ключам `mergeKey`), а внутри изменённых абзацев — по словам с учётом
// форматирования. Целиком вставленный или удалённый абзац помечается ещё и атрибутом `rev`, чтобы при
// принятии и отклонении он исчезал вместе со знаком абзаца, а не оставлял пустую строку.

import { Fragment, type Mark, type Node as PMNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { mergeKey } from './merge';
import type { Revision } from './schema';
import { diffBlocks } from './versions';

export interface RevInfo { author: string; date: string }
export interface CompareResult { doc: PMNode; inserted: number; deleted: number }

const REV = new Set(['ins', 'del']);
const children = (node: PMNode): PMNode[] => { const out: PMNode[] = []; node.forEach((n) => out.push(n)); return out; };
const plainMarks = (marks: readonly Mark[]) => marks.filter((m) => !REV.has(m.type.name));
/** Ключ блока без пометок исправлений: документ с исправлениями сравнивается по своему содержимому. */
const blockKey = (n: PMNode) => mergeKey(stripRevisionsJSON(n.toJSON()));
function stripRevisionsJSON(j: Record<string, unknown>): Record<string, unknown> {
  const out: Record<string, unknown> = { ...j };
  if (Array.isArray(j.marks)) {
    const marks = (j.marks as { type: string }[]).filter((m) => !REV.has(m.type));
    if (marks.length) out.marks = marks; else delete out.marks;
  }
  if (j.attrs && (j.attrs as Record<string, unknown>).rev) out.attrs = { ...(j.attrs as Record<string, unknown>), rev: null };
  if (Array.isArray(j.content)) out.content = (j.content as Record<string, unknown>[]).map(stripRevisionsJSON);
  return out;
}

/** Слова, пробелы и знаки по отдельности: «было/стало» показывается по словам, а не по буквам. */
const TOKEN = /[\p{L}\p{N}_’'-]+|\s+|[^\p{L}\p{N}_\s]/gu;
interface Token { key: string; node: PMNode }
function tokens(block: PMNode): Token[] {
  const out: Token[] = [];
  block.forEach((child) => {
    const mk = mergeKey(plainMarks(child.marks).map((m) => m.toJSON()));
    if (child.isText) {
      for (const m of child.text!.matchAll(TOKEN)) out.push({ key: m[0] + '\u0000' + mk, node: child.type.schema.text(m[0], child.marks) });
    } else out.push({ key: mergeKey(stripRevisionsJSON(child.toJSON())), node: child });
  });
  return out;
}

/** Наибольшая общая подпоследовательность по ключам: пары индексов совпавших элементов. */
function lcs(a: string[], b: string[], limit: number): [number, number][] | null {
  let pre = 0;
  while (pre < a.length && pre < b.length && a[pre] === b[pre]) pre++;
  let suf = 0;
  while (suf < a.length - pre && suf < b.length - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
  const n = a.length - pre - suf, m = b.length - pre - suf;
  if (n * m > limit) return null;
  const w = m + 1, t = new Uint32Array((n + 1) * w);
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) {
    t[i * w + j] = a[pre + i] === b[pre + j] ? t[(i + 1) * w + j + 1] + 1 : Math.max(t[(i + 1) * w + j], t[i * w + j + 1]);
  }
  const pairs: [number, number][] = [];
  for (let k = 0; k < pre; k++) pairs.push([k, k]);
  for (let i = 0, j = 0; i < n && j < m;) {
    if (a[pre + i] === b[pre + j]) { pairs.push([pre + i, pre + j]); i++; j++; }
    else if (t[i * w + j + 1] >= t[(i + 1) * w + j]) j++;
    else i++;
  }
  for (let k = suf; k > 0; k--) pairs.push([a.length - k, b.length - k]);
  return pairs;
}

/** Склеивает соседние текстовые узлы с одинаковыми метками. */
function joinInline(nodes: PMNode[]): PMNode[] {
  const out: PMNode[] = [];
  for (const n of nodes) {
    const last = out[out.length - 1];
    if (last && last.isText && n.isText && last.sameMarkup(n)) out[out.length - 1] = last.type.schema.text(last.text! + n.text!, last.marks);
    else out.push(n);
  }
  return out;
}

const wordsOf = (s: string) => new Set(s.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? []);
/** Похожесть двух блоков (0…1): доля общих слов. Пустые блоки одного типа считаются похожими. */
function similarity(a: PMNode, b: PMNode): number {
  if (a.type !== b.type && !(a.isTextblock && b.isTextblock)) return 0;
  const text = (n: PMNode) => n.textBetween(0, n.content.size, ' ', ' ');
  const x = wordsOf(text(a)), y = wordsOf(text(b));
  if (!x.size && !y.size) return a.type === b.type ? 1 : 0;
  let common = 0;
  for (const w of x) if (y.has(w)) common++;
  return (2 * common) / (x.size + y.size);
}

export function compareDocs(oldDoc: PMNode, newDoc: PMNode, info: RevInfo): CompareResult {
  const schema = newDoc.type.schema;
  let nextId = 1, inserted = 0, deleted = 0;
  const mark = (kind: 'ins' | 'del', id: number) => schema.marks[kind].create({ id, author: info.author, date: info.date });

  /** Весь блок вставлен или удалён: метка на каждом встроенном узле и атрибут rev на каждом абзаце. */
  const whole = (node: PMNode, kind: 'ins' | 'del'): PMNode => {
    const id = nextId++;
    if (kind === 'ins') inserted++; else deleted++;
    const m = mark(kind, id);
    const rev: Revision = { kind, id, author: info.author, date: info.date };
    const walk = (n: PMNode): PMNode => {
      if (n.isTextblock) {
        const inline = children(n).map((c) => c.mark(m.addToSet(plainMarks(c.marks))));
        return n.type.create({ ...n.attrs, rev }, inline, n.marks);
      }
      if (n.isLeaf) return n;
      return n.copy(Fragment.fromArray(children(n).map(walk)));
    };
    return walk(node);
  };

  const inlineDiff = (a: PMNode, b: PMNode): PMNode => {
    const ta = tokens(a), tb = tokens(b);
    const pairs = lcs(ta.map((t) => t.key), tb.map((t) => t.key), 4_000_000) ?? [];
    const out: PMNode[] = [];
    let i = 0, j = 0;
    const emit = (toI: number, toJ: number) => {
      if (i < toI) {
        const id = nextId++; deleted++;
        for (; i < toI; i++) out.push(ta[i].node.mark(mark('del', id).addToSet(plainMarks(ta[i].node.marks))));
      }
      if (j < toJ) {
        const id = nextId++; inserted++;
        for (; j < toJ; j++) out.push(tb[j].node.mark(mark('ins', id).addToSet(plainMarks(tb[j].node.marks))));
      }
    };
    for (const [pi, pj] of pairs) {
      emit(pi, pj);
      out.push(tb[pj].node.mark(plainMarks(tb[pj].node.marks)));
      i = pi + 1; j = pj + 1;
    }
    emit(ta.length, tb.length);
    return b.type.create({ ...b.attrs, rev: null }, joinInline(out), b.marks);
  };

  const pairBlocks = (a: PMNode, b: PMNode): PMNode[] => {
    if (a.isTextblock && b.isTextblock) return [inlineDiff(a, b)];
    // Списки, таблицы, строки и ячейки с одинаковой оболочкой сравниваются по вложенным блокам.
    // Строки с разным числом ячеек не смешиваем: сетка таблицы иначе развалится.
    if (!a.isLeaf && a.type === b.type && !(a.type.name === 'table_row' && a.childCount !== b.childCount) &&
      !a.type.spec.atom && a.type.name !== 'frame') {
      const copy = b.copy(Fragment.fromArray(sequence(children(a), children(b))));
      try { copy.check(); return [copy]; } catch { /* оболочка не принимает такой состав — правка целиком */ }
    }
    return [whole(a, 'del'), whole(b, 'ins')];
  };

  const sequence = (a: PMNode[], b: PMNode[]): PMNode[] => {
    const out: PMNode[] = [];
    for (const h of diffBlocks(a.map(blockKey), b.map(blockKey))) {
      const xs = a.slice(...h.a), ys = b.slice(...h.b);
      if (h.type === 'same') { out.push(...ys); continue; }
      // Сопоставляем изменённые блоки по похожести (как «Сравнить» в Word): правка абзаца — правка, а не «удалён и вставлен».
      const matched: [number, number][] = [];
      if (xs.length * ys.length <= 250_000) {
        // Динамика по «похожим» парам: сохраняет порядок и берёт как можно больше пар.
        const n = xs.length, m = ys.length, w = m + 1, t = new Float64Array((n + 1) * w);
        const sim = (i: number, j: number) => similarity(xs[i], ys[j]);
        const ok = (i: number, j: number) => {
          if (xs[i].type.name === 'table_row' && xs[i].childCount !== ys[j].childCount) return 0;
          const s = sim(i, j);
          return s >= 0.4 || (n === 1 && m === 1 && xs[i].type === ys[j].type) ? 1 + s : 0;
        };
        const score = new Float64Array(n * m);
        for (let i = 0; i < n; i++) for (let j = 0; j < m; j++) score[i * m + j] = ok(i, j);
        for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) {
          const s = score[i * m + j];
          t[i * w + j] = Math.max(t[(i + 1) * w + j], t[i * w + j + 1], s ? t[(i + 1) * w + j + 1] + s : 0);
        }
        for (let i = 0, j = 0; i < n && j < m;) {
          const s = score[i * m + j];
          if (s && t[i * w + j] === t[(i + 1) * w + j + 1] + s) { matched.push([i, j]); i++; j++; }
          else if (t[i * w + j] === t[(i + 1) * w + j]) i++;
          else j++;
        }
      }
      let i = 0, j = 0;
      for (const [pi, pj] of [...matched, [xs.length, ys.length] as [number, number]]) {
        for (; i < pi; i++) out.push(whole(xs[i], 'del'));
        for (; j < pj; j++) out.push(whole(ys[j], 'ins'));
        if (pi < xs.length) { out.push(...pairBlocks(xs[pi], ys[pj])); i++; j++; }
      }
    }
    return out;
  };

  const body = sequence(children(oldDoc), children(newDoc));
  const doc = newDoc.copy(Fragment.fromArray(body.length ? body : [schema.nodes.paragraph.create()]));
  doc.check();
  return { doc, inserted, deleted };
}

// --- Исправления в документе

export interface RevRange { id: number; kind: 'ins' | 'del'; from: number; to: number; author: string; date: string }

/** Все исправления по порядку: одно на id, диапазон — от первой до последней его части. */
export function listRevisions(doc: PMNode): RevRange[] {
  const map = new Map<number, RevRange>();
  const add = (id: number, kind: 'ins' | 'del', from: number, to: number, author: string, date: string) => {
    const r = map.get(id);
    if (r) { r.from = Math.min(r.from, from); r.to = Math.max(r.to, to); } else map.set(id, { id, kind, from, to, author, date });
  };
  doc.descendants((node, pos) => {
    const rev = node.attrs.rev as Revision | null | undefined;
    if (rev) add(rev.id, rev.kind, pos, pos + node.nodeSize, rev.author, rev.date);
    if (node.isInline) for (const m of node.marks) if (REV.has(m.type.name)) {
      add(m.attrs.id as number, m.type.name as 'ins' | 'del', pos, pos + node.nodeSize, m.attrs.author as string, m.attrs.date as string);
    }
    return true;
  });
  return [...map.values()].sort((a, b) => a.from - b.from || a.id - b.id);
}

export const hasRevisions = (doc: PMNode) => listRevisions(doc).length > 0;

/**
 * Принять или отклонить исправления из `ids`. Принятая вставка и отклонённое удаление остаются текстом
 * без пометки; принятое удаление и отклонённая вставка исчезают вместе с целиком помеченными абзацами,
 * строками и таблицами. Ячейка таблицы не исчезает: в ней остаётся пустой абзац.
 */
export function resolveRevisions(state: EditorState, ids: Set<number> | 'all', accept: boolean): Transaction | null {
  const doc = state.doc;
  const hit = (id: number) => ids === 'all' || ids.has(id);
  const removes = (kind: 'ins' | 'del') => (accept ? kind === 'del' : kind === 'ins');
  const removedBlock = (n: PMNode) => {
    const rev = n.attrs.rev as Revision | null | undefined;
    return !!rev && hit(rev.id) && removes(rev.kind);
  };
  /** Все абзацы внутри узла уходят — узел уходит целиком. */
  const allGone = (n: PMNode) => {
    let any = false, all = true;
    n.descendants((d) => {
      if (!all) return false;
      if (d.isTextblock) { any = true; if (!removedBlock(d)) all = false; return false; }
      return true;
    });
    return any && all;
  };
  const CELL = new Set(['table_cell', 'table_header']);
  const rebuild = (n: PMNode): PMNode[] => {
    if (n.isTextblock) {
      if (removedBlock(n)) return [];
      const inline: PMNode[] = [];
      n.forEach((c) => {
        const m = c.marks.find((x) => REV.has(x.type.name) && hit(x.attrs.id as number));
        if (!m) inline.push(c);
        else if (!removes(m.type.name as 'ins' | 'del')) inline.push(c.mark(m.removeFromSet(c.marks)));
      });
      const rev = n.attrs.rev as Revision | null;
      return [n.type.create(rev && hit(rev.id) ? { ...n.attrs, rev: null } : n.attrs, joinInline(inline), n.marks)];
    }
    if (n.isLeaf) return [n];
    if (!CELL.has(n.type.name) && allGone(n)) return [];
    const kids = children(n).flatMap(rebuild);
    if (!kids.length) {
      const filled = n.type.createAndFill(n.attrs, [], n.marks);
      return filled ? [filled] : [];
    }
    try {
      const copy = n.copy(Fragment.fromArray(kids));
      copy.check();
      return [copy];
    } catch {
      return [n.type.createAndFill(n.attrs, Fragment.fromArray(kids), n.marks) ?? n];
    }
  };
  // Перестраиваем только верхние блоки, где есть затронутые исправления.
  const tr = state.tr;
  const tops: [number, PMNode][] = [];
  doc.forEach((node, offset) => {
    let touched = false;
    node.descendants((d) => {
      if (touched) return false;
      const rev = d.attrs.rev as Revision | null | undefined;
      if (rev && hit(rev.id)) touched = true;
      else if (d.isInline && d.marks.some((m) => REV.has(m.type.name) && hit(m.attrs.id as number))) touched = true;
      return !touched;
    });
    if (!touched && (node.attrs.rev as Revision | null | undefined) && hit((node.attrs.rev as Revision).id)) touched = true;
    if (touched) tops.push([offset, node]);
  });
  if (!tops.length) return null;
  for (const [offset, node] of tops.reverse()) {
    const from = tr.mapping.map(offset);
    tr.replaceWith(from, from + node.nodeSize, rebuild(node));
  }
  if (!tr.doc.childCount) tr.insert(0, state.schema.nodes.paragraph.create());
  return tr;
}
