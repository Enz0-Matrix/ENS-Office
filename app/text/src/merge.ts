// Офлайн-слияние трёх копий. Сравнение блоков общее с историей версий.
import { Fragment, type Node as PMNode } from 'prosemirror-model';
import { diffBlocks } from './versions';

export type MergeChoice = 'base' | 'mine' | 'theirs';
export type MergePart =
  | { kind: 'nodes'; nodes: PMNode[] }
  | { kind: 'container'; node: PMNode; children: MergePart[] }
  | { kind: 'conflict'; id: number; base: PMNode[]; mine: PMNode[]; theirs: PMNode[] };
export interface MergePlan { doc: PMNode; parts: MergePart[]; conflicts: number }
interface Change { from: number; to: number; nodes: PMNode[]; side: 'mine' | 'theirs' }

/** Контрольная сумма байтов (двойной хэш и длина) для сравнения бинарных данных без копирования. */
function bytesKey(bytes: Uint8Array): string {
  let h1 = 0x811c9dc5, h2 = 5381;
  for (let i = 0; i < bytes.length; i++) {
    h1 = Math.imul(h1 ^ bytes[i], 0x01000193);
    h2 = (Math.imul(h2, 33) + bytes[i]) | 0;
  }
  return `bytes:${bytes.length}:${(h1 >>> 0).toString(16)}:${(h2 >>> 0).toString(16)}`;
}
const byKey = (a: [string, string], b: [string, string]) => (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0);

/**
 * Каноническая строка для сравнения значений. Порядок ключей, `undefined` вместо отсутствующего свойства
 * и порядок записей Map/Set не должны превращать одинаковые данные в конфликт; Map, Set, Date
 * и бинарные данные сравниваются по содержимому, а не как пустые объекты.
 */
export function mergeKey(value: unknown): string {
  if (value === null || value === undefined) return 'null';
  if (typeof value === 'bigint') return value + 'n';
  if (Array.isArray(value)) return '[' + value.map(mergeKey).join(',') + ']';
  if (value instanceof ArrayBuffer) return bytesKey(new Uint8Array(value));
  if (ArrayBuffer.isView(value)) return bytesKey(new Uint8Array(value.buffer, value.byteOffset, value.byteLength));
  if (value instanceof Date) return 'date:' + value.getTime();
  if (value instanceof Map) return 'map{' + [...value].map(([k, v]): [string, string] => [mergeKey(k), mergeKey(v)])
    .sort(byKey).map(([k, v]) => k + '=>' + v).join(',') + '}';
  if (value instanceof Set) return 'set[' + [...value].map(mergeKey).sort().join(',') + ']';
  if (typeof value === 'object') return '{' + Object.entries(value)
    .filter(([, v]) => v !== undefined).sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
    .map(([k, v]) => JSON.stringify(k) + ':' + mergeKey(v)).join(',') + '}';
  return JSON.stringify(value) ?? 'null';
}

/** Текст для окна выбора: бинарные данные не раздувают просмотр, длинный вывод обрезается. */
export function describeMeta(value: unknown, limit = 20000): string {
  const text = JSON.stringify(value, (_key, v: unknown) => {
    if (v instanceof ArrayBuffer || ArrayBuffer.isView(v)) return `[бинарные данные: ${v.byteLength} байт]`;
    if (v instanceof Map) return Object.fromEntries(v);
    if (v instanceof Set) return [...v];
    return v;
  }, 2) ?? 'null';
  return text.length > limit ? text.slice(0, limit) + '\n… (обрезано)' : text;
}

/** Доля блоков основы, найденных в копии. Низкое значение — вероятно, выбрана не та основа. */
export function baseOverlap(base: PMNode, other: PMNode): number {
  if (!base.childCount) return 1;
  const left = new Map<string, number>();
  base.forEach((n) => { const k = mergeKey(n.toJSON()); left.set(k, (left.get(k) ?? 0) + 1); });
  let kept = 0;
  other.forEach((n) => {
    const k = mergeKey(n.toJSON()), count = left.get(k) ?? 0;
    if (count > 0) { kept++; left.set(k, count - 1); }
  });
  return kept / base.childCount;
}
const children = (node: PMNode): PMNode[] => { const out: PMNode[] = []; node.forEach((n) => out.push(n)); return out; };
const same = (a: PMNode[], b: PMNode[]) => a.length === b.length && a.every((n, i) => n.eq(b[i]));

export function planMerge(base: PMNode, mine: PMNode, theirs: PMNode): MergePlan {
  let conflicts = 0;
  const sequence = (original: PMNode[], local: PMNode[], remote: PMNode[]): MergePart[] => {
    if (same(local, remote)) return [{ kind: 'nodes', nodes: local }];
    if (same(original, local)) return [{ kind: 'nodes', nodes: remote }];
    if (same(original, remote)) return [{ kind: 'nodes', nodes: local }];
    const keys = (nodes: PMNode[]) => nodes.map((n) => mergeKey(n.toJSON()));
    const baseKeys = keys(original);
    const changes = (nodes: PMNode[], side: Change['side']): Change[] => diffBlocks(baseKeys, keys(nodes))
      .filter((h) => h.type === 'change').map((h) => ({ from: h.a[0], to: h.a[1], nodes: nodes.slice(...h.b), side }));
    const all = [...changes(local, 'mine'), ...changes(remote, 'theirs')]
      .sort((a, b) => a.from - b.from || a.to - b.to);
    const overlap = (a: Change, b: Change) => {
      // Вставка на границе замены неоднозначна: не угадываем её порядок.
      if (a.from === a.to) return a.from >= b.from && a.from <= b.to;
      if (b.from === b.to) return b.from >= a.from && b.from <= a.to;
      return a.from < b.to && b.from < a.to;
    };
    const out: MergePart[] = [];
    let cursor = 0;
    for (let i = 0; i < all.length;) {
      const group = [all[i++]];
      while (i < all.length && group.some((c) => overlap(c, all[i]))) group.push(all[i++]);
      const from = Math.min(...group.map((c) => c.from));
      const to = Math.max(...group.map((c) => c.to));
      if (from > cursor) out.push({ kind: 'nodes', nodes: original.slice(cursor, from) });
      const old = original.slice(from, to);
      const apply = (side: Change['side']) => {
        const nodes: PMNode[] = [];
        let at = from;
        for (const c of group.filter((c) => c.side === side)) {
          nodes.push(...original.slice(at, c.from), ...c.nodes); at = c.to;
        }
        nodes.push(...original.slice(at, to));
        return nodes;
      };
      const a = apply('mine'), b = apply('theirs');
      if (same(a, b)) out.push({ kind: 'nodes', nodes: a });
      else if (same(old, a)) out.push({ kind: 'nodes', nodes: b });
      else if (same(old, b)) out.push({ kind: 'nodes', nodes: a });
      else if (old.length === 1 && a.length === 1 && b.length === 1 && !old[0].isTextblock && !old[0].isLeaf &&
        old[0].sameMarkup(a[0]) && old[0].sameMarkup(b[0])) {
        // Внутри таблиц и списков сравниваем вложенные абзацы, сохраняя оболочку.
        out.push({ kind: 'container', node: old[0], children: sequence(children(old[0]), children(a[0]), children(b[0])) });
      } else out.push({ kind: 'conflict', id: conflicts++, base: old, mine: a, theirs: b });
      cursor = to;
    }
    if (cursor < original.length) out.push({ kind: 'nodes', nodes: original.slice(cursor) });
    return out;
  };
  const parts = sequence(children(base), children(mine), children(theirs));
  return { doc: mine, parts, conflicts };
}

/** Неразрешённый конфликт никогда не получает неявного варианта по умолчанию. */
export function resolveMerge(plan: MergePlan, choices: ReadonlyMap<number, MergeChoice>): PMNode {
  const collect = (parts: MergePart[]): PMNode[] => parts.flatMap((p): PMNode[] => {
    if (p.kind === 'nodes') return p.nodes;
    if (p.kind === 'container') return [p.node.copy(Fragment.fromArray(collect(p.children)))];
    const choice = choices.get(p.id);
    if (!choice) throw new Error('Сначала выберите вариант каждого конфликта.');
    return p[choice];
  });
  const nodes = collect(plan.parts);
  const doc = plan.doc.copy(Fragment.fromArray(nodes.length ? nodes : [plan.doc.type.schema.nodes.paragraph.create()]));
  doc.check();
  return doc;
}
