// История версий и журнал изменений документа (как журнал изменений в ENS-Table, но с возвратом к версии).
//
// Версия — это список хэшей верхнеуровневых блоков (абзацы, списки, таблицы) плюс параметры страницы.
// Сами блоки лежат в общем хранилище и не повторяются: сотая версия документа весит столько, сколько в ней
// изменилось. Картинки тоже в общем хранилище, по одной копии. Вся история сжимается и вкладывается в сам docx
// (customXml, Word такие части сохраняет), поэтому путешествует вместе с файлом.

import { deflateSync, inflateSync, strFromU8, strToU8 } from 'fflate';
import type { Node as PMNode } from 'prosemirror-model';
import type { DocMeta } from './docx';
import { schema } from './schema';

export type VersionKind = 'save' | 'auto' | 'named' | 'open' | 'before' | 'outside';

export interface Version {
  id: string;
  t: number;
  u: string;
  kind: VersionKind;
  label: string;
  /** Хэш содержимого и страницы целиком. */
  hash: string;
  blocks: string[];
  meta: { page: DocMeta['page']; defaults: DocMeta['defaults']; hf: DocMeta['hf'] };
  words: number;
}

export interface LogEntry {
  /** Кто. */
  u: string;
  t: number;
  /** Чем: «ENS-Tekst 0.1.0» или «другая программа». */
  app: string;
  kind: 'save' | 'outside';
  /** Что изменилось по сравнению с предыдущей записью. */
  ch: { added: number; removed: number; changed: number; words: number; notes: string[] };
  /** Хэш содержимого файла на момент записи (по нему замечаем правку в другой программе). */
  h: string;
}

export interface History {
  versions: Version[];
  /** Хэш блока → JSON блока (картинки заменены ссылками ens-img:хэш). */
  blocks: Record<string, string>;
  images: Record<string, string>;
  log: LogEntry[];
}

export const emptyHistory = (): History => ({ versions: [], blocks: {}, images: {}, log: [] });

const MAX_VERSIONS = 40;
const MAX_NAMED = 60;
const MAX_LOG = 300;
const MAX_BYTES = 12 * 1024 * 1024;

/** Быстрый 53-битный хэш строки (cyrb53): для сравнения содержимого, не для защиты. */
export function hashString(s: string, seed = 0): string {
  let h1 = 0xdeadbeef ^ seed;
  let h2 = 0x41c6ce57 ^ seed;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    h1 = Math.imul(h1 ^ c, 2654435761);
    h2 = Math.imul(h2 ^ c, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  return (4294967296 * (2097151 & h2) + (h1 >>> 0)).toString(36);
}

const DATA_URL = /"src":"(data:[^"]+)"/g;

const wordsOf = (doc: PMNode): number => (doc.textBetween(0, doc.content.size, ' ', ' ').match(/[\p{L}\p{N}][\p{L}\p{N}'’-]*/gu) ?? []).length;

/** Хэш блоков и страницы: одинаков у одинакового содержимого, где бы документ ни хранился. */
export function contentHash(blockIds: string[], meta: DocMeta): string {
  return hashString(`${blockIds.join(',')}|${JSON.stringify(meta.page)}|${JSON.stringify(meta.defaults)}|${JSON.stringify(meta.hf)}`);
}

/** Раскладывает документ на блоки и кладёт их в хранилище. Возвращает хэши блоков. */
export function storeBlocks(h: History, doc: PMNode): string[] {
  const ids: string[] = [];
  doc.forEach((node) => {
    const json = JSON.stringify(node.toJSON()).replace(DATA_URL, (_m, url: string) => {
      const ih = hashString(url);
      h.images[ih] = url;
      return `"src":"ens-img:${ih}"`;
    });
    const id = hashString(json);
    h.blocks[id] ??= json;
    ids.push(id);
  });
  return ids;
}

/** Хэш документа без записи в хранилище (чтобы узнать, менялось ли что-то). */
export function docHash(doc: PMNode, meta: DocMeta): string {
  const ids: string[] = [];
  doc.forEach((node) => {
    ids.push(hashString(JSON.stringify(node.toJSON()).replace(DATA_URL, (_m, url: string) => `"src":"ens-img:${hashString(url)}"`)));
  });
  return contentHash(ids, meta);
}

export function addVersion(h: History, doc: PMNode, meta: DocMeta, o: { u: string; kind: VersionKind; label?: string; t?: number }): Version {
  const blocks = storeBlocks(h, doc);
  const hash = contentHash(blocks, meta);
  const v: Version = {
    id: `v${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`,
    t: o.t ?? Date.now(), u: o.u, kind: o.kind, label: o.label ?? '', hash, blocks, words: wordsOf(doc),
    meta: JSON.parse(JSON.stringify({ page: meta.page, defaults: meta.defaults, hf: meta.hf })),
  };
  h.versions.push(v);
  return v;
}

/** Документ версии: блоки и картинки возвращаются из хранилища. */
export function versionDoc(h: History, v: Version): PMNode {
  const nodes = v.blocks.map((id) => {
    const json = (h.blocks[id] ?? '{}').replace(/"src":"ens-img:([^"]+)"/g, (_m, ih: string) => `"src":${JSON.stringify(h.images[ih] ?? '')}`);
    return schema.nodeFromJSON(JSON.parse(json));
  });
  return schema.nodes.doc.create(null, nodes);
}

/** Один блок из хранилища как узел документа. */
export function nodeFromBlock(h: Pick<History, 'blocks' | 'images'>, id: string): PMNode {
  const json = (h.blocks[id] ?? '{}').replace(/"src":"ens-img:([^"]+)"/g, (_m, ih: string) => `"src":${JSON.stringify(h.images[ih] ?? '')}`);
  return schema.nodeFromJSON(JSON.parse(json));
}

export function versionMeta(v: Version): { page: DocMeta['page']; defaults: DocMeta['defaults']; hf: DocMeta['hf'] } {
  return JSON.parse(JSON.stringify(v.meta));
}

export function addLog(h: History, e: LogEntry): void {
  h.log.push(e);
  if (h.log.length > MAX_LOG) h.log.splice(0, h.log.length - MAX_LOG);
}

/** Прореживание: именные версии и сохранения живут дольше всего, автоверсии — меньше всего. */
export function trim(h: History): void {
  const rank = (k: VersionKind) => (k === 'named' ? 3 : k === 'save' ? 2 : 1);
  while (h.versions.length > MAX_VERSIONS || h.versions.filter((v) => v.kind === 'named').length > MAX_NAMED) {
    let idx = -1;
    let best = 99;
    h.versions.forEach((v, i) => {
      if (rank(v.kind) < best) { best = rank(v.kind); idx = i; }
    });
    // Самую раннюю из наименее ценных (forEach идёт от старых к новым, берём первую найденную).
    if (idx < 0) break;
    h.versions.splice(idx, 1);
  }
  collect(h);
}

/** Убирает блоки и картинки, на которые не ссылается ни одна версия. */
export function collect(h: History): void {
  const used = new Set<string>();
  for (const v of h.versions) for (const b of v.blocks) used.add(b);
  for (const k of Object.keys(h.blocks)) if (!used.has(k)) delete h.blocks[k];
  const imgs = new Set<string>();
  for (const json of Object.values(h.blocks)) for (const m of json.matchAll(/ens-img:([a-z0-9]+)/g)) imgs.add(m[1]);
  for (const k of Object.keys(h.images)) if (!imgs.has(k)) delete h.images[k];
}

// --- Запись в файл

const toBase64 = (b: Uint8Array): string => {
  let s = '';
  for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode(...b.subarray(i, i + 0x8000));
  return btoa(s);
};
const fromBase64 = (s: string): Uint8Array => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

export function serialize(h: History): string {
  trim(h);
  let raw = deflateSync(strToU8(JSON.stringify(h)), { level: 6 });
  // Слишком тяжёлая история: отбрасываем самые старые необязательные версии, пока не влезет.
  while (raw.length > MAX_BYTES && h.versions.length > 1) {
    const i = h.versions.findIndex((v) => v.kind !== 'named');
    h.versions.splice(i < 0 ? 0 : i, 1);
    collect(h);
    raw = deflateSync(strToU8(JSON.stringify(h)), { level: 6 });
  }
  return toBase64(raw);
}

/** Читает историю из файла; повреждённая или чужая — пустая (документ от этого открывается). */
export function parse(text: string | undefined): History {
  if (!text) return emptyHistory();
  try {
    const h = JSON.parse(strFromU8(inflateSync(fromBase64(text.trim())))) as History;
    if (!Array.isArray(h.versions) || typeof h.blocks !== 'object' || !Array.isArray(h.log)) return emptyHistory();
    h.images ??= {};
    return h;
  } catch {
    return emptyHistory();
  }
}

// --- Сравнение

export type Hunk =
  | { type: 'same'; a: [number, number]; b: [number, number] }
  | { type: 'change'; a: [number, number]; b: [number, number] };

/** Сравнение двух списков блоков (по хэшам): общие куски и отличия. */
export function diffBlocks(a: string[], b: string[]): Hunk[] {
  let pre = 0;
  while (pre < a.length && pre < b.length && a[pre] === b[pre]) pre++;
  let suf = 0;
  while (suf < a.length - pre && suf < b.length - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
  const am = a.slice(pre, a.length - suf);
  const bm = b.slice(pre, b.length - suf);
  const out: Hunk[] = [];
  const push = (h: Hunk) => {
    const last = out[out.length - 1];
    if (last && last.type === h.type && last.a[1] === h.a[0] && last.b[1] === h.b[0]) { last.a[1] = h.a[1]; last.b[1] = h.b[1]; } else out.push(h);
  };
  if (pre) push({ type: 'same', a: [0, pre], b: [0, pre] });
  if (am.length * bm.length > 6_000_000) {
    // Слишком много отличий для точного сравнения: считаем средний кусок одним изменением.
    if (am.length || bm.length) push({ type: 'change', a: [pre, pre + am.length], b: [pre, pre + bm.length] });
  } else if (am.length || bm.length) {
    // Наибольшая общая подпоследовательность.
    const n = am.length;
    const m = bm.length;
    const w = m + 1;
    const t = new Uint32Array((n + 1) * w);
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) t[i * w + j] = am[i] === bm[j] ? t[(i + 1) * w + j + 1] + 1 : Math.max(t[(i + 1) * w + j], t[i * w + j + 1]);
    let i = 0;
    let j = 0;
    let ci = 0;
    let cj = 0;
    const flush = () => {
      if (i > ci || j > cj) push({ type: 'change', a: [pre + ci, pre + i], b: [pre + cj, pre + j] });
    };
    while (i < n || j < m) {
      if (i < n && j < m && am[i] === bm[j]) {
        flush();
        push({ type: 'same', a: [pre + i, pre + i + 1], b: [pre + j, pre + j + 1] });
        i++; j++; ci = i; cj = j;
      } else if (j < m && (i === n || t[i * w + j + 1] >= t[(i + 1) * w + j])) j++;
      else i++;
    }
    flush();
  }
  if (suf) push({ type: 'same', a: [a.length - suf, a.length], b: [b.length - suf, b.length] });
  return out;
}

/** Текст блока для показа в сравнении: абзацы, пункты списков, строки таблиц. */
export function blockText(json: string, images: Record<string, string> = {}): string {
  void images;
  const walk = (n: { type: string; text?: string; content?: unknown[]; attrs?: Record<string, unknown> }, prefix = ''): string => {
    switch (n.type) {
      case 'text': return n.text ?? '';
      case 'hard_break': return '\n';
      case 'image': return '[рисунок]';
      case 'field': return n.attrs?.kind === 'pages' ? '‹число страниц›' : '‹номер страницы›';
      case 'bullet_list':
      case 'ordered_list':
        return (n.content ?? []).map((li, i) => `${prefix}${n.type === 'bullet_list' ? '• ' : `${i + 1}. `}${walk(li as never, prefix + '  ')}`).join('\n');
      case 'table': return (n.content ?? []).map((r) => walk(r as never)).join('\n');
      case 'table_row': return (n.content ?? []).map((c) => walk(c as never)).join(' │ ');
      default: return (n.content ?? []).map((c) => walk(c as never, prefix)).join(n.type === 'list_item' || n.type === 'table_cell' ? '\n' : '');
    }
  };
  try {
    return walk(JSON.parse(json));
  } catch {
    return '';
  }
}

export interface Span { t: 'same' | 'del' | 'ins'; text: string }

/** Различия внутри изменённого абзаца по словам. */
export function wordDiff(a: string, b: string): Span[] {
  const ta = a.split(/(\s+)/).filter(Boolean);
  const tb = b.split(/(\s+)/).filter(Boolean);
  if (ta.length * tb.length > 250_000) return [{ t: 'del', text: a }, { t: 'ins', text: b }];
  const w = tb.length + 1;
  const t = new Uint32Array((ta.length + 1) * w);
  for (let i = ta.length - 1; i >= 0; i--) for (let j = tb.length - 1; j >= 0; j--) t[i * w + j] = ta[i] === tb[j] ? t[(i + 1) * w + j + 1] + 1 : Math.max(t[(i + 1) * w + j], t[i * w + j + 1]);
  const out: Span[] = [];
  const add = (type: Span['t'], text: string) => {
    const last = out[out.length - 1];
    if (last && last.t === type) last.text += text; else out.push({ t: type, text });
  };
  let i = 0;
  let j = 0;
  while (i < ta.length || j < tb.length) {
    if (i < ta.length && j < tb.length && ta[i] === tb[j]) { add('same', ta[i]); i++; j++; }
    else if (j < tb.length && (i === ta.length || t[i * w + j + 1] >= t[(i + 1) * w + j])) { add('ins', tb[j]); j++; }
    else { add('del', ta[i]); i++; }
  }
  return out;
}

/** Краткое описание изменений между двумя списками блоков (для журнала). */
export function summarize(h: History, from: string[], to: string[], fromWords: number, toWords: number): LogEntry['ch'] {
  let added = 0;
  let removed = 0;
  let changed = 0;
  for (const hunk of diffBlocks(from, to)) {
    if (hunk.type !== 'change') continue;
    const x = hunk.a[1] - hunk.a[0];
    const y = hunk.b[1] - hunk.b[0];
    const both = Math.min(x, y);
    changed += both;
    removed += x - both;
    added += y - both;
  }
  void h;
  return { added, removed, changed, words: toWords - fromWords, notes: [] };
}
