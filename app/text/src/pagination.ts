// Постраничная вёрстка. Текст остаётся одним потоком ProseMirror (ввод, выделение и отмена работают как
// обычно), а страницы — это подложка (.papers) и распорки в потоке в местах разрыва:
//   * между абзацами — верхний отступ у абзаца, который начинает страницу (Decoration.node);
//   * внутри абзаца — строчная распорка перед первой строкой новой страницы (Decoration.widget);
//   * внутри таблицы — строка-распорка между строками таблицы.
// Положение разрывов считается по измерению DOM в «естественных» координатах (за вычетом уже
// поставленных распорок), поэтому расчёт не зависит от собственных результатов и устойчив: если
// ничего не изменилось, повторный расчёт даёт те же распорки и на этом заканчивается.
// Правила как в Word: «висячие» строки (не меньше двух строк абзаца на странице), заголовок
// не остаётся один в конце страницы, разрыв страницы перед абзацем, строки таблицы не рвутся.
//
// Колонки. В разделе с колонками поток идёт одной полосой шириной в колонку, а колонка считается
// «виртуальной страницей» высотой рабочей области (виртуальные страницы стоят в потоке вплотную).
// Блоки (и куски абзаца после разрыва) на экран ставит относительное смещение (position: relative):
// вправо — на свою колонку, вверх — на свой лист. Абзац, перетекающий в следующую колонку, остаётся
// одним элементом: хвост его строк сдвигает строчная декорация (Decoration.inline), поэтому ввод,
// выделение, поиск позиции по щелчку и печать работают как обычно. Лишняя высота потока, набранная
// колонками, снимается отрицательным отступом перед следующим разделом и в конце документа.

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey } from 'prosemirror-state';
import { TableMap } from 'prosemirror-tables';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { PageSetup } from './docx';
import { listNotes, renderNote } from './footnotes';
import { sectionsOf } from './sections';

export const PX = 4 / 3;
/** Зазор между листами на экране, px (при печати 0: страницы вплотную, как на бумаге); perRow — листов в ряду
 * (режим «Несколько страниц»: листы рядами, как в Word; 1 — столбиком). */
export const layout = { gap: 16, perRow: 1 };

export interface Geo {
  /** Высота листа, px. */
  H: number;
  top: number;
  bottom: number;
  /** Высота рабочей области страницы. */
  C: number;
  /** Шаг страниц: лист + зазор. */
  P: number;
}

export function geometry(p: PageSetup): Geo {
  const H = p.h * PX;
  const top = p.top * PX;
  const bottom = p.bottom * PX;
  return { H, top, bottom, C: Math.max(40, H - top - bottom), P: H + layout.gap };
}

/**
 * Виды декораций вёрстки. Распорки (сдвигают поток вниз на h): pad — отступ сверху у блока, padb — отступ снизу
 * у блока (в колонках: пустое место внизу колонки принадлежит абзацу над ним), neg — отрицательный отступ
 * (блок пустой высоты перед блоком), split — строчная распорка, row — строка-распорка таблицы.
 * Смещения (поток не меняют): rel — блок на экране, seg — хвост абзаца в другой колонке; w — ширина колонки.
 */
type Kind = 'pad' | 'padb' | 'neg' | 'split' | 'row' | 'rel' | 'seg' | 'w';
interface Spec { pg: Kind; h: number; key?: string }
const SPACERS = new Set<Kind>(['pad', 'padb', 'neg', 'split', 'row']);

/** Колонки раздела в px: ширины, начало каждой от левого поля, промежуток. */
export interface ColGeo { n: number; w: number[]; x: number[]; gap: number; equal: boolean; sep: boolean }

/** Колонки раздела (null — одна). Неравные ширины (w:col) приводятся к ширине текста. */
export function columnsOf(p: PageSetup): ColGeo | null {
  const n = Math.min(12, Math.max(1, Math.round(p.cols ?? 1)));
  if (n < 2) return null;
  const text = Math.max(36, p.w - p.left - p.right);
  const gap = Math.min(Math.max(0, p.colSpace ?? 35.4), (text - 18 * n) / (n - 1));
  const free = text - gap * (n - 1);
  const own = p.colW && p.colW.length === n && p.colW.every((x) => x > 0) ? p.colW : null;
  const sum = own ? own.reduce((a, b) => a + b, 0) : 0;
  const wPt = own ? own.map((x) => (x * free) / sum) : Array.from({ length: n }, () => free / n);
  const equal = wPt.every((x) => Math.abs(x - wPt[0]) < 0.5);
  const w = wPt.map((x) => x * PX);
  const x: number[] = [];
  let acc = 0;
  for (let k = 0; k < n; k++) { x.push(acc); acc += w[k] + gap * PX; }
  return { n, w, x, gap: gap * PX, equal, sep: !!p.colSep };
}

/** Лист на экране: положение по вертикали и размер. `page` — параметры раздела, `eff` — с учётом колонтитулов.
 * `cols` — колонки листа; `sepH` — высота разделителя (по самой длинной колонке листа), px. */
/** Сноски внизу страницы (колонки): место на листе и сноски по порядку. */
export interface NoteArea { x: number; y: number; w: number; h: number; list: { label: string; content: unknown }[] }
export interface PageBox { top: number; H: number; W: number; page: PageSetup; eff: PageSetup; cols?: ColGeo; sepH?: number; /** Левый край листа на подложке (листы рядами), px. */ left?: number; notes?: NoteArea[] }

/** Отступ над сносками: линия-разделитель и промежуток, px. */
export const NOTE_SEP = 14;
const noteCache = new Map<string, number>();
let measureBox: HTMLElement | null = null;
/** Высота текста сноски при ширине w (тем же элементом, что и на листе). */
function noteHeight(label: string, content: unknown, w: number): number {
  const key = `${Math.round(w)}|${label}|${JSON.stringify(content)}`;
  const had = noteCache.get(key);
  if (had !== undefined) return had;
  if (!measureBox || !measureBox.isConnected) {
    measureBox = document.createElement('div');
    measureBox.className = 'fn-area fn-measure';
    measureBox.setAttribute('aria-hidden', 'true');
    document.body.append(measureBox);
  }
  measureBox.style.width = `${w}px`;
  const el = renderNote(label, content);
  measureBox.replaceChildren(el);
  const h = el.getBoundingClientRect().height / (measureBox.getBoundingClientRect().width / (measureBox.offsetWidth || 1) || 1);
  if (noteCache.size > 2000) noteCache.clear();
  noteCache.set(key, h);
  return h;
}
/** Шрифты загрузились или изменились параметры документа — высоты сносок меряются заново. */
export function resetNoteCache(): void { noteCache.clear(); }
interface PagState { deco: DecorationSet; pages: number; boxes: PageBox[]; rev: number; /** Растёт при правке и запросе пересчёта (не от своих распорок). */ ext: number }

/** Параметры одного раздела для расчёта: настоящие и с учётом колонтитулов. */
interface SecGeo { g: Geo; page: PageSetup; eff: PageSetup; cols: ColGeo | null }

export const pagKey = new PluginKey<PagState>('pagination');

const r1 = (v: number) => Math.round(v * 10) / 10;

function splitDom(h: number): HTMLElement {
  const s = document.createElement('span');
  s.className = 'pg-split';
  s.contentEditable = 'false';
  s.style.height = `${h}px`;
  return s;
}

/** Отрицательный отступ: пустой блок, его поле сливается с полями соседей, поэтому сдвиг ровно h. */
function negDom(h: number): HTMLElement {
  const d = document.createElement('div');
  d.className = 'pg-neg';
  d.contentEditable = 'false';
  d.style.marginTop = `${h}px`;
  return d;
}

function rowDom(h: number, cols: number): HTMLElement {
  const tr = document.createElement('tr');
  tr.className = 'pg-gap';
  tr.contentEditable = 'false';
  const td = document.createElement('td');
  td.colSpan = cols;
  td.style.height = `${h}px`;
  tr.append(td);
  return tr;
}

interface Item { sect?: boolean; pos: number; end: number; node: PMNode; el: HTMLElement; row: boolean; cols: number; /** Узел, которому ставится верхний отступ: у первого абзаца пункта списка — сам пункт (иначе номер остался бы на прошлой странице). */ padPos: number; padEnd: number }
interface Line { pos: number; top: number; bottom: number }
interface Planned { kind: Kind; pos: number; end: number; h: number; cols: number; padPos: number; padEnd: number; style?: string }

/** Расчёт разрывов по текущему DOM. */
function plan(view: EditorView, secs: SecGeo[]): { out: Planned[]; pages: number; boxes: PageBox[]; items: number } {
  const dom = view.dom as HTMLElement;
  const box = dom.getBoundingClientRect();
  const scale = box.width / (dom.offsetWidth || 1) || 1;
  const Y = (v: number) => (v - box.top) / scale;

  const found = (pagKey.getState(view.state)?.deco.find() ?? []);
  const applied = found
    .filter((d) => SPACERS.has((d.spec as Spec).pg))
    .map((d) => ({ from: d.from, to: d.to, kind: (d.spec as Spec).pg, h: (d.spec as Spec).h }))
    .sort((a, b) => a.from - b.from);
  // Уже поставленные смещения колонок: блоки (по началу элемента) и хвосты абзацев.
  const relAt = new Map<number, number>();
  const segs: { from: number; to: number; dy: number }[] = [];
  for (const d of found) {
    const sp = d.spec as Spec;
    if (sp.pg === 'rel') relAt.set(d.from, sp.h);
    else if (sp.pg === 'seg') segs.push({ from: d.from, to: d.to, dy: sp.h });
  }
  // Сколько высоты уже добавлено распорками до позиции (в вёрстке сейчас).
  const shiftBefore = (pos: number): number => {
    let s = 0;
    for (const e of applied) {
      if (e.from > pos) break;
      if (e.kind === 'padb' ? e.to <= pos : e.kind === 'row' || e.kind === 'neg' ? e.from <= pos : e.from < pos) s += e.h;
    }
    return s;
  };
  const shiftEnd = (end: number): number => {
    let s = 0;
    for (const e of applied) { if (e.from >= end) break; s += e.h; }
    return s;
  };
  const shiftAt = (q: number): number => {
    let s = 0;
    for (const e of applied) { if (e.from > q) break; if (e.kind !== 'padb' || e.to <= q) s += e.h; }
    return s;
  };

  const items: Item[] = [];
  view.state.doc.descendants((node, pos) => {
    // Надписи стоят вне потока: их абзацы страницы не делят.
    if (node.type.name === 'frame') return false;
    if (node.type.name === 'section_break') {
      const el = view.nodeDOM(pos) as HTMLElement | null;
      if (el) items.push({ sect: true, pos, end: pos + node.nodeSize, node, el, row: false, cols: 0, padPos: pos, padEnd: pos + node.nodeSize });
      return false;
    }
    if (node.type.name === 'table') {
      const cols = TableMap.get(node).width;
      node.forEach((row, off) => {
        const p = pos + 1 + off;
        const el = view.nodeDOM(p) as HTMLElement | null;
        if (el) items.push({ pos: p, end: p + row.nodeSize, node: row, el, row: true, cols, padPos: p, padEnd: p + row.nodeSize });
      });
      return false;
    }
    if (node.isTextblock) {
      const el = view.nodeDOM(pos) as HTMLElement | null;
      if (el) {
        const $p = view.state.doc.resolve(pos);
        const inItem = $p.parent.type.name === 'list_item' && $p.index() === 0;
        items.push({ pos, end: pos + node.nodeSize, node, el, row: false, cols: 0, padPos: inItem ? pos - 1 : pos, padEnd: inItem ? pos - 1 + $p.parent.nodeSize : pos + node.nodeSize });
      }
      return false;
    }
    return true;
  });

  // Смещение блока колонкой (по вертикали), уже поставленное в вёрстке.
  const relOf = (it: Item): number => relAt.get(it.pos) ?? relAt.get(it.padPos) ?? 0;
  /** Ширина колонки, по которой блок верхнего уровня свёрстан сейчас (с отступами абзаца), px. */
  const curW = (it: Item): number => {
    const $p = view.state.doc.resolve(it.pos);
    const top = $p.depth ? $p.before(1) : it.pos;
    const node = view.state.doc.nodeAt(top);
    const el = view.nodeDOM(top) as HTMLElement | null;
    if (!node || !el) return 0;
    const own = node.type.name === 'paragraph' || node.type.name === 'heading';
    return el.offsetWidth + (own ? (((node.attrs.indent as number | null) ?? 0) + ((node.attrs.rindent as number | null) ?? 0)) * PX : 0);
  };
  const nat = items.map((it) => {
    const r = it.el.getBoundingClientRect();
    const rel = relOf(it);
    return { top: Y(r.top) - shiftBefore(it.pos) - rel, bottom: Y(r.bottom) - shiftEnd(it.end) - rel, bottomView: r.bottom };
  });

  /** Строки абзаца: начальные позиции и естественные границы. */
  const lineBoxes = (it: Item, i: number): Line[] => {
    const start = it.pos + 1;
    const end = it.pos + 1 + it.node.content.size;
    // Хвосты абзаца в других колонках сдвинуты на экране: координаты строк приводим к потоку.
    const rel = relOf(it);
    const own = segs.filter((s) => s.from >= start && s.from <= end);
    const dyAt = (q: number): number => {
      let d = rel;
      for (const s of own) if (s.from <= q && (q < s.to || (q === s.to && q === end))) d = rel + s.dy;
      return d * scale;
    };
    const coords = (q: number) => {
      const c = view.coordsAtPos(q, 1);
      const d = dyAt(q);
      return { top: c.top - d, bottom: c.bottom - d };
    };
    const starts = [start];
    for (let guard = 0; guard < 2000; guard++) {
      const q = starts[starts.length - 1];
      const c = coords(q);
      const limit = c.top + (c.bottom - c.top) / 2;
      if (coords(end).top <= limit) break;
      let lo = q + 1;
      let hi = end;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (coords(mid).top > limit) hi = mid; else lo = mid + 1;
      }
      starts.push(lo);
    }
    const lh = parseFloat(getComputedStyle(it.el).lineHeight);
    const lines: Line[] = starts.map((q) => {
      const c = coords(q);
      const rect = c.bottom - c.top;
      const half = Number.isFinite(lh) && lh > rect ? (lh - rect) / 2 : 0;
      return { pos: q, top: Y(c.top - half * scale) - shiftAt(q), bottom: 0 };
    });
    for (let j = 0; j < lines.length; j++) lines[j].bottom = j + 1 < lines.length ? lines[j + 1].top : nat[i].bottom;
    return lines;
  };

  // Сноски (не концевые): номер по порядку и текст; у блока — те, что в нём стоят.
  const labels = new Map<number, { label: string; content: unknown }>();
  for (const n of listNotes(view.state.doc)) if (!n.end) labels.set(n.pos, { label: n.label, content: view.state.doc.nodeAt(n.pos)?.attrs.content ?? [] });
  const notesIn = (it: Item): number[] => {
    const res: number[] = [];
    if (!labels.size) return res;
    it.node.descendants((c, p) => { if (c.type.name === 'footnote' && labels.has(it.pos + 1 + p)) res.push(it.pos + 1 + p); return !c.isAtom; });
    return res;
  };
  /** Сноски по виртуальным страницам (ключ — раздел и номер): позиции и высота занятого места. */
  const placed = new Map<string, { pos: number[]; h: number }>();

  const out: Planned[] = [];
  let acc = 0;
  // Текущий раздел: y0 — где начинается текст на его первой странице; «виртуальные страницы» (страницы или,
  // в разделе с колонками, колонки) идут с шагом vP, страницы на экране — с шагом g.P.
  let secIdx = 0;
  const segOf = (y0: number, sg: SecGeo) => ({ y0, g: sg.g, cols: sg.cols, vP: sg.cols ? sg.g.C : sg.g.P });
  let seg = segOf(secs[0].g.top, secs[0]);
  let lastBottom = seg.y0;
  let startOfSection = false;
  const boxes: PageBox[] = [];
  // Листы рядами («Несколько страниц»): поток остаётся столбиком, лист g ставится в ряд floor(g/k), место g%k.
  // Ряды одной высоты — по самому высокому листу, места одной ширины — по самому широкому.
  const k = Math.max(1, Math.floor(layout.perRow));
  const grid = k > 1;
  const rowH = Math.max(...secs.map((x) => x.g.H)) + layout.gap;
  const colStep = Math.max(...secs.map((x) => x.page.w * PX)) + layout.gap;
  /** Номер первого листа текущего раздела во всём документе. */
  let segFirst = 0;
  const gridTop = (g: number) => Math.floor(g / k) * rowH;
  const gridLeft = (g: number) => (g % k) * colStep;
  /** Высота разделителя колонок по листам раздела. */
  let sepH: number[] = [];
  const vIdx = (y: number) => Math.max(0, Math.floor((y - seg.y0 + 0.5) / seg.vP));
  const pageOfV = (v: number) => (seg.cols ? Math.floor(v / seg.cols.n) : v);
  /** Смещение виртуальной страницы v на экран (колонки, листы рядами): по вертикали и по горизонтали. */
  const dyOf = (v: number) => {
    const p = pageOfV(v);
    const col = seg.cols ? p * seg.g.P - v * seg.vP : 0;
    return grid ? col + gridTop(segFirst + p) - (seg.y0 - seg.g.top + p * seg.g.P) : col;
  };
  const dxOf = (v: number) => (seg.cols ? seg.cols.x[v % seg.cols.n] : 0) + (grid ? gridLeft(segFirst + pageOfV(v)) : 0);
  const markSep = (v: number, h: number) => {
    const p = pageOfV(v);
    sepH[p] = Math.max(sepH[p] ?? 0, Math.min(seg.g.C, h));
  };
  const colW = (v: number) => (seg.cols ? seg.cols.w[v % seg.cols.n] : (secs[secIdx].page.w - secs[secIdx].page.left - secs[secIdx].page.right) * PX);
  const noteH = (pos: number, v: number) => { const l = labels.get(pos)!; return noteHeight(l.label, l.content, colW(v)); };
  /** Сколько места внизу виртуальной страницы v займут сноски, если добавить к ним ещё list. Не больше 60% высоты. */
  const reserve = (v: number, list: number[]): number => {
    const had = placed.get(`${secIdx}:${v}`);
    const h = (had?.h ?? 0) + list.reduce((s, p) => s + noteH(p, v), 0);
    return h > 0 ? Math.min(seg.g.C * 0.6, h + NOTE_SEP) : 0;
  };
  const place = (v: number, list: number[]) => {
    if (!list.length) return;
    const k = `${secIdx}:${v}`;
    const had = placed.get(k) ?? { pos: [], h: 0 };
    for (const p of list) { had.pos.push(p); had.h += noteH(p, v); }
    placed.set(k, had);
  };
  const closeSeg = (count: number) => {
    const top0 = seg.y0 - seg.g.top;
    const sg = secs[secIdx];
    for (let j = 0; j < count; j++) {
      const g = boxes.length;
      const b: PageBox = { top: grid ? gridTop(g) : top0 + j * seg.g.P, H: seg.g.H, W: sg.page.w * PX, page: sg.page, eff: sg.eff };
      if (grid) b.left = gridLeft(g);
      if (sg.cols) { b.cols = sg.cols; b.sepH = r1(sepH[j] ?? 0); }
      const per = sg.cols ? sg.cols.n : 1;
      for (let c = 0; c < per; c++) {
        const v = j * per + c;
        const pl = placed.get(`${secIdx}:${v}`);
        if (!pl?.pos.length) continue;
        const h = Math.min(seg.g.C * 0.6, pl.h + NOTE_SEP);
        (b.notes ??= []).push({ x: sg.page.left * PX + (sg.cols ? sg.cols.x[c] : 0), y: seg.g.top + seg.g.C - h, w: colW(v), h: r1(h), list: pl.pos.map((p) => labels.get(p)!) });
      }
      boxes.push(b);
    }
    sepH = [];
    segFirst = boxes.length;
  };
  // Смещения колонок: открытые элементы (li держит смещение своего первого абзаца — вместе с номером).
  const relStack: { to: number; dx: number; dy: number }[] = [];
  /** Абзац предыдущего элемента раздела: отступ при переходе в новую колонку ставится ему снизу. */
  let prevText: Item | null = null;
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const { top: nTop, bottom: nBottom } = nat[i];
    if (it.sect) {
      // Разрыв раздела: следующий раздел начинается с новой страницы, у неё могут быть другие размер и поля.
      const pIdx = pageOfV(vIdx(lastBottom - 1));
      closeSeg(pIdx + 1);
      const nextTop = seg.y0 - seg.g.top + (pIdx + 1) * seg.g.P;
      secIdx = Math.min(secIdx + 1, secs.length - 1);
      seg = segOf(nextTop + secs[secIdx].g.top, secs[secIdx]);
      startOfSection = true;
      lastBottom = seg.y0;
      prevText = null;
      continue;
    }
    const h0 = nBottom - nTop;
    const mine = notesIn(it);
    let pad = 0;
    const splits: { pos: number; h: number; v: number }[] = [];
    const cols = seg.cols;
    let forced: 'page' | 'col' | null = !it.row && it.node.attrs.pageBreakBefore ? 'page' : !it.row && it.node.attrs.colBreakBefore ? (cols ? 'col' : 'page') : null;
    const heading = !it.row && it.node.type.name === 'heading';
    const region = (a: number) => {
      const p = Math.max(0, Math.floor((a - seg.y0 + 0.5) / seg.vP));
      const start = p * seg.vP + seg.y0;
      return { start, end: start + seg.g.C, v: p };
    };
    if (startOfSection) {
      // Первый блок раздела встаёт в начало текста новой страницы (после колонок поток бывает и ниже — тогда
      // отступ отрицательный).
      pad = seg.y0 - (nTop + acc);
      startOfSection = false;
      forced = null;
    }

    for (let guard = 0; guard < 80; guard++) {
      const a = nTop + acc + pad;
      const R = region(a);
      if (forced) {
        const f = forced;
        forced = null;
        if (f === 'page' && cols) {
          if (a - R.start > 1 || R.v % cols.n) { pad += seg.y0 + (Math.floor(R.v / cols.n) + 1) * cols.n * seg.vP - a; continue; }
        } else if (a - R.start > 1) { pad += R.start + seg.vP - a; continue; }
      }
      if (a >= R.end - reserve(R.v, []) - 0.5) { pad += R.start + seg.vP - a; continue; }
      // Колонки разной ширины: блок сейчас свёрстан по ширине прежней колонки. Если новая уже, он станет выше —
      // высоту берём с запасом (пропорционально ширине), иначе расчёт качался бы между колонками.
      const h = cols && !cols.equal && !it.row ? h0 * Math.max(1, curW(it) / cols.w[R.v % cols.n]) : h0;
      if (a + h <= R.end - reserve(R.v, mine) + 0.5) {
        // Заголовок не остаётся один внизу страницы: за ним должны поместиться хотя бы две строки.
        if (heading && i + 1 < items.length && !items[i + 1].sect && a - R.start > 1) {
          const n = nat[i + 1];
          const need = n.top - nTop + Math.min(n.bottom - n.top, 2 * h0);
          if (a + need > R.end - reserve(R.v, mine) + 0.5) { pad += R.start + seg.vP - a; continue; }
        }
        break;
      }
      // Не помещается.
      if (it.row) {
        if (a - R.start > 1) { pad += R.start + seg.vP - a; continue; }
        break; // строка выше страницы — не рвём
      }
      const atTop = a - R.start <= 1;
      // Колонки разной ширины: абзац переносится целиком (строки одного абзаца не могут быть разной длины).
      if (cols && !cols.equal && !atTop) { pad += R.start + seg.vP - a; continue; }
      const lines = lineBoxes(it, i);
      const n = lines.length;
      // Сноски строк from..to (включительно): строка сноски — та, в которой стоит её номер.
      const lineNotes = (from: number, to: number) => mine.filter((p) => p >= lines[from].pos && (to + 1 >= n || p < lines[to + 1].pos));
      let k = 0;
      while (k < n && lines[k].bottom + acc + pad <= R.end - reserve(R.v, lineNotes(0, k)) + 0.5) k++;
      if (k >= n) break;
      if (k < 2 && n >= 2 && !atTop) { pad += R.start + seg.vP - a; continue; }
      if (n - k < 2) {
        if (k >= 3) k -= 1;
        else if (!atTop) { pad += R.start + seg.vP - a; continue; }
      }
      k = Math.max(1, k);
      // Строчная распорка не сдвигает строку меньше, чем на высоту строки (у её строки есть «подпорка» шрифта).
      // На странице остаток всегда больше строки (поля), а в колонке бывает меньше — тогда хвост встаёт чуть ниже.
      let pitch = Infinity;
      for (let m = 1; m < n; m++) pitch = Math.min(pitch, lines[m].top - lines[m - 1].top);
      // С запасом: распорка выше подпорки сдвигает ровно на свою высоту (иначе расчёт качается на доли пикселя).
      pitch = Number.isFinite(pitch) ? Math.ceil(pitch) + 1 : 0;
      let local = acc + pad;
      let end = R.end;
      let next = R.start + seg.vP;
      let j = k;
      let from = 0;
      for (let safety = 0; safety < 200 && j < n; safety++) {
        // Остаток внизу колонки меньше строки: в следующую колонку уходит на строку больше, и она начинается ровно
        // сверху (распорка ниже строки сдвинула бы хвост на целую строку). Висячие строки при этом не нарушаем.
        if (cols && next - (lines[j].top + local) < pitch && j - 1 - from >= (from === 0 && atTop ? 1 : 2)) j -= 1;
        from = j;
        const spacer = Math.max(next - (lines[j].top + local), pitch);
        splits.push({ pos: lines[j].pos, h: spacer, v: vIdx(next) });
        if (cols) markSep(vIdx(next) - 1, seg.g.C);
        local += spacer;
        end += seg.vP;
        next += seg.vP;
        let m = j;
        const vNext = vIdx(next - seg.vP);
        while (m < n && lines[m].bottom + local <= end - reserve(vNext, lineNotes(j, m)) + 0.5) m++;
        if (m >= n) break;
        if (n - m < 2 && m - 1 >= j + 2) m -= 1;
        j = Math.max(m, j + 1);
      }
      break;
    }

    if (mine.length) {
      // Сноски — на страницу (колонку) строки со своим номером.
      const v0n = vIdx(nTop + acc + pad);
      const bySeg = new Map<number, number[]>();
      for (const p of mine) {
        let v = v0n;
        for (const s of splits) if (p >= s.pos) v = s.v;
        bySeg.set(v, [...(bySeg.get(v) ?? []), p]);
      }
      for (const [v, list] of bySeg) place(v, list);
    }
    const padR = r1(pad);
    if (pad < -0.5) out.push({ kind: 'neg', pos: it.padPos, end: it.padPos, h: padR, cols: 0, padPos: it.padPos, padEnd: it.padPos });
    else if (pad > 0.5) {
      // В колонках пустое место внизу колонки отдаём абзацу над ним (отступ снизу): отступ сверху у блока
      // в следующей колонке лежал бы на экране поверх чужого текста и перехватывал щелчки.
      if ((cols || grid) && !it.row && prevText) out.push({ kind: 'padb', pos: prevText.pos, end: prevText.end, h: padR, cols: 0, padPos: prevText.pos, padEnd: prevText.end });
      else out.push({ kind: it.row ? 'row' : 'pad', pos: it.pos, end: it.end, h: padR, cols: it.cols, padPos: it.padPos, padEnd: it.padEnd });
    }
    for (const s of splits) out.push({ kind: 'split', pos: s.pos, end: s.pos, h: r1(s.h), cols: 0, padPos: s.pos, padEnd: s.pos });
    const accBefore = acc;
    acc += (Math.abs(pad) > 0.5 ? padR : 0) + splits.reduce((s, x) => s + r1(x.h), 0);
    lastBottom = nBottom + acc;

    if (cols || grid) {
      // Смещение на экран: блок — по виртуальной странице начала текста, хвосты — по своим.
      const v0 = vIdx(nTop + accBefore + (Math.abs(pad) > 0.5 ? padR : 0));
      const dy0 = dyOf(v0), dx0 = dxOf(v0);
      const own = it.padPos !== it.pos ? { from: it.padPos, to: it.padEnd } : { from: it.pos, to: it.end };
      while (relStack.length && relStack[relStack.length - 1].to <= own.from) relStack.pop();
      const parent = relStack[relStack.length - 1];
      const ldx = dx0 - (parent?.dx ?? 0), ldy = dy0 - (parent?.dy ?? 0);
      out.push({ kind: 'rel', pos: own.from, end: own.to, h: r1(dy0), cols: 0, padPos: own.from, padEnd: own.to, style: `position:relative;left:${r1(ldx)}px;top:${r1(ldy)}px` });
      relStack.push({ to: own.to, dx: dx0, dy: dy0 });
      const contentEnd = it.pos + 1 + it.node.content.size;
      splits.forEach((s, k) => {
        const to = k + 1 < splits.length ? splits[k + 1].pos : contentEnd;
        if (to <= s.pos) return;
        const dy = dyOf(s.v) - dy0, dx = dxOf(s.v) - dx0;
        out.push({ kind: 'seg', pos: s.pos, end: to, h: r1(dy), cols: 0, padPos: s.pos, padEnd: to, style: `position:relative;left:${r1(dx)}px;top:${r1(dy)}px` });
      });
      if (cols && !cols.equal) {
        // Ширина блока верхнего уровня — ширина его колонки (у абзаца, выше колонки, — самой узкой).
        const $p = view.state.doc.resolve(it.pos);
        if ($p.depth === 0 || (it.row && $p.depth === 1)) {
          const top = it.row ? $p.before(1) : it.pos;
          const node = view.state.doc.nodeAt(top);
          const vs = [v0, ...splits.map((s) => s.v)];
          const w = Math.min(...vs.map((v) => cols.w[v % cols.n]));
          if (node && !it.row) out.push({ kind: 'w', pos: top, end: top + node.nodeSize, h: 0, cols: 0, padPos: top, padEnd: top + node.nodeSize, style: widthCss(node, w) });
        }
      }
      const vLast = vIdx(lastBottom - 1);
      if (cols) markSep(vLast, lastBottom - (seg.y0 + vLast * seg.vP));
    }
    prevText = it.row ? null : it;
  }
  const vEnd = vIdx(lastBottom - 1);
  // Поток после колонок (и при листах рядами) длиннее, чем листы на экране: лишнее снимаем в конце документа.
  // Считаем до closeSeg: он сдвигает номер первого листа раздела.
  const tail = r1(dyOf(vEnd));
  closeSeg(Math.max(1, pageOfV(vEnd) + 1));
  if (tail < -0.5) { const e = view.state.doc.content.size; out.push({ kind: 'neg', pos: e, end: e, h: tail, cols: 0, padPos: e, padEnd: e }); }
  return { out, pages: boxes.length, boxes, items: items.length };
}

/** Ширина блока верхнего уровня в колонке, px: у абзаца — за вычетом его отступов. */
function widthCss(node: PMNode, w: number): string {
  const own = node.type.name === 'paragraph' || node.type.name === 'heading';
  const il = own ? ((node.attrs.indent as number | null) ?? 0) : 0;
  const ir = own ? ((node.attrs.rindent as number | null) ?? 0) : 0;
  return `width:${r1(w - (il + ir) * PX)}px;box-sizing:border-box`;
}

function toDecos(out: Planned[]): Decoration[] {
  return out.map((p) => {
    if (p.kind === 'pad') return Decoration.node(p.padPos, p.padEnd, { class: 'pg-pad', style: `padding-top:${p.h}px` }, { pg: 'pad', h: p.h });
    if (p.kind === 'padb') return Decoration.node(p.padPos, p.padEnd, { class: 'pg-padb', style: `padding-bottom:${p.h}px` }, { pg: 'padb', h: p.h });
    if (p.kind === 'neg') return Decoration.widget(p.pos, () => negDom(p.h), { side: -1, key: `neg${p.h}`, pg: 'neg', h: p.h });
    if (p.kind === 'rel') return Decoration.node(p.padPos, p.padEnd, { class: 'pg-col', style: p.style ?? '' }, { pg: 'rel', h: p.h });
    if (p.kind === 'seg') return Decoration.inline(p.padPos, p.padEnd, { class: 'pg-colseg', style: p.style ?? '' }, { pg: 'seg', h: p.h });
    if (p.kind === 'w') return Decoration.node(p.padPos, p.padEnd, { style: p.style ?? '' }, { pg: 'w', h: 0, key: p.style });
    if (p.kind === 'row') return Decoration.widget(p.pos, () => rowDom(p.h, p.cols), { side: -1, key: `row${p.h}:${p.cols}`, pg: 'row', h: p.h });
    return Decoration.widget(p.pos, () => splitDom(p.h), { side: -1, key: `split${p.h}`, pg: 'split', h: p.h });
  });
}

const boxKey = (b: PageBox) => `${Math.round(b.left ?? 0)}:${Math.round(b.top)}:${b.W}:${b.H}:${b.cols ? `${b.cols.n},${b.cols.sep ? 1 : 0},${b.sepH}` : ''}:${b.notes ? JSON.stringify(b.notes) : ''}`;
type Mark = { from: number; to: number; spec: Spec };
const close = (a: Mark[], b: Mark[]): boolean => a.length === b.length && a.every((x, i) => {
  const y = b[i];
  return x.from === y.from && x.to === y.to && x.spec.pg === y.spec.pg && Math.abs(x.spec.h - y.spec.h) <= 0.3 && (x.spec.pg !== 'w' || x.spec.key === y.spec.key);
});
const sig = (list: { from: number; to: number; spec: Spec }[]) => list.map((d) => `${d.spec.pg}:${d.from}:${d.to}:${d.spec.h}:${d.spec.key ?? ''}`).join('|');

/** Поля по горизонтали для блоков разделов, отличающихся от первого: страницы разной ширины стоят по центру листа-подложки. */
function marginDecos(doc: PMNode, base: PageSetup): Decoration[] | null {
  const secs = sectionsOf(doc, base);
  if (secs.length < 2) return null;
  const wmax = Math.max(...secs.map((x) => x.page.w));
  const out: Decoration[] = [];
  let si = 0;
  doc.forEach((node, off) => {
    while (si + 1 < secs.length && off >= secs[si + 1].from) si++;
    const p = secs[si].page;
    const side = (wmax - p.w) / 2;
    // Стиль декорации идёт после стиля узла и заменяет его margin: отступы абзаца (слева, справа) прибавляем сами,
    // иначе абзац с отступом и выступом (3 см и 3 см) начинал первую строку левее листа раздела.
    const own = node.type.name === 'paragraph' || node.type.name === 'heading';
    const il = own ? ((node.attrs.indent as number | null) ?? 0) : 0;
    const ir = own ? ((node.attrs.rindent as number | null) ?? 0) : 0;
    // Отрицательный отступ уводит текст в поле, но не за край листа раздела.
    const ml = Math.max(side, side + p.left + il);
    // --pl: от края листа раздела до текста — им ограничен выступ первой строки (schema.ts, paraStyle).
    out.push(Decoration.node(off, off + node.nodeSize, { style: `margin-left:${ml}pt;margin-right:${side + p.right + ir}pt;--pl:${p.left}pt` }));
  });
  return out;
}

/** Ширина блоков верхнего уровня в разделах с колонками равной ширины (неравные ставит расчёт вёрстки). */
function columnDecos(doc: PMNode, base: PageSetup): Decoration[] {
  const secs = sectionsOf(doc, base);
  const out: Decoration[] = [];
  if (!secs.some((x) => (x.page.cols ?? 1) > 1)) return out;
  let si = 0;
  doc.forEach((node, off) => {
    while (si + 1 < secs.length && off >= secs[si + 1].from) si++;
    const c = columnsOf(secs[si].page);
    if (!c || node.type.name === 'section_break' || node.type.name === 'frame') return;
    // Таблица — по ширине самой узкой колонки; строки, ушедшие в другую колонку, обёртка не обрезает.
    if (node.type.name === 'table') out.push(Decoration.node(off, off + node.nodeSize, { class: 'pg-colt', style: `width:${r1(Math.min(...c.w))}px;max-width:${r1(Math.min(...c.w))}px` }));
    else if (c.equal) out.push(Decoration.node(off, off + node.nodeSize, { style: widthCss(node, c.w[0]) }));
  });
  return out;
}

/**
 * Вызывает onPages при смене числа или размеров страниц.
 * getBase — параметры первого раздела; eff — они же с учётом высоты колонтитулов.
 */
export function pagination(getBase: () => PageSetup, eff: (p: PageSetup) => PageSetup, onPages: (pages: number, boxes: PageBox[]) => void): Plugin<PagState> {
  let marginCache: { doc: PMNode; base: string; list: Decoration[] | null } | null = null;
  return new Plugin<PagState>({
    key: pagKey,
    state: {
      init: () => ({ deco: DecorationSet.empty, pages: 1, boxes: [], rev: 0, ext: 0 }),
      apply(tr, prev) {
        const meta = tr.getMeta(pagKey) as { deco?: DecorationSet; pages?: number; boxes?: PageBox[]; replan?: boolean } | undefined;
        if (meta?.deco) return { deco: meta.deco, pages: meta.pages ?? prev.pages, boxes: meta.boxes ?? prev.boxes, rev: prev.rev + 1, ext: prev.ext };
        if (tr.docChanged) return { ...prev, deco: prev.deco.map(tr.mapping, tr.doc), rev: prev.rev + 1, ext: prev.ext + 1 };
        if (meta?.replan) return { ...prev, rev: prev.rev + 1, ext: prev.ext + 1 };
        return prev;
      },
    },
    props: {
      decorations(state) {
        const st = pagKey.getState(state);
        if (!st) return null;
        const base = getBase();
        const key = JSON.stringify(base);
        if (!marginCache || marginCache.doc !== state.doc || marginCache.base !== key) {
          // Поля разделов и ширина колонок — одной декорацией на узел (две на одном узле add() теряет).
          const styles = new Map<number, { to: number; style: string; cls: string }>();
          for (const d of [...(marginDecos(state.doc, base) ?? []), ...columnDecos(state.doc, base)]) {
            const a = (d as unknown as { type: { attrs: { style: string; class?: string } } }).type.attrs;
            const had = styles.get(d.from);
            styles.set(d.from, { to: d.to, style: had ? `${had.style};${a.style}` : a.style, cls: [had?.cls, a.class].filter(Boolean).join(' ') });
          }
          const list = [...styles].map(([from, x]) => Decoration.node(from, x.to, x.cls ? { style: x.style, class: x.cls } : { style: x.style }));
          marginCache = { doc: state.doc, base: key, list: list.length ? list : null };
        }
        // add() и create() портят переданный массив — отдаём копию.
        return marginCache.list ? st.deco.add(state.doc, [...marginCache.list]) : st.deco;
      },
    },
    view(view) {
      let seen = -1;
      let passes = 0;
      let seenExt = 0;
      let scheduled = false;
      let lastSig = '';
      let alive = true;
      /** В больших документах расчёт откладывается: ввод не ждёт измерения всех абзацев. */
      let big = false;
      const run = () => {
        scheduled = false;
        if (!alive) return;
        const st = pagKey.getState(view.state);
        if (!st) return;
        const secs: SecGeo[] = sectionsOf(view.state.doc, getBase()).map((x) => { const e = eff(x.page); return { g: geometry(e), page: x.page, eff: e, cols: columnsOf(x.page) }; });
        const t0 = performance.now();
        const { out, pages, boxes, items } = plan(view, secs);
        big = items > 600;
        stats.ms = performance.now() - t0;
        stats.items = items;
        const old = st.deco.find().map((d) => ({ from: d.from, to: d.to, spec: d.spec as Spec }));
        // DecorationSet.create портит переданный массив: набор строим один раз.
        const set = DecorationSet.create(view.state.doc, toDecos(out));
        const fresh = set.find().map((d) => ({ from: d.from, to: d.to, spec: d.spec as Spec }));
        const boxSig = boxes.map(boxKey).join('|');
        const stateSig = st.boxes.map(boxKey).join('|');
        if (boxSig !== lastSig) { lastSig = boxSig; onPages(pages, boxes); }
        const order = (a: { from: number; to: number; spec: Spec }, b: { from: number; to: number; spec: Spec }) => a.from - b.from || a.to - b.to || (a.spec.pg < b.spec.pg ? -1 : a.spec.pg > b.spec.pg ? 1 : 0);
        // Обычно расчёт сходится за 2–3 прохода и сравнение точное; дальше — с допуском в доли пикселя:
        // субпиксельное округление вёрстки иначе качало бы расчёт.
        const same = passes >= 3 ? close(old.sort(order), fresh.sort(order)) : sig(old.sort(order)) === sig(fresh.sort(order));
        stats.stuck = false;
        if (same && boxSig === stateSig) { passes = 0; return; }
        // Защита от зацикливания. Неравным колонкам нужно больше проходов: ширина блока зависит от его колонки.
        if (++passes > 16) { stats.stuck = true; return; }
        view.dispatch(view.state.tr.setMeta(pagKey, { deco: same ? st.deco : set, pages, boxes }).setMeta('addToHistory', false));
      };
      const schedule = () => {
        if (scheduled) return;
        scheduled = true;
        if (big) setTimeout(run, 80); else queueMicrotask(run);
      };
      const onLoad = () => replan(view);
      view.dom.addEventListener('load', onLoad, true);
      document.fonts?.addEventListener?.('loadingdone', onLoad);
      schedule();
      return {
        update(v) {
          const st = pagKey.getState(v.state);
          const rev = st?.rev ?? 0;
          // Правка или запрос пересчёта — новый отсчёт проходов (иначе после упора в предел вёрстка замерла бы).
          if (st && st.ext !== seenExt) { seenExt = st.ext; passes = 0; }
          if (rev === seen) return;
          seen = rev;
          schedule();
        },
        destroy() {
          alive = false;
          view.dom.removeEventListener('load', onLoad, true);
          document.fonts?.removeEventListener?.('loadingdone', onLoad);
        },
      };
    },
  });
}

/** Пересчитать вёрстку (изменились поля, шрифты, картинки). */
export function replan(view: EditorView): void {
  resetNoteCache();
  view.dispatch(view.state.tr.setMeta(pagKey, { replan: true }).setMeta('addToHistory', false));
}

/** Время последнего расчёта (для проверок). */
export const stats = { ms: 0, items: 0, /** Расчёт не сошёлся за отведённые проходы. */ stuck: false };

export const pageCount = (view: EditorView): number => pagKey.getState(view.state)?.pages ?? 1;
export const pageBoxes = (view: EditorView): PageBox[] => pagKey.getState(view.state)?.boxes ?? [];
