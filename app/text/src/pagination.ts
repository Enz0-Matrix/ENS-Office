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

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey } from 'prosemirror-state';
import { TableMap } from 'prosemirror-tables';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { PageSetup } from './docx';
import { sectionsOf } from './sections';

export const PX = 4 / 3;
/** Зазор между листами на экране, px (при печати 0: страницы вплотную, как на бумаге). */
export const layout = { gap: 16 };

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

interface Spec { pg: 'pad' | 'split' | 'row'; h: number; key?: string }

/** Лист на экране: положение по вертикали и размер. `page` — параметры раздела, `eff` — с учётом колонтитулов. */
export interface PageBox { top: number; H: number; W: number; page: PageSetup; eff: PageSetup }
interface PagState { deco: DecorationSet; pages: number; boxes: PageBox[]; rev: number }

/** Параметры одного раздела для расчёта: настоящие и с учётом колонтитулов. */
interface SecGeo { g: Geo; page: PageSetup; eff: PageSetup }

export const pagKey = new PluginKey<PagState>('pagination');

const r1 = (v: number) => Math.round(v * 10) / 10;

function splitDom(h: number): HTMLElement {
  const s = document.createElement('span');
  s.className = 'pg-split';
  s.contentEditable = 'false';
  s.style.height = `${h}px`;
  return s;
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
interface Planned { kind: Spec['pg']; pos: number; end: number; h: number; cols: number; padPos: number; padEnd: number }

/** Расчёт разрывов по текущему DOM. */
function plan(view: EditorView, secs: SecGeo[]): { out: Planned[]; pages: number; boxes: PageBox[]; items: number } {
  const dom = view.dom as HTMLElement;
  const box = dom.getBoundingClientRect();
  const scale = box.width / (dom.offsetWidth || 1) || 1;
  const Y = (v: number) => (v - box.top) / scale;

  const applied = (pagKey.getState(view.state)?.deco.find() ?? [])
    .map((d) => ({ from: d.from, kind: (d.spec as Spec).pg, h: (d.spec as Spec).h }))
    .sort((a, b) => a.from - b.from);
  // Сколько высоты уже добавлено распорками до позиции (в вёрстке сейчас).
  const shiftBefore = (pos: number): number => {
    let s = 0;
    for (const e of applied) {
      if (e.from > pos) break;
      if (e.kind === 'row' ? e.from <= pos : e.from < pos) s += e.h;
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
    for (const e of applied) { if (e.from > q) break; s += e.h; }
    return s;
  };

  const items: Item[] = [];
  view.state.doc.descendants((node, pos) => {
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

  const nat = items.map((it) => {
    const r = it.el.getBoundingClientRect();
    return { top: Y(r.top) - shiftBefore(it.pos), bottom: Y(r.bottom) - shiftEnd(it.end), bottomView: r.bottom };
  });

  const coords = (q: number) => view.coordsAtPos(q, 1);

  /** Строки абзаца: начальные позиции и естественные границы. */
  const lineBoxes = (it: Item, i: number): Line[] => {
    const start = it.pos + 1;
    const end = it.pos + 1 + it.node.content.size;
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

  const out: Planned[] = [];
  let acc = 0;
  // Текущий раздел: y0 — где начинается текст на его первой странице; страницы раздела идут с шагом seg.g.P.
  let secIdx = 0;
  let seg = { y0: secs[0].g.top, g: secs[0].g };
  let lastBottom = seg.y0;
  let startOfSection = false;
  const boxes: PageBox[] = [];
  const closeSeg = (count: number) => {
    const top0 = seg.y0 - seg.g.top;
    const sg = secs[secIdx];
    for (let j = 0; j < count; j++) boxes.push({ top: top0 + j * seg.g.P, H: seg.g.H, W: sg.page.w * PX, page: sg.page, eff: sg.eff });
  };
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const { top: nTop, bottom: nBottom } = nat[i];
    if (it.sect) {
      // Разрыв раздела: следующий раздел начинается с новой страницы, у неё могут быть другие размер и поля.
      const pIdx = Math.max(0, Math.floor((lastBottom - seg.y0 - 0.5) / seg.g.P));
      closeSeg(pIdx + 1);
      const nextTop = seg.y0 - seg.g.top + (pIdx + 1) * seg.g.P;
      secIdx = Math.min(secIdx + 1, secs.length - 1);
      seg = { y0: nextTop + secs[secIdx].g.top, g: secs[secIdx].g };
      startOfSection = true;
      lastBottom = seg.y0;
      continue;
    }
    const h = nBottom - nTop;
    let pad = 0;
    const splits: { pos: number; h: number }[] = [];
    let forced = !it.row && !!it.node.attrs.pageBreakBefore;
    const heading = !it.row && it.node.type.name === 'heading';
    const region = (a: number) => {
      const p = Math.max(0, Math.floor((a - seg.y0 + 0.5) / seg.g.P));
      const start = p * seg.g.P + seg.y0;
      return { start, end: start + seg.g.C };
    };
    if (startOfSection) {
      // Первый блок раздела встаёт в начало текста новой страницы.
      pad = Math.max(0, seg.y0 - (nTop + acc));
      startOfSection = false;
      forced = false;
    }

    for (let guard = 0; guard < 80; guard++) {
      const a = nTop + acc + pad;
      const R = region(a);
      if (forced) {
        forced = false;
        if (a - R.start > 1) { pad += R.start + seg.g.P - a; continue; }
      }
      if (a >= R.end - 0.5) { pad += R.start + seg.g.P - a; continue; }
      if (a + h <= R.end + 0.5) {
        // Заголовок не остаётся один внизу страницы: за ним должны поместиться хотя бы две строки.
        if (heading && i + 1 < items.length && !items[i + 1].sect && a - R.start > 1) {
          const n = nat[i + 1];
          const need = n.top - nTop + Math.min(n.bottom - n.top, 2 * h);
          if (a + need > R.end + 0.5) { pad += R.start + seg.g.P - a; continue; }
        }
        break;
      }
      // Не помещается.
      if (it.row) {
        if (a - R.start > 1) { pad += R.start + seg.g.P - a; continue; }
        break; // строка выше страницы — не рвём
      }
      const lines = lineBoxes(it, i);
      const n = lines.length;
      let k = 0;
      while (k < n && lines[k].bottom + acc + pad <= R.end + 0.5) k++;
      if (k >= n) break;
      const atTop = a - R.start <= 1;
      if (k < 2 && n >= 2 && !atTop) { pad += R.start + seg.g.P - a; continue; }
      if (n - k < 2) {
        if (k >= 3) k -= 1;
        else if (!atTop) { pad += R.start + seg.g.P - a; continue; }
      }
      k = Math.max(1, k);
      let local = acc + pad;
      let end = R.end;
      let next = R.start + seg.g.P;
      let j = k;
      for (let safety = 0; safety < 200 && j < n; safety++) {
        const spacer = next - (lines[j].top + local);
        splits.push({ pos: lines[j].pos, h: spacer });
        local += spacer;
        end += seg.g.P;
        next += seg.g.P;
        let m = j;
        while (m < n && lines[m].bottom + local <= end + 0.5) m++;
        if (m >= n) break;
        if (n - m < 2 && m - 1 >= j + 2) m -= 1;
        j = Math.max(m, j + 1);
      }
      break;
    }

    if (pad > 0.5) out.push({ kind: it.row ? 'row' : 'pad', pos: it.pos, end: it.end, h: r1(pad), cols: it.cols, padPos: it.padPos, padEnd: it.padEnd });
    for (const s of splits) out.push({ kind: 'split', pos: s.pos, end: s.pos, h: r1(s.h), cols: 0, padPos: s.pos, padEnd: s.pos });
    acc += r1(pad) + splits.reduce((s, x) => s + r1(x.h), 0);
    lastBottom = nBottom + acc;
  }
  closeSeg(Math.max(1, Math.floor((lastBottom - seg.y0 - 0.5) / seg.g.P) + 1));
  return { out, pages: boxes.length, boxes, items: items.length };
}

function toDecos(out: Planned[]): Decoration[] {
  return out.map((p) => {
    if (p.kind === 'pad') return Decoration.node(p.padPos, p.padEnd, { class: 'pg-pad', style: `padding-top:${p.h}px` }, { pg: 'pad', h: p.h });
    if (p.kind === 'row') return Decoration.widget(p.pos, () => rowDom(p.h, p.cols), { side: -1, key: `row${p.h}:${p.cols}`, pg: 'row', h: p.h });
    return Decoration.widget(p.pos, () => splitDom(p.h), { side: -1, key: `split${p.h}`, pg: 'split', h: p.h });
  });
}

const sig = (list: { from: number; spec: Spec }[]) => list.map((d) => `${d.spec.pg}:${d.from}:${d.spec.h}`).join('|');

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
    out.push(Decoration.node(off, off + node.nodeSize, { style: `margin-left:${side + p.left}pt;margin-right:${side + p.right}pt` }));
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
      init: () => ({ deco: DecorationSet.empty, pages: 1, boxes: [], rev: 0 }),
      apply(tr, prev) {
        const meta = tr.getMeta(pagKey) as { deco?: DecorationSet; pages?: number; boxes?: PageBox[]; replan?: boolean } | undefined;
        if (meta?.deco) return { deco: meta.deco, pages: meta.pages ?? prev.pages, boxes: meta.boxes ?? prev.boxes, rev: prev.rev + 1 };
        if (tr.docChanged) return { ...prev, deco: prev.deco.map(tr.mapping, tr.doc), rev: prev.rev + 1 };
        if (meta?.replan) return { ...prev, rev: prev.rev + 1 };
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
          marginCache = { doc: state.doc, base: key, list: marginDecos(state.doc, base) };
        }
        // add() и create() портят переданный массив — отдаём копию.
        return marginCache.list ? st.deco.add(state.doc, [...marginCache.list]) : st.deco;
      },
    },
    view(view) {
      let seen = -1;
      let passes = 0;
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
        const secs: SecGeo[] = sectionsOf(view.state.doc, getBase()).map((x) => { const e = eff(x.page); return { g: geometry(e), page: x.page, eff: e }; });
        const t0 = performance.now();
        const { out, pages, boxes, items } = plan(view, secs);
        big = items > 600;
        stats.ms = performance.now() - t0;
        stats.items = items;
        const old = st.deco.find().map((d) => ({ from: d.from, spec: d.spec as Spec }));
        // DecorationSet.create портит переданный массив: набор строим один раз.
        const set = DecorationSet.create(view.state.doc, toDecos(out));
        const fresh = set.find().map((d) => ({ from: d.from, spec: d.spec as Spec }));
        const boxSig = boxes.map((b) => `${Math.round(b.top)}:${b.W}:${b.H}`).join('|');
        const stateSig = st.boxes.map((b) => `${Math.round(b.top)}:${b.W}:${b.H}`).join('|');
        if (boxSig !== lastSig) { lastSig = boxSig; onPages(pages, boxes); }
        const same = sig(old.sort((a, b) => a.from - b.from)) === sig(fresh.sort((a, b) => a.from - b.from));
        if (same && boxSig === stateSig) { passes = 0; return; }
        if (++passes > 6) return; // защита от зацикливания
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
          const rev = pagKey.getState(v.state)?.rev ?? 0;
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
  view.dispatch(view.state.tr.setMeta(pagKey, { replan: true }).setMeta('addToHistory', false));
}

/** Время последнего расчёта (для проверок). */
export const stats = { ms: 0, items: 0 };

export const pageCount = (view: EditorView): number => pagKey.getState(view.state)?.pages ?? 1;
export const pageBoxes = (view: EditorView): PageBox[] => pagKey.getState(view.state)?.boxes ?? [];
