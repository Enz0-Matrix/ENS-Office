// Линейки сверху и слева: поля страницы, отступы абзаца (первая строка, выступ, слева, справа), границы столбцов таблицы.
// Ноль — у границы текста; единица — сантиметр; масштаб учитывается. Показывают страницу, на которой стоит курсор.

import type { Node as PMNode } from 'prosemirror-model';
import { TableMap } from 'prosemirror-tables';
import type { EditorView } from 'prosemirror-view';
import type { PageSetup } from './docx';
import { schema } from './schema';
import { columnWidths, setColumnWidths, tableAt } from './tables-ui';

export interface RulerHost {
  view: () => EditorView;
  /** Листы-подложки (по ним берётся положение страницы на экране). */
  papers: () => HTMLElement;
  zoom: () => number;
  /** Номер текущей страницы с нуля и её параметры (поля без учёта колонтитулов). */
  pageIndex: () => number;
  page: () => PageSetup;
  changePage: (next: Partial<PageSetup>) => void;
  locked: () => boolean;
}

const PX_CM = 96 / 2.54;
const PX_PT = 4 / 3;
/** Толщина линеек, px. */
const RH = 30;
const el = <K extends keyof HTMLElementTagNameMap>(tag: K, cls: string): HTMLElementTagNameMap[K] => { const e = document.createElement(tag); e.className = cls; return e; };

let H: RulerHost;
let grid: HTMLElement, hr: HTMLElement, vr: HTMLElement, hin: HTMLElement, vin: HTMLElement, tip: HTMLElement;
let hsvg: SVGSVGElement, vsvg: SVGSVGElement;
let mLeft: HTMLElement, mRight: HTMLElement, mTop: HTMLElement, mBottom: HTMLElement;
let iFirst: HTMLElement, iHang: HTMLElement, iLeft: HTMLElement, iRight: HTMLElement;
let colMarks: HTMLElement[] = [];
let rowMarks: HTMLElement[] = [];
let visible = true;
let raf = 0;
let dragging = false;
let guide: HTMLElement | null = null;

const KEY = 'ens-text.ruler';
export const rulerVisible = (): boolean => visible;

export function setRulerVisible(on: boolean): void {
  visible = on;
  grid.classList.toggle('no-ruler', !on);
  try { localStorage.setItem(KEY, on ? '1' : '0'); } catch { /* без хранилища — до закрытия окна */ }
  updateRuler();
}

const cmText = (pt: number) => `${(Math.round((pt / 28.3465) * 100) / 100).toString().replace('.', ',')} см`;

interface ParaInfo { L: number; F: number; R: number }

/** Абзац под курсором, если он обычный (не в списке и не в таблице): его отступы. */
function paraAtCursor(): { info: ParaInfo; node: PMNode } | null {
  const v = H.view();
  const $f = v.state.selection.$from;
  const node = $f.parent;
  if (!node.isTextblock || (node.type !== schema.nodes.paragraph && node.type !== schema.nodes.heading)) return null;
  for (let d = $f.depth; d > 0; d--) if ($f.node(d).type === schema.nodes.list_item) return null;
  return { node, info: { L: (node.attrs.indent as number | null) ?? 0, F: (node.attrs.first as number | null) ?? 0, R: (node.attrs.rindent as number | null) ?? 0 } };
}

/** Применить отступы ко всем абзацам выделения. */
function applyIndents(next: ParaInfo): void {
  const v = H.view();
  const { from, to } = v.state.selection;
  const tr = v.state.tr;
  v.state.doc.nodesBetween(from, to, (n, pos) => {
    if (n.type !== schema.nodes.paragraph && n.type !== schema.nodes.heading) return true;
    const $p = v.state.doc.resolve(pos);
    for (let d = $p.depth; d > 0; d--) if ($p.node(d).type === schema.nodes.list_item) return false;
    const r = (x: number) => Math.round(x * 10) / 10 || null;
    tr.setNodeMarkup(pos, undefined, { ...n.attrs, indent: r(next.L), first: r(next.F), rindent: r(next.R) });
    return false;
  });
  v.dispatch(tr);
}

function showTip(x: number, y: number, text: string): void {
  tip.hidden = false;
  tip.textContent = text;
  tip.style.left = `${x + 14}px`;
  tip.style.top = `${y + 14}px`;
}
const hideTip = () => { tip.hidden = true; guide?.remove(); guide = null; };

function showGuide(rect: DOMRect, x: number | null, y: number | null): void {
  if (!guide) { guide = el('div', 'ruler-guide'); document.body.append(guide); }
  guide.style.left = `${x ?? rect.left}px`;
  guide.style.top = `${y ?? rect.top}px`;
  guide.style.width = x != null ? '0' : `${rect.width}px`;
  guide.style.height = y != null ? '0' : `${rect.height}px`;
}

/** Перетаскивание метки: move получает смещение в экранных px и координаты мыши; up — итог. */
function track(e: MouseEvent, move: (dx: number, dy: number, ev: MouseEvent) => void, up: (dx: number, dy: number) => void): void {
  e.preventDefault();
  e.stopPropagation();
  dragging = true;
  const sx = e.clientX, sy = e.clientY;
  let dx = 0, dy = 0;
  const mv = (ev: MouseEvent) => { dx = ev.clientX - sx; dy = ev.clientY - sy; move(dx, dy, ev); };
  const end = () => { document.removeEventListener('mousemove', mv); dragging = false; hideTip(); up(dx, dy); updateRuler(); };
  document.addEventListener('mousemove', mv);
  document.addEventListener('mouseup', end, { once: true });
}

function ticks(len: number, zero: number, vertical: boolean): string {
  const z = H.zoom();
  const cm = PX_CM * z;
  const parts: string[] = [];
  const low = Math.floor(-zero / cm) - 1, high = Math.ceil((len - zero) / cm) + 1;
  const put = (pos: number, size: number) => {
    if (pos < 0 || pos > len) return;
    const [x1, y1, x2, y2] = vertical ? [RH - 2 - size, pos, RH - 2, pos] : [pos, RH - 2 - size, pos, RH - 2];
    parts.push(`<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}"/>`);
  };
  for (let k = low; k <= high; k++) {
    const p = zero + k * cm;
    put(p, 8);
    if (k !== 0 && p >= 0 && p <= len) {
      const n = Math.abs(k);
      parts.push(vertical ? `<text x="5" y="${p + 3}" font-size="9">${n}</text>` : `<text x="${p + 2}" y="${RH - 12}" font-size="9">${n}</text>`);
    }
    put(p + cm / 2, 5);
    if (cm >= 30) for (const q of [1, 2, 3, 4, 6, 7, 8, 9]) put(p + (cm * q) / 10, 2);
  }
  return parts.join('');
}

function doUpdate(): void {
  raf = 0;
  if (!visible || !H) return;
  const paper = H.papers().children[H.pageIndex()] as HTMLElement | undefined;
  if (!paper) return;
  const pr = paper.getBoundingClientRect();
  const hrr = hr.getBoundingClientRect();
  const vrr = vr.getBoundingClientRect();
  const z = H.zoom();
  const page = H.page();
  const ml = page.left * PX_PT * z, mr = page.right * PX_PT * z, mt = page.top * PX_PT * z, mb = page.bottom * PX_PT * z;
  const W = pr.width, Hh = pr.height;

  hin.style.left = `${pr.left - hrr.left}px`;
  hin.style.width = `${W}px`;
  hsvg.setAttribute('width', String(W));
  hsvg.setAttribute('height', String(hrr.height));
  hsvg.innerHTML = `<rect x="0" y="0" width="${ml}" height="${RH}" fill="#c5c9cf"/><rect x="${W - mr}" y="0" width="${mr}" height="${RH}" fill="#c5c9cf"/>` +
    `<rect x="${ml}" y="0" width="${Math.max(0, W - ml - mr)}" height="${RH}" fill="#fff"/><g stroke="#555" fill="#555" stroke-width="1">${ticks(W, ml, false)}</g>`;
  mLeft.style.left = `${ml}px`;
  mRight.style.left = `${W - mr}px`;

  vin.style.top = `${pr.top - vrr.top}px`;
  vin.style.height = `${Hh}px`;
  vsvg.setAttribute('width', String(vrr.width));
  vsvg.setAttribute('height', String(Hh));
  vsvg.innerHTML = `<rect x="0" y="0" width="${RH}" height="${mt}" fill="#c5c9cf"/><rect x="0" y="${Hh - mb}" width="${RH}" height="${mb}" fill="#c5c9cf"/>` +
    `<rect x="0" y="${mt}" width="${RH}" height="${Math.max(0, Hh - mt - mb)}" fill="#fff"/><g stroke="#555" fill="#555" stroke-width="1">${ticks(Hh, mt, true)}</g>`;
  mTop.style.top = `${mt}px`;
  mBottom.style.top = `${Hh - mb}px`;

  // Метки отступов абзаца или границы столбцов таблицы.
  const tbl = !H.locked() ? tableAt(H.view().state) : null;
  // В таблице вместо отступов абзаца — границы столбцов и строк.
  const para = !H.locked() && !tbl ? paraAtCursor() : null;
  for (const m of [iFirst, iHang, iLeft, iRight]) m.hidden = !para;
  if (para) {
    const { L, F, R } = para.info;
    iFirst.style.left = `${ml + (L + F) * PX_PT * z}px`;
    iHang.style.left = `${ml + L * PX_PT * z}px`;
    iLeft.style.left = `${ml + L * PX_PT * z}px`;
    iRight.style.left = `${W - mr - R * PX_PT * z}px`;
  }
  for (const m of colMarks) m.remove();
  for (const m of rowMarks) m.remove();
  colMarks = [];
  rowMarks = [];
  if (tbl) {
    const { edgesX, edgesY } = tableEdges(tbl, pr);
    edgesX.forEach((x, i) => {
      const m = el('div', 'rc');
      m.style.left = `${x}px`;
      m.title = 'Граница столбца: перетащите, чтобы изменить ширину';
      m.addEventListener('mousedown', (e) => dragColumn(e, i, tbl.pos, z));
      hin.append(m);
      colMarks.push(m);
    });
    edgesY.forEach((y, i) => {
      const m = el('div', 'rr');
      m.style.top = `${y}px`;
      m.title = 'Граница строки: перетащите, чтобы изменить высоту';
      m.addEventListener('mousedown', (e) => dragRow(e, i, tbl.pos, z));
      vin.append(m);
      rowMarks.push(m);
    });
  }
}

/** Границы столбцов (x) и строк (y) таблицы под курсором в координатах листа на экране. */
function tableEdges(tbl: { pos: number; node: PMNode }, pr: DOMRect): { edgesX: number[]; edgesY: number[] } {
  const v = H.view();
  const map = TableMap.get(tbl.node);
  const start = tbl.pos + 1;
  const xs: number[] = [];
  const ys: number[] = [];
  // Первая строка — границы столбцов; первый столбец — границы строк.
  const edges = (cellPos: number, axis: 'x' | 'y', end: 'a' | 'b') => {
    const dom = v.nodeDOM(start + cellPos) as HTMLElement | null;
    if (!dom) return null;
    const r = dom.getBoundingClientRect();
    return axis === 'x' ? (end === 'a' ? r.left : r.right) - pr.left : (end === 'a' ? r.top : r.bottom) - pr.top;
  };
  const first = edges(map.map[0], 'x', 'a');
  if (first != null) xs.push(first);
  const seen = new Set<number>();
  for (let c = 0; c < map.width; c++) {
    const pos = map.map[c];
    if (seen.has(pos)) continue;
    seen.add(pos);
    const x = edges(pos, 'x', 'b');
    if (x != null) xs.push(x);
  }
  const seenR = new Set<number>();
  const y0 = edges(map.map[0], 'y', 'a');
  if (y0 != null) ys.push(y0);
  for (let r = 0; r < map.height; r++) {
    const pos = map.map[r * map.width];
    if (seenR.has(pos)) continue;
    seenR.add(pos);
    const y = edges(pos, 'y', 'b');
    if (y != null) ys.push(y);
  }
  return { edgesX: xs, edgesY: ys };
}

function dragColumn(e: MouseEvent, edge: number, tablePos: number, z: number): void {
  const v = H.view();
  const cols = columnWidths(v, tablePos);
  if (!cols.length) return;
  const orig = [...cols];
  const hrr = hin.getBoundingClientRect();
  track(e, (dx, _dy, ev) => {
    const d = dx / z;
    const next = [...orig];
    if (edge === 0) return; // левый край таблицы не двигаем
    const i = edge - 1;
    const w = Math.max(20, orig[i] + d);
    next[i] = w;
    if (edge < orig.length) next[edge] = Math.max(20, orig[edge] - (w - orig[i]));
    showGuide(hrr, ev.clientX, null);
    showTip(ev.clientX, ev.clientY, `${Math.round(next[i])} px`);
  }, (dx) => {
    const d = dx / z;
    if (!d || edge === 0) return;
    const i = edge - 1;
    const next = [...orig];
    const w = Math.max(20, orig[i] + d);
    next[i] = w;
    if (edge < orig.length) next[edge] = Math.max(20, orig[edge] - (w - orig[i]));
    setColumnWidths(H.view(), tablePos, next);
  });
}

function dragRow(e: MouseEvent, edge: number, tablePos: number, z: number): void {
  if (edge === 0) return;
  const v = H.view();
  const tbl = v.state.doc.nodeAt(tablePos);
  if (!tbl) return;
  const rowIdx = edge - 1;
  const row = tbl.child(rowIdx);
  const dom = v.nodeDOM(tablePos + 1 + (() => { let o = 0; for (let i = 0; i < rowIdx; i++) o += tbl.child(i).nodeSize; return o; })()) as HTMLElement | null;
  const h0 = dom ? dom.getBoundingClientRect().height / z : 20;
  track(e, (dx, dy, ev) => {
    showGuide(vin.getBoundingClientRect(), null, ev.clientY);
    showTip(ev.clientX, ev.clientY, cmText((h0 + dy / z) / PX_PT));
  }, (_dx, dy) => {
    const h = Math.max(12, h0 + dy / z);
    let off = 0;
    for (let i = 0; i < rowIdx; i++) off += tbl.child(i).nodeSize;
    const pos = tablePos + 1 + off;
    const cur = H.view().state.doc.nodeAt(pos);
    if (!cur) return;
    H.view().dispatch(H.view().state.tr.setNodeMarkup(pos, undefined, { ...cur.attrs, height: Math.round((h / PX_PT) * 10) / 10 }));
    void row;
  });
}

export function updateRuler(): void {
  if (raf) return;
  raf = requestAnimationFrame(doUpdate);
}

export function initRuler(host: RulerHost): void {
  H = host;
  grid = document.getElementById('deskgrid')!;
  hr = document.getElementById('hruler')!;
  vr = document.getElementById('vruler')!;
  hin = document.getElementById('hr-inner')!;
  vin = document.getElementById('vr-inner')!;
  tip = document.getElementById('ruler-tip')!;
  hsvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  vsvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  hin.append(hsvg);
  vin.append(vsvg);
  mLeft = el('div', 'rm'); mRight = el('div', 'rm'); mTop = el('div', 'rm'); mBottom = el('div', 'rm');
  mLeft.title = 'Левое поле'; mRight.title = 'Правое поле'; mTop.title = 'Верхнее поле'; mBottom.title = 'Нижнее поле';
  hin.append(mLeft, mRight);
  vin.append(mTop, mBottom);
  iFirst = el('div', 'ri ri-first'); iHang = el('div', 'ri ri-hang'); iLeft = el('div', 'ri ri-left'); iRight = el('div', 'ri ri-right');
  iFirst.title = 'Отступ первой строки'; iHang.title = 'Выступ (отступ остальных строк)'; iLeft.title = 'Отступ слева'; iRight.title = 'Отступ справа';
  hin.append(iFirst, iHang, iLeft, iRight);

  // Поля страницы: тянем границу серой и белой частей.
  const marginDrag = (m: HTMLElement, key: 'left' | 'right' | 'top' | 'bottom') => m.addEventListener('mousedown', (e) => {
    const page = H.page();
    const z = H.zoom();
    const paper = H.papers().children[H.pageIndex()] as HTMLElement;
    const pr = paper.getBoundingClientRect();
    const horizontal = key === 'left' || key === 'right';
    const start = page[key];
    const limit = horizontal ? (page.w - 72 - (key === 'left' ? page.right : page.left)) : (page.h - 72 - (key === 'top' ? page.bottom : page.top));
    const value = (d: number) => Math.min(limit, Math.max(0, Math.round((start + ((key === 'left' || key === 'top' ? 1 : -1) * d) / (PX_PT * z)) * 2) / 2));
    track(e, (dx, dy, ev) => {
      const val = value(horizontal ? dx : dy);
      const pos = (key === 'left' ? pr.left + val * PX_PT * z : key === 'right' ? pr.right - val * PX_PT * z : key === 'top' ? pr.top + val * PX_PT * z : pr.bottom - val * PX_PT * z);
      showGuide(pr, horizontal ? pos : null, horizontal ? null : pos);
      showTip(ev.clientX, ev.clientY, cmText(val));
    }, (dx, dy) => {
      const d = horizontal ? dx : dy;
      if (Math.abs(d) < 1) return;
      H.changePage({ [key]: value(d) } as Partial<PageSetup>);
    });
  });
  marginDrag(mLeft, 'left'); marginDrag(mRight, 'right'); marginDrag(mTop, 'top'); marginDrag(mBottom, 'bottom');

  // Отступы абзаца.
  const indentDrag = (m: HTMLElement, kind: 'first' | 'hang' | 'left' | 'right') => m.addEventListener('mousedown', (e) => {
    const p = paraAtCursor();
    if (!p) return;
    const z = H.zoom();
    const page = H.page();
    const paper = H.papers().children[H.pageIndex()] as HTMLElement;
    const pr = paper.getBoundingClientRect();
    const textW = page.w - page.left - page.right;
    const { L, F, R } = p.info;
    const calc = (dx: number): ParaInfo => {
      const d = dx / (PX_PT * z);
      const snap = (x: number) => Math.round(x * 2) / 2;
      if (kind === 'first') return { L, F: snap(Math.max(-L, Math.min(textW - 20 - L, F + d))), R };
      if (kind === 'hang') { const nl = snap(Math.max(0, Math.min(textW - 20, L + d))); return { L: nl, F: F - (nl - L), R }; }
      if (kind === 'left') return { L: snap(Math.max(0, Math.min(textW - 20, L + d))), F, R };
      return { L, F, R: snap(Math.max(0, Math.min(textW - 20 - L, R - d))) };
    };
    track(e, (dx, _dy, ev) => {
      const n = calc(dx);
      const x = kind === 'right' ? pr.right - (page.right + n.R) * PX_PT * z : pr.left + (page.left + (kind === 'first' ? n.L + n.F : n.L)) * PX_PT * z;
      showGuide(pr, x, null);
      showTip(ev.clientX, ev.clientY, cmText(kind === 'first' ? n.L + n.F : kind === 'right' ? n.R : n.L));
    }, (dx) => { if (Math.abs(dx) >= 1) applyIndents(calc(dx)); });
  });
  indentDrag(iFirst, 'first'); indentDrag(iHang, 'hang'); indentDrag(iLeft, 'left'); indentDrag(iRight, 'right');

  document.getElementById('desk')!.addEventListener('scroll', updateRuler);
  window.addEventListener('resize', updateRuler);
  try { visible = localStorage.getItem(KEY) !== '0'; } catch { visible = true; }
  grid.classList.toggle('no-ruler', !visible);
  void dragging;
  updateRuler();
}
