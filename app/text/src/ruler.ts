// Линейки сверху и слева: поля страницы, отступы абзаца (первая строка, выступ, слева, справа), границы столбцов таблицы.
// Ноль — у границы текста; единица — сантиметр; масштаб учитывается. Показывают страницу, на которой стоит курсор.
// Перетаскивание «живое»: метка, поля, отступы, столбцы и строки меняются на экране за мышью (декорации, без шагов
// в документе), а в документ изменение записывается один раз при отпускании — одна запись в истории отмены.

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey } from 'prosemirror-state';
import { TableMap } from 'prosemirror-tables';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
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
  /** Показать поля страницы до записи (null — вернуть как в документе). */
  previewPage?: (next: Partial<PageSetup> | null) => void;
  locked: () => boolean;
}

// ---------------------------------------------------------------- предпросмотр перетаскивания

const previewKey = new PluginKey<DecorationSet>('ruler-preview');

/** Декорации предпросмотра: пока тянем метку, абзацы и строки таблицы показываются с новыми размерами. */
export const rulerPreviewPlugin = new Plugin<DecorationSet>({
  key: previewKey,
  state: {
    init: () => DecorationSet.empty,
    apply: (tr, set) => {
      const m = tr.getMeta(previewKey) as DecorationSet | null | undefined;
      if (m !== undefined) return m ?? DecorationSet.empty;
      return tr.docChanged ? DecorationSet.empty : set;
    },
  },
  props: { decorations: (st) => previewKey.getState(st) },
});

/** Обновить предпросмотр без dispatch: в историю и в «изменено» это не попадает. */
function preview(decos: Decoration[] | null): void {
  const v = H.view();
  if (!previewKey.getState(v.state)) return; // плагина нет (окно колонтитула) — без предпросмотра
  const tr = v.state.tr.setMeta(previewKey, decos ? DecorationSet.create(v.state.doc, decos) : null).setMeta('addToHistory', false);
  v.updateState(v.state.apply(tr));
}

/** Временные значения для самой линейки во время перетаскивания. */
let livePage: Partial<PageSetup> | null = null;
let liveIndent: ParaInfo | null = null

const PX_CM = 96 / 2.54;
const PX_PT = 4 / 3;
/** Толщина линеек, px. */
const RH = 22;
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
  const end = () => {
    document.removeEventListener('mousemove', mv);
    dragging = false;
    hideTip();
    livePage = null;
    liveIndent = null;
    // Сначала снять предпросмотр, затем записать итог одной транзакцией.
    preview(null);
    up(dx, dy);
    updateRuler();
  };
  document.addEventListener('mousemove', mv);
  document.addEventListener('mouseup', end, { once: true });
}

function ticks(len: number, zero: number, vertical: boolean): string {
  const z = H.zoom();
  const cm = PX_CM * z;
  const parts: string[] = [];
  const low = Math.floor(-zero / cm) - 1, high = Math.ceil((len - zero) / cm) + 1;
  // Деления на целых пикселях (+0,5): линия в 1 px остаётся чёткой, а не размазывается на два.
  const px = (v: number) => Math.round(v) + 0.5;
  const put = (pos: number, size: number) => {
    if (pos < 0 || pos > len) return;
    const q = px(pos);
    const [x1, y1, x2, y2] = vertical ? [RH - 1 - size, q, RH - 1, q] : [q, RH - 1 - size, q, RH - 1];
    parts.push(`<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}"/>`);
  };
  for (let k = low; k <= high; k++) {
    const p = zero + k * cm;
    if (k !== 0 && p >= 0 && p <= len) {
      // Цифра сантиметра вместо длинного деления, как в Word.
      const n = Math.abs(k);
      parts.push(vertical ? `<text x="${RH / 2}" y="${p + 3}" text-anchor="middle" font-size="8.5" stroke="none">${n}</text>` : `<text x="${p}" y="${RH - 8}" text-anchor="middle" font-size="8.5" stroke="none">${n}</text>`);
    } else put(p, 6);
    put(p + cm / 2, 4);
    // Миллиметры — при масштабе от ~65 %.
    if (cm >= 24) for (const q of [1, 2, 3, 4, 6, 7, 8, 9]) put(p + (cm * q) / 10, 2);
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
  const page = { ...H.page(), ...livePage };
  const ml = page.left * PX_PT * z, mr = page.right * PX_PT * z, mt = page.top * PX_PT * z, mb = page.bottom * PX_PT * z;
  const W = pr.width, Hh = pr.height;

  hin.style.left = `${pr.left - hrr.left}px`;
  hin.style.width = `${W}px`;
  hsvg.setAttribute('width', String(W));
  hsvg.setAttribute('height', String(hrr.height));
  hsvg.innerHTML = `<rect x="0" y="0" width="${ml}" height="${RH}" fill="#c5c9cf"/><rect x="${W - mr}" y="0" width="${mr}" height="${RH}" fill="#c5c9cf"/>` +
    `<rect x="${ml}" y="0" width="${Math.max(0, W - ml - mr)}" height="${RH}" fill="#fff"/><g stroke="#666" fill="#444" stroke-width="1" shape-rendering="crispEdges" font-family="Segoe UI, Carlito, Arimo, sans-serif">${ticks(W, ml, false)}</g>`;
  mLeft.style.left = `${ml}px`;
  mRight.style.left = `${W - mr}px`;
  // Поля хранятся в пунктах текущего раздела, экранная шкала — в сантиметрах.
  for (const [handle, key] of [[mLeft, 'left'], [mRight, 'right'], [mTop, 'top'], [mBottom, 'bottom']] as const) {
    const horizontal = key === 'left' || key === 'right';
    const max = Math.max(0, (horizontal ? page.w - page.left - page.right : page.h - page.top - page.bottom) + page[key] - 72);
    handle.setAttribute('aria-valuemax', String(max));
    handle.setAttribute('aria-valuenow', String(page[key]));
    handle.setAttribute('aria-valuetext', cmText(page[key]));
    handle.setAttribute('aria-disabled', String(H.locked()));
  }

  vin.style.top = `${pr.top - vrr.top}px`;
  vin.style.height = `${Hh}px`;
  vsvg.setAttribute('width', String(vrr.width));
  vsvg.setAttribute('height', String(Hh));
  vsvg.innerHTML = `<rect x="0" y="0" width="${RH}" height="${mt}" fill="#c5c9cf"/><rect x="0" y="${Hh - mb}" width="${RH}" height="${mb}" fill="#c5c9cf"/>` +
    `<rect x="0" y="${mt}" width="${RH}" height="${Math.max(0, Hh - mt - mb)}" fill="#fff"/><g stroke="#666" fill="#444" stroke-width="1" shape-rendering="crispEdges" font-family="Segoe UI, Carlito, Arimo, sans-serif">${ticks(Hh, mt, true)}</g>`;
  mTop.style.top = `${mt}px`;
  mBottom.style.top = `${Hh - mb}px`;

  // Метки отступов абзаца или границы столбцов таблицы.
  const tbl = !H.locked() ? tableAt(H.view().state) : null;
  // В таблице вместо отступов абзаца — границы столбцов и строк.
  const para = !H.locked() && !tbl ? paraAtCursor() : null;
  for (const m of [iFirst, iHang, iLeft, iRight]) m.hidden = !para;
  if (para) {
    const { L, F, R } = liveIndent ?? para.info;
    iFirst.style.left = `${ml + (L + F) * PX_PT * z}px`;
    iHang.style.left = `${ml + L * PX_PT * z}px`;
    iLeft.style.left = `${ml + L * PX_PT * z}px`;
    iRight.style.left = `${W - mr - R * PX_PT * z}px`;
  }
  if (dragging) return; // метки столбцов и строк не пересоздаём под мышью
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
  const mark = colMarks[edge];
  const mark0 = mark ? parseFloat(mark.style.left) : 0;
  // Ширины столбцов на экране: <col> внутри colgroup (вид таблицы не реагирует на их правку), ширина таблицы — сумма.
  const tdom = (v.nodeDOM(tablePos) as HTMLElement | null)?.querySelector('table') ?? null;
  const colEls = tdom ? [...tdom.querySelectorAll<HTMLElement>(':scope > colgroup > col')] : [];
  const saved = { cols: colEls.map((c) => c.style.width), table: tdom?.style.width ?? '', minW: tdom?.style.minWidth ?? '' };
  const live = (next: number[] | null) => {
    if (!tdom || colEls.length !== orig.length) return;
    colEls.forEach((c, k) => { c.style.width = next ? `${next[k]}px` : saved.cols[k]; });
    tdom.style.width = next ? `${next.reduce((a, b) => a + b, 0)}px` : saved.table;
    tdom.style.minWidth = next ? '' : saved.minW;
  };
  track(e, (dx, _dy, ev) => {
    const d = dx / z;
    const next = [...orig];
    if (edge === 0) return; // левый край таблицы не двигаем
    const i = edge - 1;
    const w = Math.max(20, orig[i] + d);
    next[i] = w;
    if (edge < orig.length) next[edge] = Math.max(20, orig[edge] - (w - orig[i]));
    live(next);
    if (mark) mark.style.left = `${mark0 + (w - orig[i]) * z}px`;
    showGuide(hrr, ev.clientX, null);
    showTip(ev.clientX, ev.clientY, `${Math.round(next[i])} px`);
  }, (dx) => {
    live(null);
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
  const rowPos = tablePos + 1 + (() => { let o = 0; for (let i = 0; i < rowIdx; i++) o += tbl.child(i).nodeSize; return o; })();
  const mark = rowMarks[edge];
  const mark0 = mark ? parseFloat(mark.style.top) : 0;
  track(e, (dx, dy, ev) => {
    const h = Math.max(12, h0 + dy / z);
    preview([Decoration.node(rowPos, rowPos + row.nodeSize, { style: `height:${h}px` })]);
    if (mark) mark.style.top = `${mark0 + (h - h0) * z}px`;
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
  // Подсказки при наведении — что можно тянуть; aria-label остаётся коротким (по нему ищут и озвучивают).
  const hint = (m: HTMLElement, text: string) => { m.dataset.hint = text; };
  hint(mLeft, 'Левое поле — тяните границу серой и белой части (или стрелки ← →)');
  hint(mRight, 'Правое поле — тяните границу серой и белой части');
  hint(mTop, 'Верхнее поле — тяните границу'); hint(mBottom, 'Нижнее поле — тяните границу');
  hin.append(mLeft, mRight);
  vin.append(mTop, mBottom);
  iFirst = el('div', 'ri ri-first'); iHang = el('div', 'ri ri-hang'); iLeft = el('div', 'ri ri-left'); iRight = el('div', 'ri ri-right');
  iFirst.title = 'Отступ первой строки — тяните; можно увести в поле'; iHang.title = 'Выступ (остальные строки) — тяните'; iLeft.title = 'Отступ слева (весь абзац) — тяните; можно увести в поле'; iRight.title = 'Отступ справа — тяните';
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
      livePage = { [key]: val };
      H.previewPage?.(livePage);
      doUpdate();
    }, (dx, dy) => {
      H.previewPage?.(null);
      const d = horizontal ? dx : dy;
      if (Math.abs(d) < 1) return;
      H.changePage({ [key]: value(d) } as Partial<PageSetup>);
    });
  });
  marginDrag(mLeft, 'left'); marginDrag(mRight, 'right'); marginDrag(mTop, 'top'); marginDrag(mBottom, 'bottom');

  // Та же операция доступна без мыши: стрелка в сторону текста увеличивает поле.
  // Шаг — 1 мм (Shift — 5 мм): поля по ГОСТ задают в миллиметрах, линейка показывает сантиметры.
  const KEY_STEP_PT = 72 / 25.4;
  for (const [handle, key] of [[mLeft, 'left'], [mRight, 'right'], [mTop, 'top'], [mBottom, 'bottom']] as const) {
    handle.tabIndex = 0;
    handle.setAttribute('role', 'slider');
    handle.setAttribute('aria-label', handle.title);
    handle.title = handle.dataset.hint ?? handle.title;
    handle.setAttribute('aria-valuemin', '0');
    handle.addEventListener('keydown', (e) => {
      if (H.locked() || e.altKey || e.ctrlKey || e.metaKey) return;
      const inward = { left: 'ArrowRight', right: 'ArrowLeft', top: 'ArrowDown', bottom: 'ArrowUp' }[key];
      const outward = { left: 'ArrowLeft', right: 'ArrowRight', top: 'ArrowUp', bottom: 'ArrowDown' }[key];
      if (e.key !== inward && e.key !== outward) return;
      e.preventDefault();
      e.stopPropagation();
      const page = H.page();
      const max = (key === 'left' || key === 'right' ? page.w - page.left - page.right : page.h - page.top - page.bottom) + page[key] - 72;
      const delta = (e.shiftKey ? 5 : 1) * KEY_STEP_PT * (e.key === inward ? 1 : -1);
      const next = Math.max(0, Math.min(max, page[key] + delta));
      if (next !== page[key]) H.changePage({ [key]: next });
    });
  }

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
      // Как в Word: отступы можно увести в поле (на серую часть), но не дальше края листа.
      const minL = -page.left, minR = -page.right;
      if (kind === 'first') return { L, F: snap(Math.max(minL - L, Math.min(textW - 20 - L, F + d))), R };
      if (kind === 'hang') { const nl = snap(Math.max(minL, Math.min(textW - 20, L + d))); return { L: nl, F: Math.max(minL - nl, F - (nl - L)), R }; }
      if (kind === 'left') return { L: snap(Math.max(minL - Math.min(0, F), Math.min(textW - 20, L + d))), F, R };
      return { L, F, R: snap(Math.max(minR, Math.min(textW - 20 - L, R - d))) };
    };
    const v = H.view();
    const { from, to } = v.state.selection;
    const paras: { pos: number; end: number }[] = [];
    v.state.doc.nodesBetween(from, to, (n, pos) => {
      if (n.type !== schema.nodes.paragraph && n.type !== schema.nodes.heading) return true;
      const $p = v.state.doc.resolve(pos);
      for (let d = $p.depth; d > 0; d--) if ($p.node(d).type === schema.nodes.list_item) return false;
      paras.push({ pos, end: pos + n.nodeSize });
      return false;
    });
    track(e, (dx, _dy, ev) => {
      const n = calc(dx);
      liveIndent = n;
      preview(paras.map((q) => Decoration.node(q.pos, q.end, { style: `margin-left:${n.L}pt;text-indent:${n.F}pt;margin-right:${n.R}pt` })));
      doUpdate();
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
