// Печать и PDF (Файл → Печать, Ctrl+P). Лист превращается в HTML-таблицу
// с тем же оформлением и печатается системным окном: там и принтеры, и
// «Сохранить как PDF» (Windows: «Microsoft Print to PDF», РЕД ОС: «Печать в файл»).

import type { BorderItem, CellStyle, Model } from './engine';
import { fontStack } from './fonts';

export type Paper = 'A4' | 'A3' | 'A5' | 'Letter';
type Margins = 'normal' | 'narrow' | 'wide' | 'custom';

export interface PrintOptions {
  what: 'sheet' | 'selection';
  landscape: boolean;
  paper: Paper;
  fit: 'none' | 'width' | 'page';
  margins: Margins;
  /** Поля, мм: верх, право, низ, лево (для «Настраиваемые»). */
  marginMm: [number, number, number, number];
  gridlines: boolean;
  /** Заголовок (первая строка или закреплённые строки) — на каждой странице. */
  repeatHeader: boolean;
  /** По центру страницы по горизонтали. */
  centerH: boolean;
}

export const DEFAULT_PRINT: PrintOptions = {
  what: 'sheet', landscape: false, paper: 'A4', fit: 'width', margins: 'normal', marginMm: [19, 18, 19, 18],
  gridlines: false, repeatHeader: false, centerH: false,
};

interface Range { r1: number; c1: number; r2: number; c2: number }

export const PAPERS: Record<Paper, [number, number]> = { A4: [210, 297], A3: [297, 420], A5: [148, 210], Letter: [216, 279] };
// Как в Excel: обычные, узкие, широкие (мм: верх, право, низ, лево).
const MARGINS_MM: Record<Exclude<Margins, 'custom'>, [number, number, number, number]> = {
  normal: [19, 18, 19, 18], narrow: [19, 6, 19, 6], wide: [25, 25, 25, 25],
};

export function marginsOf(o: PrintOptions): [number, number, number, number] {
  return o.margins === 'custom' ? o.marginMm : MARGINS_MM[o.margins];
}

function paperOf(o: PrintOptions): [number, number] {
  const [w, h] = PAPERS[o.paper] ?? PAPERS.A4;
  return o.landscape ? [h, w] : [w, h];
}

/** Параметры листа из книги (Worksheet.ens_page_setup, из xlsx — pageSetup Excel). */
export function fromSaved(json: string): PrintOptions {
  const o: PrintOptions = { ...DEFAULT_PRINT };
  let v: Partial<PrintOptions> = {};
  try { v = json ? JSON.parse(json) : {}; } catch { /* по умолчанию */ }
  if (typeof v.landscape === 'boolean') o.landscape = v.landscape;
  if (v.paper && v.paper in PAPERS) o.paper = v.paper;
  if (v.fit === 'none' || v.fit === 'width' || v.fit === 'page') o.fit = v.fit;
  for (const k of ['gridlines', 'repeatHeader', 'centerH'] as const) if (typeof v[k] === 'boolean') o[k] = v[k];
  if (Array.isArray(v.marginMm) && v.marginMm.length === 4 && v.marginMm.every((x) => typeof x === 'number')) {
    o.marginMm = v.marginMm;
    const preset = (Object.keys(MARGINS_MM) as (keyof typeof MARGINS_MM)[])
      .find((k) => MARGINS_MM[k].every((x, i) => Math.abs(x - o.marginMm[i]) < 1.5));
    o.margins = preset ?? 'custom';
  }
  return o;
}

export function toSaved(o: PrintOptions): string {
  const { landscape, paper, fit, gridlines, repeatHeader, centerH } = o;
  return JSON.stringify({ landscape, paper, fit, marginMm: marginsOf(o), gridlines, repeatHeader, centerH });
}
const MAX_CELLS = 400_000;
const ROW_SCALE = 22 / 25;

function borderCss(b: BorderItem | undefined, model: Model): string {
  if (!b?.style) return '';
  const color = b.color ? model.resolveColor(b.color) || '#000' : '#000';
  const st = b.style;
  if (st === 'double') return `3px double ${color}`;
  const width = st === 'thick' ? 3 : st.startsWith('medium') ? 2 : 1;
  const kind = st === 'dotted' ? 'dotted' : st.includes('dash') ? 'dashed' : 'solid';
  return `${width}px ${kind} ${color}`;
}

function cellCss(st: CellStyle | null, model: Model, type: number): string {
  const css: string[] = [];
  const f = st?.font;
  if (f) {
    css.push(`font: ${f.i ? 'italic ' : ''}${f.b ? 'bold ' : ''}${Math.round((f.sz ?? 11) * 1.33)}px ${fontStack(f.name)}`);
    if (f.color) css.push(`color: ${model.resolveColor(f.color)}`);
    const deco = [f.u ? 'underline' : '', f.strike ? 'line-through' : ''].filter(Boolean).join(' ');
    if (deco) css.push(`text-decoration: ${deco}`);
  }
  if (st?.fill?.color) css.push(`background: ${model.resolveColor(st.fill.color)}`);
  let h = st?.alignment?.horizontal ?? 'general';
  if (h === 'general') h = type === 1 ? 'right' : type === 4 || type === 16 ? 'center' : 'left';
  if (h === 'centercontinuous') h = 'center';
  css.push(`text-align: ${['left', 'right', 'center', 'justify'].includes(h) ? h : 'left'}`);
  const v = st?.alignment?.vertical ?? 'bottom';
  css.push(`vertical-align: ${v === 'top' ? 'top' : v === 'center' ? 'middle' : 'bottom'}`);
  if (st?.alignment?.wrap_text) css.push('white-space: pre-wrap');
  const b = st?.border;
  for (const side of ['top', 'right', 'bottom', 'left'] as const) {
    const v2 = borderCss(b?.[side], model);
    if (v2) css.push(`border-${side}: ${v2}`);
  }
  return css.join('; ');
}

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

/** Что печатаем: выделение или занятая часть листа. */
export function printRange(model: Model, sheet: number, opts: PrintOptions, sel: Range): Range | null {
  if (opts.what === 'selection' && (sel.r1 !== sel.r2 || sel.c1 !== sel.c2)) {
    return { r1: sel.r1, c1: sel.c1, r2: Math.min(sel.r2, 1_048_576), c2: Math.min(sel.c2, 16_384) };
  }
  const [r1, c1, r2, c2] = model.getSheetDimension(sheet);
  if (r1 === r2 && c1 === c2 && model.getFormattedCellValue(sheet, r1, c1) === '') return null;
  return { r1: 1, c1: 1, r2, c2 };
}

/** Таблица для печати и предпросмотра. Возвращает её ширину и высоту в px. */
export function buildTable(model: Model, sheet: number, r: Range, opts: PrintOptions): { html: string; width: number; height: number } {
  if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) > MAX_CELLS) {
    throw new Error('Слишком большая область для печати. Выделите нужный диапазон и выберите «Выделенный диапазон».');
  }
  const merges = (() => { try { return model.getMergedCells(sheet); } catch { return []; } })();
  const covered = new Set<string>();
  const spans = new Map<string, [number, number]>();
  for (const m of merges) {
    spans.set(`${m.row}:${m.column}`, [m.height, m.width]);
    for (let rr = m.row; rr < m.row + m.height; rr++) {
      for (let cc = m.column; cc < m.column + m.width; cc++) if (rr !== m.row || cc !== m.column) covered.add(`${rr}:${cc}`);
    }
  }
  const cols: number[] = [];
  let width = 0;
  for (let c = r.c1; c <= r.c2; c++) {
    const w = Math.round(model.getColumnWidth(sheet, c));
    cols.push(w);
    width += w;
  }
  let height = 0;
  const rowHtml = (row: number): string => {
    const h = Math.round(model.getRowHeight(sheet, row) * ROW_SCALE);
    if (h === 0) return ''; // скрытая строка не печатается
    height += h;
    let tds = '';
    for (let c = r.c1; c <= r.c2; c++) {
      if (cols[c - r.c1] === 0 || covered.has(`${row}:${c}`)) continue;
      const span = spans.get(`${row}:${c}`);
      const st = (() => { try { return model.getCellStyle(sheet, row, c).style; } catch { return null; } })();
      const text = model.getFormattedCellValue(sheet, row, c);
      const attrs = span ? ` rowspan="${span[0]}" colspan="${span[1]}"` : '';
      tds += `<td${attrs} style="${esc(cellCss(st, model, text ? model.getCellType(sheet, row, c) : 0))}">${esc(text)}</td>`;
    }
    return `<tr style="height:${h}px">${tds}</tr>`;
  };
  // Заголовок на каждой странице: закреплённые строки (если область начинается в них), иначе первая.
  const frozen = (() => { try { return model.getFrozenRowsCount(sheet); } catch { return 0; } })();
  const headEnd = opts.repeatHeader ? (frozen >= r.r1 ? Math.min(frozen, r.r2) : r.r1) : r.r1 - 1;
  let first = '';
  for (let row = r.r1; row <= headEnd; row++) first += rowHtml(row);
  let body = '';
  for (let row = headEnd + 1; row <= r.r2; row++) body += rowHtml(row);
  const colgroup = `<colgroup>${cols.map((w) => (w ? `<col style="width:${w}px">` : '')).join('')}</colgroup>`;
  const head = opts.repeatHeader ? `<thead>${first}</thead><tbody>${body}</tbody>` : `<tbody>${body}</tbody>`;
  const cls = `print-table${opts.gridlines ? ' grid' : ''}`;
  const center = opts.centerH ? 'margin:0 auto;' : '';
  return { html: `<table class="${cls}" style="${center}width:${width}px">${colgroup}${head}</table>`, width, height };
}

/** Размер области печати на странице (без полей), px. */
function pageArea(opts: PrintOptions): [number, number] {
  const [t, rt, b, l] = marginsOf(opts);
  const [w, h] = paperOf(opts);
  return [Math.max(20, w - l - rt) * 96 / 25.4, Math.max(20, h - t - b) * 96 / 25.4];
}

/** Масштаб, чтобы вписать в страницу. */
function zoomFor(opts: PrintOptions, width: number, height: number): number {
  const [pxW, pxH] = pageArea(opts);
  if (opts.fit === 'width') return Math.min(1, pxW / width);
  if (opts.fit === 'page') return Math.min(1, pxW / width, pxH / Math.max(1, height));
  return 1;
}

/** Сколько страниц получится (оценка: по ширине × по высоте). */
function pageCount(opts: PrintOptions, width: number, height: number, z: number): number {
  const [pxW, pxH] = pageArea(opts);
  return Math.max(1, Math.ceil((width * z) / pxW - 0.01)) * Math.max(1, Math.ceil((height * z) / pxH - 0.01));
}

const PAPER_LABEL: Record<Paper, string> = { A4: 'A4', A3: 'A3', A5: 'A5', Letter: 'Letter' };

/** Предпросмотр: лист бумаги с таблицей в масштабе. */
export function renderPreview(host: HTMLElement, model: Model, sheet: number, opts: PrintOptions, sel: Range): void {
  const r = printRange(model, sheet, opts, sel);
  if (!r) { host.innerHTML = '<p class="bs-muted">На листе нет данных для печати.</p>'; return; }
  let t;
  try { t = buildTable(model, sheet, r, opts); } catch (e) { host.innerHTML = `<p class="bs-muted">${esc(String((e as Error).message ?? e))}</p>`; return; }
  const [mt, mr, mb, ml] = marginsOf(opts);
  const [paperW, paperH] = paperOf(opts);
  const z = zoomFor(opts, t.width, t.height);
  const pages = pageCount(opts, t.width, t.height, z);
  host.innerHTML = `<div class="print-paper" style="aspect-ratio:${paperW}/${paperH}">
    <div class="print-sheet" style="padding:${mt}mm ${mr}mm ${mb}mm ${ml}mm">
      <div class="print-zoom" style="zoom:${z}">${t.html}</div></div></div>
    <p class="bs-muted">${opts.fit === 'none' ? 'Без масштабирования' : `Масштаб ${Math.round(z * 100)}%`} · ${PAPER_LABEL[opts.paper]} ·
      ${opts.landscape ? 'альбомная' : 'книжная'} · страниц: ${pages} · строк: ${r.r2 - r.r1 + 1}, столбцов: ${r.c2 - r.c1 + 1}</p>`;
  // Бумага в предпросмотре — уменьшенная копия: подгоняем масштаб под ширину окна.
  const paper = host.querySelector<HTMLElement>('.print-paper')!;
  const inner = host.querySelector<HTMLElement>('.print-sheet')!;
  const pxPaperW = paperW * 96 / 25.4;
  inner.style.transform = `scale(${paper.clientWidth / pxPaperW})`;
  inner.style.width = `${pxPaperW}px`;
  inner.style.height = `${paperH * 96 / 25.4}px`;
}

/** Печать: системное окно (принтер или PDF). */
export function printSheet(model: Model, sheet: number, opts: PrintOptions, sel: Range, title: string): void {
  const r = printRange(model, sheet, opts, sel);
  if (!r) throw new Error('На листе нет данных для печати.');
  const t = buildTable(model, sheet, r, opts);
  const [mt, mr, mb, ml] = marginsOf(opts);
  let root = document.getElementById('print-root');
  if (!root) {
    root = document.createElement('div');
    root.id = 'print-root';
    document.body.append(root);
  }
  // Поля страницы — нулевые: тогда браузер не печатает в них свои дату, имя и
  // путь к файлу. Наши поля — пустые строки шапки и подвала внешней таблицы
  // (повторяются на каждой странице) и отступы слева/справа.
  const size = opts.paper === 'Letter' ? 'letter' : opts.paper;
  root.innerHTML = `<style>@page { size: ${size} ${opts.landscape ? 'landscape' : 'portrait'}; margin: 0; }</style>
    <table class="print-page" style="padding: 0 ${mr}mm 0 ${ml}mm; width: 100%">
      <thead><tr><td style="height:${mt}mm"></td></tr></thead>
      <tfoot><tr><td style="height:${mb}mm"></td></tr></tfoot>
      <tbody><tr><td><div class="print-zoom" style="zoom:${zoomFor(opts, t.width, t.height)}">${t.html}</div></td></tr></tbody>
    </table>`;
  const old = document.title;
  document.title = title; // имя PDF по умолчанию
  document.body.classList.add('printing');
  const done = () => {
    document.body.classList.remove('printing');
    document.title = old;
    root!.innerHTML = '';
    window.removeEventListener('afterprint', done);
  };
  window.addEventListener('afterprint', done);
  window.print();
}
