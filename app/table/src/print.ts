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
  /** Колонтитулы в записи Excel: &L/&C/&R — части, &P — страница, &N — страниц,
   *  &F — файл, &A — лист, &D — дата, &T — время. */
  header: string;
  footer: string;
}

export const DEFAULT_PRINT: PrintOptions = {
  what: 'sheet', landscape: false, paper: 'A4', fit: 'width', margins: 'normal', marginMm: [19, 18, 19, 18],
  gridlines: false, repeatHeader: false, centerH: false, header: '', footer: '',
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
  for (const k of ['header', 'footer'] as const) if (typeof v[k] === 'string') o[k] = v[k];
  if (Array.isArray(v.marginMm) && v.marginMm.length === 4 && v.marginMm.every((x) => typeof x === 'number')) {
    o.marginMm = v.marginMm;
    const preset = (Object.keys(MARGINS_MM) as (keyof typeof MARGINS_MM)[])
      .find((k) => MARGINS_MM[k].every((x, i) => Math.abs(x - o.marginMm[i]) < 1.5));
    o.margins = preset ?? 'custom';
  }
  return o;
}

export function toSaved(o: PrintOptions): string {
  const { landscape, paper, fit, gridlines, repeatHeader, centerH, header, footer } = o;
  return JSON.stringify({ landscape, paper, fit, marginMm: marginsOf(o), gridlines, repeatHeader, centerH, header, footer });
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

/** Что печатаем: выделение, область печати листа или занятая часть листа. */
export function printRange(model: Model, sheet: number, opts: PrintOptions, sel: Range): Range | null {
  if (opts.what === 'selection' && (sel.r1 !== sel.r2 || sel.c1 !== sel.c2)) {
    return { r1: sel.r1, c1: sel.c1, r2: Math.min(sel.r2, 1_048_576), c2: Math.min(sel.c2, 16_384) };
  }
  const area = printAreaOf(model, sheet);
  if (area) return area;
  const [r1, c1, r2, c2] = model.getSheetDimension(sheet);
  if (r1 === r2 && c1 === c2 && model.getFormattedCellValue(sheet, r1, c1) === '') return null;
  return { r1: 1, c1: 1, r2, c2 };
}

/** Сквозные строки и столбцы листа: [строка1, строка2, столбец1, столбец2], 0 — не заданы. */
export function printTitles(model: Model, sheet: number): number[] {
  try { const t = model.getPrintTitles(sheet); return t.length === 4 ? t : [0, 0, 0, 0]; } catch { return [0, 0, 0, 0]; }
}

/** Область печати листа (Excel: _xlnm.Print_Area) или null. */
export function printAreaOf(model: Model, sheet: number): Range | null {
  try {
    const a = model.getPrintArea(sheet);
    return a.length === 4 ? { r1: a[0], c1: a[1], r2: a[2], c2: a[3] } : null;
  } catch { return null; }
}

interface PrintRow { html: string; h: number; /** нельзя начинать с неё страницу (внутри объединения) */ keep: boolean }
interface Built { colgroup: string; cls: string; width: number; head: PrintRow[]; body: PrintRow[] }

/** Строки таблицы для печати и предпросмотра. */
function buildRows(model: Model, sheet: number, r: Range, opts: PrintOptions, colList: number[]): Built {
  if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) > MAX_CELLS) {
    throw new Error('Слишком большая область для печати. Выделите нужный диапазон и выберите «Выделенный диапазон».');
  }
  const merges = (() => { try { return model.getMergedCells(sheet); } catch { return []; } })();
  const covered = new Set<string>();
  const insideMerge = new Set<number>();
  const spans = new Map<string, [number, number]>();
  for (const m of merges) {
    spans.set(`${m.row}:${m.column}`, [m.height, m.width]);
    for (let rr = m.row; rr < m.row + m.height; rr++) {
      if (rr > m.row) insideMerge.add(rr);
      for (let cc = m.column; cc < m.column + m.width; cc++) if (rr !== m.row || cc !== m.column) covered.add(`${rr}:${cc}`);
    }
  }
  const cols: number[] = [];
  let width = 0;
  for (const c of colList) {
    const w = Math.round(model.getColumnWidth(sheet, c));
    cols.push(w);
    width += w;
  }
  const rowOf = (row: number): PrintRow | null => {
    const h = Math.round(model.getRowHeight(sheet, row) * ROW_SCALE);
    if (h === 0) return null; // скрытая строка не печатается
    let tds = '';
    for (const [i, c] of colList.entries()) {
      if (cols[i] === 0 || covered.has(`${row}:${c}`)) continue;
      const span = spans.get(`${row}:${c}`);
      const st = (() => { try { return model.getCellStyle(sheet, row, c).style; } catch { return null; } })();
      const text = model.getFormattedCellValue(sheet, row, c);
      const attrs = span ? ` rowspan="${span[0]}" colspan="${span[1]}"` : '';
      tds += `<td${attrs} style="${esc(cellCss(st, model, text ? model.getCellType(sheet, row, c) : 0))}">${esc(text)}</td>`;
    }
    return { html: `<tr style="height:${h}px">${tds}</tr>`, h, keep: insideMerge.has(row) };
  };
  // Заголовок на каждой странице: закреплённые строки (если область начинается в них), иначе первая.
  const frozen = (() => { try { return model.getFrozenRowsCount(sheet); } catch { return 0; } })();
  const titles = printTitles(model, sheet);
  const head: PrintRow[] = [];
  const body: PrintRow[] = [];
  if (titles[0]) {
    // Сквозные строки листа (Excel: Print_Titles) — на каждой странице, даже если они вне области печати.
    for (let row = titles[0]; row <= titles[1]; row++) { const pr = rowOf(row); if (pr) head.push(pr); }
    for (let row = r.r1; row <= r.r2; row++) {
      if (row >= titles[0] && row <= titles[1]) continue;
      const pr = rowOf(row);
      if (pr) body.push(pr);
    }
  } else {
    const headEnd = opts.repeatHeader ? (frozen >= r.r1 ? Math.min(frozen, r.r2) : r.r1) : r.r1 - 1;
    for (let row = r.r1; row <= r.r2; row++) {
      const pr = rowOf(row);
      if (pr) (row <= headEnd ? head : body).push(pr);
    }
  }
  const colgroup = `<colgroup>${cols.map((w) => (w ? `<col style="width:${w}px">` : '')).join('')}</colgroup>`;
  return { colgroup, cls: `print-table${opts.gridlines ? ' grid' : ''}`, width, head, body };
}

function tableHtml(b: Built, opts: PrintOptions, rows: PrintRow[]): string {
  const center = opts.centerH ? 'margin:0 auto;' : '';
  const head = b.head.length ? `<thead>${b.head.map((x) => x.html).join('')}</thead>` : '';
  return `<table class="${b.cls}" style="${center}width:${b.width}px">${b.colgroup}${head}<tbody>${rows.map((x) => x.html).join('')}</tbody></table>`;
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

/** Разбивка строк на страницы (как будет напечатано): заголовок — на каждой. */
function paginate(b: Built, opts: PrintOptions): { pages: PrintRow[][]; zoom: number } {
  const headH = b.head.reduce((s, x) => s + x.h, 0);
  const total = headH + b.body.reduce((s, x) => s + x.h, 0);
  const zoom = zoomFor(opts, b.width, total);
  // Небольшой запас: в печати строки с переносом бывают чуть выше.
  const avail = Math.max(20, (pageArea(opts)[1] / zoom) * 0.98 - headH);
  const pages: PrintRow[][] = [];
  let cur: PrintRow[] = [];
  let used = 0;
  for (const row of b.body) {
    if (cur.length && used + row.h > avail && !row.keep) {
      pages.push(cur);
      cur = [];
      used = 0;
    }
    cur.push(row);
    used += row.h;
  }
  if (cur.length || !pages.length) pages.push(cur);
  return { pages, zoom };
}

interface PrintPage { b: Built; rows: PrintRow[]; zoom: number }

/** Все страницы: сначала вниз, затем вправо (как в Excel). Если таблица шире страницы и масштаб
 *  не задан — столбцы делятся на части, сквозные столбцы повторяются в каждой. */
function layout(model: Model, sheet: number, r: Range, opts: PrintOptions): { pages: PrintPage[]; chunks: number } {
  const titles = printTitles(model, sheet);
  const titleCols: number[] = [];
  if (titles[2]) for (let c = titles[2]; c <= titles[3]; c++) titleCols.push(c);
  const own: number[] = [];
  for (let c = r.c1; c <= r.c2; c++) if (!titleCols.includes(c)) own.push(c);
  const widthOf = (c: number) => Math.round(model.getColumnWidth(sheet, c));
  const chunks: number[][] = [];
  if (opts.fit !== 'none') {
    chunks.push(own);
  } else {
    const avail = pageArea(opts)[0] - titleCols.reduce((s, c) => s + widthOf(c), 0);
    let cur: number[] = [], used = 0;
    for (const c of own) {
      const w = widthOf(c);
      if (cur.length && used + w > avail) { chunks.push(cur); cur = []; used = 0; }
      cur.push(c);
      used += w;
    }
    if (cur.length || !chunks.length) chunks.push(cur);
  }
  const pages: PrintPage[] = [];
  for (const chunk of chunks) {
    const b = buildRows(model, sheet, r, opts, [...titleCols, ...chunk]);
    const p = paginate(b, opts);
    for (const rows of p.pages) pages.push({ b, rows, zoom: p.zoom });
  }
  return { pages, chunks: chunks.length };
}

// --- Колонтитулы (запись Excel: &L, &C, &R, &P, &N, &D, &T, &F, &A) ---

export const HEADER_PRESETS: [string, string][] = [
  ['', 'Нет'], ['&C&F', 'Имя файла'], ['&C&A', 'Имя листа'], ['&L&F&R&A', 'Имя файла и лист'], ['&R&D', 'Дата печати'],
];
export const FOOTER_PRESETS: [string, string][] = [
  ['', 'Нет'], ['&CСтраница &P', 'Страница 1'], ['&CСтраница &P из &N', 'Страница 1 из N'],
  ['&L&F&RСтраница &P из &N', 'Имя файла; страница 1 из N'], ['&L&A&RСтраница &P из &N', 'Имя листа; страница 1 из N'],
];

interface HfContext { page: number; pages: number; file: string; sheet: string }

/** Колонтитул Excel → три части (слева, по центру, справа). Шрифтовые коды пропускаются. */
export function hfParts(code: string, ctx: HfContext): [string, string, string] {
  const parts: [string, string, string] = ['', '', ''];
  let sec = 1;
  const now = new Date();
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch !== '&') { parts[sec] += ch; continue; }
    const n = code[++i] ?? '';
    switch (n.toUpperCase()) {
      case 'L': sec = 0; break;
      case 'C': sec = 1; break;
      case 'R': sec = 2; break;
      case 'P': parts[sec] += String(ctx.page); break;
      case 'N': parts[sec] += String(ctx.pages); break;
      case 'F': parts[sec] += ctx.file; break;
      case 'A': parts[sec] += ctx.sheet; break;
      case 'D': parts[sec] += now.toLocaleDateString('ru-RU'); break;
      case 'T': parts[sec] += now.toLocaleTimeString('ru-RU', { hour: '2-digit', minute: '2-digit' }); break;
      case '&': parts[sec] += '&'; break;
      case '"': { const end = code.indexOf('"', i + 1); i = end < 0 ? code.length : end; break; } // &"Шрифт,Жирный"
      case 'K': i += 6; break; // &Kцвет
      default:
        if (/\d/.test(n)) { while (/\d/.test(code[i + 1] ?? '')) i++; } // &12 — размер шрифта
        // &B, &I, &U, &S, &E, &X, &Y — начертание: пропускаем
    }
  }
  return parts;
}

function hfHtml(code: string, ctx: HfContext, pos: string): string {
  if (!code) return '';
  const [l, c, r] = hfParts(code, ctx);
  return `<div class="print-hf" style="${pos}"><span>${esc(l)}</span><span>${esc(c)}</span><span>${esc(r)}</span></div>`;
}

/** Одна страница: бумага, поля, колонтитулы, часть таблицы. */
function pageHtml(b: Built, opts: PrintOptions, rows: PrintRow[], zoom: number, ctx: HfContext, scaleForPreview = false): string {
  const [mt, mr, mb, ml] = marginsOf(opts);
  const [w, h] = paperOf(opts);
  const size = scaleForPreview ? '' : `width:${w}mm;height:${h - 0.5}mm;`;
  return `<div class="print-pg" style="${size}padding:${mt}mm ${mr}mm ${mb}mm ${ml}mm">
    ${hfHtml(opts.header, ctx, `left:${ml}mm;right:${mr}mm;top:${Math.max(3, mt / 2 - 2)}mm`)}${hfHtml(opts.footer, ctx, `left:${ml}mm;right:${mr}mm;bottom:${Math.max(3, mb / 2 - 2)}mm`)}
    <div class="print-zoom" style="zoom:${zoom}">${tableHtml(b, opts, rows)}</div></div>`;
}

const PAPER_LABEL: Record<Paper, string> = { A4: 'A4', A3: 'A3', A5: 'A5', Letter: 'Letter' };
let previewPage = 0;

/** Предпросмотр: страница с таблицей в масштабе, листание страниц. */
export function renderPreview(host: HTMLElement, model: Model, sheet: number, opts: PrintOptions, sel: Range, file = ''): void {
  const r = printRange(model, sheet, opts, sel);
  if (!r) { host.innerHTML = '<p class="bs-muted">На листе нет данных для печати.</p>'; return; }
  let lay: ReturnType<typeof layout>;
  try { lay = layout(model, sheet, r, opts); } catch (e) { host.innerHTML = `<p class="bs-muted">${esc(String((e as Error).message ?? e))}</p>`; return; }
  const { pages, chunks } = lay;
  const zoom = pages[0].zoom;
  previewPage = Math.max(0, Math.min(previewPage, pages.length - 1));
  const [paperW, paperH] = paperOf(opts);
  const sheetName = model.getWorksheetsProperties()[sheet]?.name ?? '';
  const ctx: HfContext = { page: previewPage + 1, pages: pages.length, file, sheet: sheetName };
  const area = printAreaOf(model, sheet) && opts.what === 'sheet' ? ' · область печати' : '';
  host.innerHTML = `<div class="print-paper" style="aspect-ratio:${paperW}/${paperH}">
      <div class="print-sheet">${pageHtml(pages[previewPage].b, opts, pages[previewPage].rows, pages[previewPage].zoom, ctx, true)}</div></div>
    <div class="print-nav">
      <button data-pg="-1" ${previewPage ? '' : 'disabled'} title="Предыдущая страница">◀</button>
      <span>Страница ${previewPage + 1} из ${pages.length}</span>
      <button data-pg="1" ${previewPage < pages.length - 1 ? '' : 'disabled'} title="Следующая страница">▶</button>
    </div>
    <p class="bs-muted">${opts.fit === 'none' ? 'Без масштабирования' : `Масштаб ${Math.round(zoom * 100)}%`} · ${PAPER_LABEL[opts.paper]} ·
      ${opts.landscape ? 'альбомная' : 'книжная'}${area} · строк: ${r.r2 - r.r1 + 1}, столбцов: ${r.c2 - r.c1 + 1}
      ${chunks > 1 ? `<br>Таблица шире страницы: она разбита по столбцам на части (${chunks}), страницы идут сначала вниз, потом вправо.` : ''}</p>`;
  host.querySelectorAll<HTMLButtonElement>('[data-pg]').forEach((btn) => btn.addEventListener('click', () => {
    previewPage += Number(btn.dataset.pg);
    renderPreview(host, model, sheet, opts, sel, file);
  }));
  // Бумага в предпросмотре — уменьшенная копия: подгоняем масштаб под ширину окна.
  const paper = host.querySelector<HTMLElement>('.print-paper')!;
  const inner = host.querySelector<HTMLElement>('.print-sheet')!;
  const pxPaperW = paperW * 96 / 25.4;
  inner.style.transform = `scale(${paper.clientWidth / pxPaperW})`;
  inner.style.width = `${pxPaperW}px`;
  inner.style.height = `${paperH * 96 / 25.4}px`;
}

export function resetPreviewPage(): void { previewPage = 0; }

/** Печать: системное окно (принтер или PDF). Страницы раскладываем сами — так на
 *  каждой будут колонтитулы с номером страницы, в любом браузерном движке. */
export function printSheet(model: Model, sheet: number, opts: PrintOptions, sel: Range, title: string): void {
  const r = printRange(model, sheet, opts, sel);
  if (!r) throw new Error('На листе нет данных для печати.');
  const { pages } = layout(model, sheet, r, opts);
  const sheetName = model.getWorksheetsProperties()[sheet]?.name ?? '';
  let root = document.getElementById('print-root');
  if (!root) {
    root = document.createElement('div');
    root.id = 'print-root';
    document.body.append(root);
  }
  // Поля страницы у браузера — нулевые: тогда он не печатает в них свои дату, имя и
  // путь к файлу. Наши поля и колонтитулы — внутри каждой страницы.
  const size = opts.paper === 'Letter' ? 'letter' : opts.paper;
  root.innerHTML = `<style>@page { size: ${size} ${opts.landscape ? 'landscape' : 'portrait'}; margin: 0; }</style>
    ${pages.map((pg, i) => pageHtml(pg.b, opts, pg.rows, pg.zoom, { page: i + 1, pages: pages.length, file: title, sheet: sheetName })).join('')}`;
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
