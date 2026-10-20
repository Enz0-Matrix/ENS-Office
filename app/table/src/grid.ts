// Сетка ENS-Table на Canvas: рисуются только видимые ячейки.
// Требования пользователя (docs/03-table-ux-requirements.md):
//  §1 активная ячейка видна всегда, в неактивном окне — бледнее, но заметно;
//  §2 «перекрестие» (фокус ячейки) до активной ячейки, настраиваемое.

import { columnName, type Model } from './engine';
import { fontStack } from './fonts';

export interface Selection {
  /** Активная ячейка. */
  row: number;
  column: number;
  /** Выделенный диапазон (включительно). */
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface FocusCellOptions {
  enabled: boolean;
  /** shade — затемнять ячейки, border — подсвечивать границы полосы. */
  mode: 'shade' | 'border';
  /** toCell — от заголовка до ячейки, full — вся строка и весь столбец. */
  extent: 'toCell' | 'full';
  color: string;
}

export interface GridTheme {
  activeBorder: string;
  /** Цвет рамки, когда окно неактивно. */
  activeBorderInactive: string;
  activeWidth: number;
  selectionFill: string;
  headerActiveFill: string;
  headerActiveText: string;
}

const ROW_H = 22;
const HEADER_H = 22;
const LAST_ROW = 1_048_576;
const LAST_COLUMN = 16_384;

export class Grid {
  sheet = 0;
  sel: Selection = { row: 1, column: 1, r1: 1, c1: 1, r2: 1, c2: 1 };
  focusCell: FocusCellOptions = { enabled: false, mode: 'shade', extent: 'toCell', color: '#107c41' };
  theme: GridTheme = {
    activeBorder: '#107c41',
    activeBorderInactive: '#5fa880',
    activeWidth: 3,
    selectionFill: 'rgba(16,124,65,0.12)',
    headerActiveFill: '#caead8',
    headerActiveText: '#0b5a2f',
  };
  /** Окно приложения в фокусе (влияет только на яркость выделения). */
  windowActive = true;
  onSelect: (sel: Selection) => void = () => {};
  /** Диапазон, выбираемый для формулы (пунктирная рамка, как в Excel). */
  pointRange: { r1: number; c1: number; r2: number; c2: number } | null = null;
  /** Щелчок или протяжка по ячейке во время ввода формулы. Вернуть true,
   *  если щелчок обработан (тогда выделение листа не меняется). */
  onPointer: (row: number, column: number, phase: 'down' | 'drag', shift: boolean) => boolean = () => false;
  private pointing = false;
  /** Двойной щелчок или F2: начать редактирование. */
  onEditRequest: (initial?: string) => void = () => {};

  readonly scroller: HTMLDivElement;
  private readonly spacer: HTMLDivElement;
  private readonly canvas: HTMLCanvasElement;
  private readonly ctx: CanvasRenderingContext2D;
  private headerW = 50;
  private colX: number[] = [0];
  private extentRows = 1000;
  private extentCols = 52;
  private dragging = false;
  private frame = 0;

  constructor(host: HTMLElement, private model: () => Model) {
    this.scroller = document.createElement('div');
    this.scroller.className = 'grid-scroller';
    this.scroller.tabIndex = 0;
    this.spacer = document.createElement('div');
    this.spacer.className = 'grid-spacer';
    this.canvas = document.createElement('canvas');
    this.canvas.className = 'grid-canvas';
    this.scroller.append(this.spacer, this.canvas);
    host.append(this.scroller);
    this.ctx = this.canvas.getContext('2d')!;

    this.scroller.addEventListener('scroll', () => this.render());
    new ResizeObserver(() => this.resize()).observe(this.scroller);
    this.canvas.addEventListener('mousedown', (e) => this.mouseDown(e));
    window.addEventListener('mousemove', (e) => this.mouseMove(e));
    window.addEventListener('mouseup', () => { this.dragging = false; this.pointing = false; });
    this.canvas.addEventListener('dblclick', () => this.onEditRequest());
    window.addEventListener('focus', () => { this.windowActive = true; this.render(); });
    window.addEventListener('blur', () => { this.windowActive = false; this.render(); });
  }

  // --- Геометрия ---

  private columnWidth(column: number): number {
    try {
      return Math.max(4, Math.round(this.model().getColumnWidth(this.sheet, column)));
    } catch {
      return 100;
    }
  }

  /** Левая граница столбца в пикселях листа (без заголовка строк). */
  private x(column: number): number {
    while (this.colX.length <= column) {
      const c = this.colX.length;
      this.colX.push(this.colX[c - 1] + this.columnWidth(c));
    }
    return this.colX[column - 1];
  }

  private width(column: number): number {
    return this.x(column + 1) - this.x(column);
  }

  private columnAt(px: number): number {
    let lo = 1, hi = Math.max(2, this.extentCols + 1);
    while (this.x(hi) <= px && hi < LAST_COLUMN) hi = Math.min(LAST_COLUMN, hi * 2);
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.x(mid) <= px) lo = mid; else hi = mid - 1;
    }
    return lo;
  }

  /** Ширины столбцов могли измениться (новый файл, лист). */
  invalidateLayout(): void {
    this.colX = [0];
    this.updateExtent();
    this.render();
  }

  private updateExtent(): void {
    this.extentRows = Math.min(LAST_ROW, Math.max(1000, this.sel.r2 + 200, this.extentRows));
    this.extentCols = Math.min(LAST_COLUMN, Math.max(52, this.sel.c2 + 10, this.extentCols));
    this.spacer.style.height = `${HEADER_H + this.extentRows * ROW_H}px`;
    this.spacer.style.width = `${this.headerW + this.x(this.extentCols + 1)}px`;
  }

  private resize(): void {
    const dpr = window.devicePixelRatio || 1;
    const w = this.scroller.clientWidth, h = this.scroller.clientHeight;
    this.canvas.width = Math.round(w * dpr);
    this.canvas.height = Math.round(h * dpr);
    this.canvas.style.width = `${w}px`;
    this.canvas.style.height = `${h}px`;
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    this.updateExtent();
    this.render();
  }

  private get firstRow(): number {
    return Math.floor(this.scroller.scrollTop / ROW_H) + 1;
  }

  private get firstColumn(): number {
    return this.columnAt(this.scroller.scrollLeft);
  }

  private visibleRows(): number {
    return Math.ceil((this.scroller.clientHeight - HEADER_H) / ROW_H);
  }

  /** Прямоугольник ячейки на экране (в координатах canvas). */
  cellRect(row: number, column: number): { x: number; y: number; w: number; h: number } {
    return {
      x: this.headerW + this.x(column) - this.scroller.scrollLeft,
      y: HEADER_H + (row - 1) * ROW_H - this.scroller.scrollTop,
      w: this.width(column),
      h: ROW_H,
    };
  }

  // --- Выделение и навигация ---

  /** Подвижный угол выделения (активная ячейка — неподвижный якорь). */
  corner = { row: 1, column: 1 };

  select(row: number, column: number, extend = false): void {
    row = Math.min(LAST_ROW, Math.max(1, row));
    column = Math.min(LAST_COLUMN, Math.max(1, column));
    const s = this.sel;
    this.corner = { row, column };
    if (extend) {
      this.sel = {
        row: s.row, column: s.column,
        r1: Math.min(s.row, row), r2: Math.max(s.row, row),
        c1: Math.min(s.column, column), c2: Math.max(s.column, column),
      };
    } else {
      this.sel = { row, column, r1: row, c1: column, r2: row, c2: column };
    }
    this.ensureVisible(row, column);
    this.updateExtent();
    this.render();
    this.onSelect(this.sel);
  }

  selectRange(r1: number, c1: number, r2: number, c2: number): void {
    this.sel = { row: r1, column: c1, r1, c1, r2, c2 };
    this.corner = { row: r2, column: c2 };
    this.ensureVisible(r1, c1);
    this.updateExtent();
    this.render();
    this.onSelect(this.sel);
  }

  move(dRow: number, dCol: number, extend: boolean): void {
    const from = extend ? this.corner : { row: this.sel.row, column: this.sel.column };
    this.select(from.row + dRow, from.column + dCol, extend);
  }

  pageRows(): number {
    return Math.max(1, this.visibleRows() - 1);
  }

  ensureVisible(row: number, column: number): void {
    const top = (row - 1) * ROW_H;
    const viewH = this.scroller.clientHeight - HEADER_H;
    if (top < this.scroller.scrollTop) this.scroller.scrollTop = top;
    else if (top + ROW_H > this.scroller.scrollTop + viewH) this.scroller.scrollTop = top + ROW_H - viewH;
    const left = this.x(column);
    const viewW = this.scroller.clientWidth - this.headerW;
    if (left < this.scroller.scrollLeft) this.scroller.scrollLeft = left;
    else if (left + this.width(column) > this.scroller.scrollLeft + viewW) {
      this.scroller.scrollLeft = left + this.width(column) - viewW;
    }
  }

  private hit(e: MouseEvent): { row: number; column: number; header: 'row' | 'column' | 'corner' | null } {
    const r = this.canvas.getBoundingClientRect();
    const px = e.clientX - r.left, py = e.clientY - r.top;
    const row = py < HEADER_H ? 0 : Math.floor((py - HEADER_H + this.scroller.scrollTop) / ROW_H) + 1;
    const column = px < this.headerW ? 0 : this.columnAt(px - this.headerW + this.scroller.scrollLeft);
    const header = row === 0 && column === 0 ? 'corner' : row === 0 ? 'column' : column === 0 ? 'row' : null;
    return { row, column, header };
  }

  private mouseDown(e: MouseEvent): void {
    const h = this.hit(e);
    if (!h.header && h.row >= 1 && h.column >= 1 && this.onPointer(h.row, h.column, 'down', e.shiftKey)) {
      e.preventDefault(); // фокус остаётся в поле ввода формулы
      this.pointing = true;
      return;
    }
    this.scroller.focus();
    if (h.header === 'column') this.selectRange(1, h.column, LAST_ROW, h.column);
    else if (h.header === 'row') this.selectRange(h.row, 1, h.row, LAST_COLUMN);
    else if (h.header === 'corner') this.selectRange(1, 1, LAST_ROW, LAST_COLUMN);
    else {
      this.select(h.row, h.column, e.shiftKey);
      this.dragging = true;
    }
  }

  private mouseMove(e: MouseEvent): void {
    if (this.pointing) {
      const p = this.hit(e);
      if (p.row >= 1 && p.column >= 1) this.onPointer(p.row, p.column, 'drag', true);
      return;
    }
    if (!this.dragging) return;
    const h = this.hit(e);
    if (h.row >= 1 && h.column >= 1) this.select(h.row, h.column, true);
  }

  // --- Отрисовка ---

  render(): void {
    if (this.frame) return;
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.draw();
    });
  }

  private draw(): void {
    const ctx = this.ctx;
    const model = this.model();
    const W = this.scroller.clientWidth, H = this.scroller.clientHeight;
    const sl = this.scroller.scrollLeft, st = this.scroller.scrollTop;
    const r0 = this.firstRow, rN = Math.min(LAST_ROW, r0 + this.visibleRows());
    const c0 = this.firstColumn;
    let cN = c0;
    while (cN < LAST_COLUMN && this.headerW + this.x(cN + 1) - sl < W) cN++;
    const s = this.sel;
    const gridLines = (() => { try { return model.getShowGridLines(this.sheet); } catch { return true; } })();

    // Ширина заголовков строк — по количеству цифр.
    const hw = Math.max(40, 12 + 8 * String(rN).length);
    if (hw !== this.headerW) { this.headerW = hw; this.updateExtent(); }

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, W, H);
    ctx.save();
    ctx.beginPath();
    ctx.rect(this.headerW, HEADER_H, W - this.headerW, H - HEADER_H);
    ctx.clip();

    // Заливка ячеек и «перекрестие» под текстом.
    for (let r = r0; r <= rN; r++) {
      for (let c = c0; c <= cN; c++) {
        const { x, y, w, h } = this.cellRect(r, c);
        let fill: string | undefined;
        try {
          const color = model.getCellStyle(this.sheet, r, c).style.fill.color;
          if (color) fill = model.resolveColor(color);
        } catch { /* пустая ячейка */ }
        if (fill) { ctx.fillStyle = fill; ctx.fillRect(x, y, w, h); }
      }
    }
    if (this.focusCell.enabled) this.drawFocusCell(r0, rN, c0, cN);

    // Выделенный диапазон (кроме активной ячейки).
    if (s.r1 !== s.r2 || s.c1 !== s.c2) {
      const a = this.cellRect(Math.max(s.r1, r0), Math.max(s.c1, c0));
      const b = this.cellRect(Math.min(s.r2, rN), Math.min(s.c2, cN));
      ctx.fillStyle = this.theme.selectionFill;
      ctx.fillRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y);
    }

    // Линии сетки.
    if (gridLines) {
      ctx.strokeStyle = '#e1e1e1';
      ctx.lineWidth = 1;
      ctx.beginPath();
      for (let r = r0; r <= rN + 1; r++) {
        const y = Math.round(HEADER_H + (r - 1) * ROW_H - st) + 0.5;
        ctx.moveTo(this.headerW, y); ctx.lineTo(W, y);
      }
      for (let c = c0; c <= cN + 1; c++) {
        const x = Math.round(this.headerW + this.x(c) - sl) + 0.5;
        ctx.moveTo(x, HEADER_H); ctx.lineTo(x, H);
      }
      ctx.stroke();
    }

    // Текст.
    ctx.textBaseline = 'middle';
    for (let r = r0; r <= rN; r++) {
      for (let c = c0; c <= cN; c++) {
        let text: string;
        try { text = model.getFormattedCellValue(this.sheet, r, c); } catch { continue; }
        if (!text) continue;
        const { x, y, w, h } = this.cellRect(r, c);
        let style;
        try { style = model.getCellStyle(this.sheet, r, c).style; } catch { style = undefined; }
        const f = style?.font;
        ctx.font = `${f?.i ? 'italic ' : ''}${f?.b ? 'bold ' : ''}${Math.round((f?.sz ?? 11) * 1.33)}px ${fontStack(f?.name)}`;
        ctx.fillStyle = f?.color ? model.resolveColor(f.color) : '#000';
        let align = style?.alignment?.horizontal ?? 'general';
        if (align === 'general') {
          const type = model.getCellType(this.sheet, r, c);
          align = type === 1 ? 'right' : type === 4 || type === 16 ? 'center' : 'left';
        }
        ctx.save();
        ctx.beginPath(); ctx.rect(x + 1, y, w - 2, h); ctx.clip();
        const tx = align === 'right' ? x + w - 3 : align === 'center' ? x + w / 2 : x + 3;
        ctx.textAlign = align === 'right' ? 'right' : align === 'center' ? 'center' : 'left';
        ctx.fillText(text, tx, y + h / 2 + 1);
        if (f?.u) {
          const tw = ctx.measureText(text).width;
          const ux = ctx.textAlign === 'right' ? tx - tw : ctx.textAlign === 'center' ? tx - tw / 2 : tx;
          ctx.fillRect(ux, y + h - 5, tw, 1);
        }
        ctx.restore();
      }
    }

    // Активная ячейка: толстая рамка; в неактивном окне — бледнее.
    if (s.row >= r0 && s.row <= rN && s.column >= c0 && s.column <= cN) {
      const { x, y, w, h } = this.cellRect(s.row, s.column);
      const lw = this.theme.activeWidth;
      ctx.strokeStyle = this.windowActive ? this.theme.activeBorder : this.theme.activeBorderInactive;
      ctx.lineWidth = lw;
      ctx.strokeRect(x + lw / 2 - 1, y + lw / 2 - 1, w - lw + 2, h - lw + 2);
      // Маркер заполнения в правом нижнем углу, как в Excel.
      ctx.fillStyle = ctx.strokeStyle;
      ctx.fillRect(x + w - 4, y + h - 4, 6, 6);
    }
    // Рамка диапазона.
    if (s.r1 !== s.r2 || s.c1 !== s.c2) {
      const a = this.cellRect(s.r1, s.c1), b = this.cellRect(s.r2, s.c2);
      ctx.strokeStyle = this.windowActive ? this.theme.activeBorder : this.theme.activeBorderInactive;
      ctx.lineWidth = 1.5;
      ctx.strokeRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y);
    }
    if (this.pointRange) {
      const p = this.pointRange;
      const a = this.cellRect(p.r1, p.c1), b = this.cellRect(p.r2, p.c2);
      ctx.save();
      ctx.setLineDash([5, 3]);
      ctx.lineWidth = 2;
      ctx.strokeStyle = '#2f6fd1';
      ctx.strokeRect(a.x + 1, a.y + 1, b.x + b.w - a.x - 2, b.y + b.h - a.y - 2);
      ctx.fillStyle = 'rgba(47,111,209,0.08)';
      ctx.fillRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y);
      ctx.restore();
    }
    ctx.restore();

    this.drawHeaders(r0, rN, c0, cN, W, H);
  }

  /** §2: полосы от заголовков к активной ячейке. */
  private drawFocusCell(r0: number, rN: number, c0: number, cN: number): void {
    const ctx = this.ctx, s = this.sel, fc = this.focusCell;
    const colRect = this.cellRect(r0, s.column);
    const rowRect = this.cellRect(s.row, c0);
    const rowEnd = fc.extent === 'full' ? rN : s.row - 1;
    const colEnd = fc.extent === 'full' ? cN : s.column - 1;
    const bandColH = HEADER_H + rowEnd * ROW_H - this.scroller.scrollTop - colRect.y;
    const bandRowW = this.headerW + this.x(colEnd + 1) - this.scroller.scrollLeft - rowRect.x;
    ctx.save();
    if (fc.mode === 'shade') {
      ctx.globalAlpha = 0.16;
      ctx.fillStyle = fc.color;
      if (bandColH > 0) ctx.fillRect(colRect.x, colRect.y, colRect.w, bandColH);
      if (bandRowW > 0) ctx.fillRect(rowRect.x, rowRect.y, bandRowW, rowRect.h);
      if (fc.extent === 'full') {
        // Не закрашиваем саму ячейку дважды: она и так выделена рамкой.
      }
    } else {
      ctx.strokeStyle = fc.color;
      ctx.lineWidth = 1.5;
      ctx.globalAlpha = 0.8;
      if (bandColH > 0) ctx.strokeRect(colRect.x + 0.75, colRect.y, colRect.w - 1.5, bandColH);
      if (bandRowW > 0) ctx.strokeRect(rowRect.x, rowRect.y + 0.75, bandRowW, rowRect.h - 1.5);
    }
    ctx.restore();
  }

  private drawHeaders(r0: number, rN: number, c0: number, cN: number, W: number, H: number): void {
    const ctx = this.ctx, s = this.sel;
    ctx.fillStyle = '#f5f5f5';
    ctx.fillRect(0, 0, W, HEADER_H);
    ctx.fillRect(0, 0, this.headerW, H);
    ctx.font = '12px "Segoe UI", sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.strokeStyle = '#d4d4d4';
    ctx.lineWidth = 1;
    for (let c = c0; c <= cN; c++) {
      const { x, w } = this.cellRect(1, c);
      const inSel = c >= s.c1 && c <= s.c2;
      if (inSel) {
        ctx.fillStyle = this.theme.headerActiveFill;
        ctx.fillRect(x, 0, w, HEADER_H);
        ctx.fillStyle = this.windowActive ? this.theme.activeBorder : this.theme.activeBorderInactive;
        ctx.fillRect(x, HEADER_H - 2, w, 2);
      }
      ctx.fillStyle = inSel ? this.theme.headerActiveText : '#444';
      ctx.font = `${c === s.column ? 'bold ' : ''}12px "Segoe UI", sans-serif`;
      ctx.fillText(columnName(c), x + w / 2, HEADER_H / 2);
      ctx.beginPath(); ctx.moveTo(Math.round(x + w) + 0.5, 0); ctx.lineTo(Math.round(x + w) + 0.5, HEADER_H); ctx.stroke();
    }
    for (let r = r0; r <= rN; r++) {
      const { y, h } = this.cellRect(r, 1);
      const inSel = r >= s.r1 && r <= s.r2;
      if (inSel) {
        ctx.fillStyle = this.theme.headerActiveFill;
        ctx.fillRect(0, y, this.headerW, h);
        ctx.fillStyle = this.windowActive ? this.theme.activeBorder : this.theme.activeBorderInactive;
        ctx.fillRect(this.headerW - 2, y, 2, h);
      }
      ctx.fillStyle = inSel ? this.theme.headerActiveText : '#444';
      ctx.font = `${r === s.row ? 'bold ' : ''}12px "Segoe UI", sans-serif`;
      ctx.fillText(String(r), this.headerW / 2, y + h / 2);
      ctx.beginPath(); ctx.moveTo(0, Math.round(y + h) + 0.5); ctx.lineTo(this.headerW, Math.round(y + h) + 0.5); ctx.stroke();
    }
    ctx.fillStyle = '#e8e8e8';
    ctx.fillRect(0, 0, this.headerW, HEADER_H);
    ctx.strokeStyle = '#c8c8c8';
    ctx.beginPath();
    ctx.moveTo(0, HEADER_H + 0.5); ctx.lineTo(W, HEADER_H + 0.5);
    ctx.moveTo(this.headerW + 0.5, 0); ctx.lineTo(this.headerW + 0.5, H);
    ctx.stroke();
  }
}
