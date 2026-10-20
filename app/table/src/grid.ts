// Сетка ENS-Table на Canvas: рисуются только видимые ячейки.
// Требования пользователя (docs/03-table-ux-requirements.md):
//  §1 активная ячейка видна всегда, в неактивном окне — бледнее, но заметно;
//  §2 «перекрестие» (фокус ячейки) до активной ячейки, настраиваемое.

import { columnName, type BorderItem, type CellBorder, type CellStyle, type MergedCell, type Model } from './engine';
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
/** Высота строки в движке (IronCalc: 25 = стандартная) → пиксели на экране. */
const ROW_SCALE = ROW_H / 25;
/** Насколько близко к границе заголовка курсор превращается в «⟷». */
const GRIP = 4;
const LAST_ROW = 1_048_576;
const LAST_COLUMN = 16_384;

export class Grid {
  sheet = 0;
  sel: Selection = { row: 1, column: 1, r1: 1, c1: 1, r2: 1, c2: 1 };
  focusCell: FocusCellOptions = { enabled: false, mode: 'shade', extent: 'toCell', color: '#2b7a6b' };
  theme: GridTheme = {
    activeBorder: '#2b7a6b',
    activeBorderInactive: '#6fae9f',
    activeWidth: 3,
    selectionFill: 'rgba(43,122,107,0.12)',
    headerActiveFill: '#d3ebe5',
    headerActiveText: '#1d5a4f',
  };
  /** Окно приложения в фокусе (влияет только на яркость выделения). */
  windowActive = true;
  onSelect: (sel: Selection) => void = () => {};
  /** Диапазон, выбираемый для формулы (пунктирная рамка, как в Excel). */
  pointRange: { r1: number; c1: number; r2: number; c2: number } | null = null;
  /** Ссылки редактируемой формулы на этом листе: у каждой свой цвет, как в Excel. */
  refHighlights: { r1: number; c1: number; r2: number; c2: number; color: string }[] = [];
  /** Щелчок или протяжка по ячейке во время ввода формулы. Вернуть true,
   *  если щелчок обработан (тогда выделение листа не меняется). */
  onPointer: (row: number, column: number, phase: 'down' | 'drag', shift: boolean) => boolean = () => false;
  private pointing = false;
  /** Двойной щелчок или F2: начать редактирование. */
  onEditRequest: (initial?: string) => void = () => {};
  /** Перетащили границу заголовка: новая ширина столбца / высота строки, px
   *  (null — двойной щелчок, подобрать по содержимому). */
  onResize: (kind: 'column' | 'row', index: number, px: number | null) => void = () => {};
  /** Правый щелчок по ячейке или заголовку. */
  onContextMenu: (x: number, y: number, header: 'row' | 'column' | 'corner' | null) => void = () => {};
  /** Протягивание маркера заполнения: куда дотянули. */
  onFill: (target: { r1: number; c1: number; r2: number; c2: number }) => void = () => {};
  /** Скопированный или вырезанный диапазон — пунктирная рамка, как в Excel. */
  clipRange: { r1: number; c1: number; r2: number; c2: number; sheet: number } | null = null;
  private filling: { r1: number; c1: number; r2: number; c2: number } | null = null;
  private handle = { x: -100, y: -100 };
  /** Закреплённые строки и столбцы (из книги). */
  private fr = 0;
  private fc = 0;
  private resizing: { kind: 'column' | 'row'; index: number; start: number; startSize: number; size: number } | null = null;

  readonly scroller: HTMLDivElement;
  private readonly spacer: HTMLDivElement;
  private readonly canvas: HTMLCanvasElement;
  private readonly ctx: CanvasRenderingContext2D;
  private headerW = 50;
  private colX: number[] = [0];
  private rowY: number[] = [0];
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
    window.addEventListener('mouseup', () => { this.endResize(); this.endFill(); });
    this.canvas.addEventListener('dblclick', (e) => {
      const g = this.grip(e);
      if (g) this.onResize(g.kind, g.index, null);
      else if (!this.hit(e).header) this.onEditRequest();
    });
    this.canvas.addEventListener('contextmenu', (e) => {
      e.preventDefault();
      const h = this.hit(e);
      const s = this.sel;
      // Правый щелчок вне выделения выделяет ячейку (строку, столбец), как в Excel.
      if (h.header === 'column' && !(s.c1 <= h.column && h.column <= s.c2 && s.r1 === 1 && s.r2 === LAST_ROW)) this.selectRange(1, h.column, LAST_ROW, h.column);
      else if (h.header === 'row' && !(s.r1 <= h.row && h.row <= s.r2 && s.c1 === 1 && s.c2 === LAST_COLUMN)) this.selectRange(h.row, 1, h.row, LAST_COLUMN);
      else if (!h.header && !(s.r1 <= h.row && h.row <= s.r2 && s.c1 <= h.column && h.column <= s.c2)) this.select(h.row, h.column);
      this.onContextMenu(e.clientX, e.clientY, h.header);
    });
    window.addEventListener('focus', () => { this.windowActive = true; this.render(); });
    window.addEventListener('blur', () => { this.windowActive = false; this.render(); });
  }

  // --- Геометрия ---

  private columnWidth(column: number): number {
    if (this.resizing?.kind === 'column' && this.resizing.index === column) return this.resizing.size;
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

  private rowHeight(row: number): number {
    if (this.resizing?.kind === 'row' && this.resizing.index === row) return this.resizing.size;
    try {
      return Math.max(0, Math.round(this.model().getRowHeight(this.sheet, row) * ROW_SCALE));
    } catch {
      return ROW_H;
    }
  }

  /** Верхняя граница строки в пикселях листа (без заголовка столбцов). */
  private y(row: number): number {
    while (this.rowY.length <= row) {
      const r = this.rowY.length;
      this.rowY.push(this.rowY[r - 1] + this.rowHeight(r));
    }
    return this.rowY[row - 1];
  }

  private height(row: number): number {
    return this.y(row + 1) - this.y(row);
  }

  private rowAt(py: number): number {
    let lo = 1, hi = Math.max(2, this.extentRows + 1);
    while (this.y(hi) <= py && hi < LAST_ROW) hi = Math.min(LAST_ROW, hi * 2);
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.y(mid) <= py) lo = mid; else hi = mid - 1;
    }
    return lo;
  }

  /** Высота строки на экране (px) → значение для движка. */
  static rowPxToModel(px: number): number { return px / ROW_SCALE; }

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
    this.rowY = [0];
    this.readFrozen();
    this.updateExtent();
    this.render();
  }

  private updateExtent(): void {
    this.extentRows = Math.min(LAST_ROW, Math.max(1000, this.sel.r2 + 200, this.extentRows));
    this.extentCols = Math.min(LAST_COLUMN, Math.max(52, this.sel.c2 + 10, this.extentCols));
    this.spacer.style.height = `${HEADER_H + this.y(this.extentRows + 1)}px`;
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

  private readFrozen(): void {
    try {
      this.fr = this.model().getFrozenRowsCount(this.sheet);
      this.fc = this.model().getFrozenColumnsCount(this.sheet);
    } catch {
      this.fr = this.fc = 0;
    }
  }

  /** Высота / ширина закреплённой области, px. */
  private get frozenH(): number { return this.y(this.fr + 1); }
  private get frozenW(): number { return this.x(this.fc + 1); }

  /** Первая строка прокручиваемой области. */
  private get firstRow(): number {
    return Math.max(this.fr + 1, this.rowAt(this.scroller.scrollTop + this.frozenH));
  }

  private get firstColumn(): number {
    return Math.max(this.fc + 1, this.columnAt(this.scroller.scrollLeft + this.frozenW));
  }

  private visibleRows(): number {
    const top = this.firstRow;
    let n = 0;
    while (top + n < LAST_ROW && this.y(top + n + 1) - this.scroller.scrollTop < this.scroller.clientHeight - HEADER_H) n++;
    return n;
  }

  /** Экранные координаты: закреплённые строки и столбцы не прокручиваются. */
  private sy(row: number): number {
    return HEADER_H + this.y(row) - (row <= this.fr ? 0 : this.scroller.scrollTop);
  }

  private sx(column: number): number {
    return this.headerW + this.x(column) - (column <= this.fc ? 0 : this.scroller.scrollLeft);
  }

  /** Координаты листа (px) под точкой экрана в области ячеек. */
  private sheetY(py: number): number {
    const b = py - HEADER_H;
    return b < this.frozenH ? b : b + this.scroller.scrollTop;
  }

  private sheetX(px: number): number {
    const b = px - this.headerW;
    return b < this.frozenW ? b : b + this.scroller.scrollLeft;
  }

  /** Прямоугольник ячейки на экране (в координатах canvas). */
  cellRect(row: number, column: number): { x: number; y: number; w: number; h: number } {
    return { x: this.sx(column), y: this.sy(row), w: this.width(column), h: this.height(row) };
  }

  /** Закрепить области: строки выше и столбцы левее ячейки (0 — снять). */
  setFrozen(rows: number, columns: number): void {
    this.model().setFrozenRowsCount(this.sheet, rows);
    this.model().setFrozenColumnsCount(this.sheet, columns);
    this.scroller.scrollTop = 0;
    this.scroller.scrollLeft = 0;
    this.invalidateLayout();
  }

  get frozen(): { rows: number; columns: number } {
    this.readFrozen();
    return { rows: this.fr, columns: this.fc };
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
      // Щелчок в объединённую ячейку выделяет её целиком.
      const m = this.mergeOf(row, column);
      this.sel = m
        ? { row: m.row, column: m.column, r1: m.row, c1: m.column, r2: m.row + m.height - 1, c2: m.column + m.width - 1 }
        : { row, column, r1: row, c1: column, r2: row, c2: column };
      if (m) { row = m.row; column = m.column; }
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
    let row = from.row + dRow, column = from.column + dCol;
    // Из объединённой ячейки — сразу за её край, как в Excel.
    const m = extend ? undefined : this.mergeOf(from.row, from.column);
    if (m) {
      if (dRow > 0) row = m.row + m.height; else if (dRow < 0) row = m.row - 1; else row = m.row;
      if (dCol > 0) column = m.column + m.width; else if (dCol < 0) column = m.column - 1; else column = m.column;
    }
    this.select(row, column, extend);
  }

  /** Объединённая ячейка, в которую входит (row, column). */
  mergeOf(row: number, column: number): MergedCell | undefined {
    let merges: MergedCell[] = [];
    try { merges = this.model().getMergedCells(this.sheet) ?? []; } catch { /* нет листа */ }
    return merges.find((m) => row >= m.row && row < m.row + m.height && column >= m.column && column < m.column + m.width);
  }

  /** Прямоугольник активной ячейки (объединённой — целиком). */
  activeRect(): { x: number; y: number; w: number; h: number } {
    const m = this.mergeOf(this.sel.row, this.sel.column);
    if (!m) return this.cellRect(this.sel.row, this.sel.column);
    const a = this.cellRect(m.row, m.column), b = this.cellRect(m.row + m.height - 1, m.column + m.width - 1);
    return { x: a.x, y: a.y, w: b.x + b.w - a.x, h: b.y + b.h - a.y };
  }

  pageRows(): number {
    return Math.max(1, this.visibleRows() - 1);
  }

  ensureVisible(row: number, column: number): void {
    if (row > this.fr) {
      const top = this.y(row) - this.frozenH, rh = this.height(row);
      const viewH = this.scroller.clientHeight - HEADER_H - this.frozenH;
      if (top < this.scroller.scrollTop) this.scroller.scrollTop = top;
      else if (top + rh > this.scroller.scrollTop + viewH) this.scroller.scrollTop = top + rh - viewH;
    }
    if (column > this.fc) {
      const left = this.x(column) - this.frozenW, cw = this.width(column);
      const viewW = this.scroller.clientWidth - this.headerW - this.frozenW;
      if (left < this.scroller.scrollLeft) this.scroller.scrollLeft = left;
      else if (left + cw > this.scroller.scrollLeft + viewW) this.scroller.scrollLeft = left + cw - viewW;
    }
  }

  private hit(e: MouseEvent): { row: number; column: number; header: 'row' | 'column' | 'corner' | null } {
    const r = this.canvas.getBoundingClientRect();
    const px = e.clientX - r.left, py = e.clientY - r.top;
    const row = py < HEADER_H ? 0 : this.rowAt(this.sheetY(py));
    const column = px < this.headerW ? 0 : this.columnAt(this.sheetX(px));
    const header = row === 0 && column === 0 ? 'corner' : row === 0 ? 'column' : column === 0 ? 'row' : null;
    return { row, column, header };
  }

  /** Курсор у границы заголовка: какой столбец/строку тянуть. */
  private grip(e: MouseEvent): { kind: 'column' | 'row'; index: number } | null {
    const r = this.canvas.getBoundingClientRect();
    const px = e.clientX - r.left, py = e.clientY - r.top;
    if (py < HEADER_H && px > this.headerW) {
      const sx = this.sheetX(px);
      const c = this.columnAt(sx);
      if (this.x(c + 1) - sx <= GRIP) return { kind: 'column', index: c };
      if (sx - this.x(c) <= GRIP && c > 1) return { kind: 'column', index: c - 1 };
    } else if (px < this.headerW && py > HEADER_H) {
      const sy = this.sheetY(py);
      const row = this.rowAt(sy);
      if (this.y(row + 1) - sy <= GRIP) return { kind: 'row', index: row };
      if (sy - this.y(row) <= GRIP && row > 1) return { kind: 'row', index: row - 1 };
    }
    return null;
  }

  private endFill(): void {
    const f = this.filling;
    if (!f) return;
    this.filling = null;
    document.body.style.cursor = '';
    const s = this.sel;
    if (f.r1 !== s.r1 || f.r2 !== s.r2 || f.c1 !== s.c1 || f.c2 !== s.c2) this.onFill(f);
    this.render();
  }

  private endResize(): void {
    const z = this.resizing;
    if (!z) return;
    this.resizing = null;
    document.body.style.cursor = '';
    if (z.size !== z.startSize) this.onResize(z.kind, z.index, z.size);
    this.invalidateLayout();
  }

  private mouseDown(e: MouseEvent): void {
    if (e.button !== 0) return;
    const g = this.grip(e);
    if (g) {
      e.preventDefault();
      const size = g.kind === 'column' ? this.width(g.index) : this.height(g.index);
      this.resizing = { ...g, start: g.kind === 'column' ? e.clientX : e.clientY, startSize: size, size };
      document.body.style.cursor = g.kind === 'column' ? 'col-resize' : 'row-resize';
      return;
    }
    const cr = this.canvas.getBoundingClientRect();
    if (Math.abs(e.clientX - cr.left - this.handle.x) <= 5 && Math.abs(e.clientY - cr.top - this.handle.y) <= 5) {
      // Маркер заполнения: протянуть формулу или ряд (1, 2 → 3, 4…).
      e.preventDefault();
      const s = this.sel;
      this.filling = { r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 };
      document.body.style.cursor = 'crosshair';
      return;
    }
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
    const z = this.resizing;
    if (z) {
      const pos = z.kind === 'column' ? e.clientX : e.clientY;
      z.size = Math.max(z.kind === 'column' ? 8 : 6, Math.round(z.startSize + pos - z.start));
      this.colX = [0];
      this.rowY = [0];
      this.render();
      return;
    }
    if (this.filling) {
      const h = this.hit(e), s = this.sel;
      const row = Math.max(1, h.row || s.r2), column = Math.max(1, h.column || s.c2);
      // Тянут дальше по вертикали или по горизонтали — что больше.
      const dr = row > s.r2 ? row - s.r2 : row < s.r1 ? row - s.r1 : 0;
      const dc = column > s.c2 ? column - s.c2 : column < s.c1 ? column - s.c1 : 0;
      this.filling = Math.abs(dr) >= Math.abs(dc)
        ? { r1: Math.min(s.r1, row), r2: Math.max(s.r2, row), c1: s.c1, c2: s.c2 }
        : { r1: s.r1, r2: s.r2, c1: Math.min(s.c1, column), c2: Math.max(s.c2, column) };
      this.ensureVisible(row, column);
      this.render();
      return;
    }
    if (e.target === this.canvas && !this.dragging && !this.pointing) {
      const g = this.grip(e);
      const cr = this.canvas.getBoundingClientRect();
      const onHandle = Math.abs(e.clientX - cr.left - this.handle.x) <= 5 && Math.abs(e.clientY - cr.top - this.handle.y) <= 5;
      this.canvas.style.cursor = g ? (g.kind === 'column' ? 'col-resize' : 'row-resize') : onHandle ? 'crosshair' : '';
    }
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
    this.readFrozen();
    const r0 = this.firstRow;
    let rN = r0;
    while (rN < LAST_ROW && HEADER_H + this.y(rN + 1) - st < H) rN++;
    const c0 = this.firstColumn;
    let cN = c0;
    while (cN < LAST_COLUMN && this.headerW + this.x(cN + 1) - sl < W) cN++;
    // Полосы: закреплённые строки/столбцы и прокручиваемые.
    const rowBands: [number, number][] = this.fr ? [[1, this.fr], [r0, rN]] : [[r0, rN]];
    const colBands: [number, number][] = this.fc ? [[1, this.fc], [c0, cN]] : [[c0, cN]];
    const shown = (r: number, c: number) => (r <= this.fr || (r >= r0 && r <= rN)) && (c <= this.fc || (c >= c0 && c <= cN));
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

    // Стили и объединения читаем один раз за кадр.
    const styles = new Map<number, CellStyle | null>();
    const styleOf = (r: number, c: number): CellStyle | null => {
      const key = r * 16_385 + c;
      let st = styles.get(key);
      if (st === undefined) {
        try { st = model.getCellStyle(this.sheet, r, c).style; } catch { st = null; }
        styles.set(key, st);
      }
      return st;
    };
    let merges: MergedCell[] = [];
    try { merges = model.getMergedCells(this.sheet) ?? []; } catch { merges = []; }
    const mergeAt = (r: number, c: number) =>
      merges.find((m) => r >= m.row && r < m.row + m.height && c >= m.column && c < m.column + m.width);
    const textOf = (r: number, c: number) => { try { return model.getFormattedCellValue(this.sheet, r, c); } catch { return ''; } };

    for (const [ra, rb] of rowBands) for (const [ca, cb] of colBands) {
    ctx.save();
    ctx.beginPath();
    // Прокручиваемая часть не заходит под закреплённую.
    const left = ca <= this.fc ? this.sx(ca) : Math.max(this.sx(ca), this.headerW + this.frozenW);
    const top = ra <= this.fr ? this.sy(ra) : Math.max(this.sy(ra), HEADER_H + this.frozenH);
    ctx.rect(left, top, this.sx(cb) + this.width(cb) - left, this.sy(rb) + this.height(rb) - top);
    ctx.clip();

    // Линии сетки — под заливкой: залитые ячейки их скрывают, как в Excel.
    if (gridLines) {
      ctx.strokeStyle = '#e1e1e1';
      ctx.lineWidth = 1;
      ctx.beginPath();
      for (let r = ra; r <= rb + 1; r++) {
        const y = Math.round(this.sy(r)) + 0.5;
        ctx.moveTo(this.headerW, y); ctx.lineTo(W, y);
      }
      for (let c = ca; c <= cb + 1; c++) {
        const x = Math.round(this.sx(c)) + 0.5;
        ctx.moveTo(x, HEADER_H); ctx.lineTo(x, H);
      }
      ctx.stroke();
    }

    // Ячейки: обычные и объединённые (одна большая клетка с данными левой верхней).
    const cells: { r: number; c: number; x: number; y: number; w: number; h: number; merged: boolean }[] = [];
    for (let r = ra; r <= rb; r++) {
      for (let c = ca; c <= cb; c++) {
        if (mergeAt(r, c)) continue;
        cells.push({ r, c, ...this.cellRect(r, c), merged: false });
      }
    }
    for (const m of merges) {
      const r2 = m.row + m.height - 1, c2 = m.column + m.width - 1;
      if (r2 < ra || m.row > rb || c2 < ca || m.column > cb) continue;
      const a = this.cellRect(m.row, m.column), b = this.cellRect(r2, c2);
      cells.push({ r: m.row, c: m.column, x: a.x, y: a.y, w: b.x + b.w - a.x, h: b.y + b.h - a.y, merged: true });
    }

    // Заливка (объединённая ячейка без заливки — белая: внутри нет линий сетки).
    for (const cell of cells) {
      const color = styleOf(cell.r, cell.c)?.fill?.color;
      const fill = color ? model.resolveColor(color) : cell.merged ? '#ffffff' : '';
      if (fill) { ctx.fillStyle = fill; ctx.fillRect(cell.x, cell.y, cell.w, cell.h); }
    }

    // Выделенный диапазон (кроме активной ячейки).
    if ((s.r1 !== s.r2 || s.c1 !== s.c2) && s.r2 >= ra && s.r1 <= rb && s.c2 >= ca && s.c1 <= cb) {
      const a = this.cellRect(Math.max(s.r1, ra), Math.max(s.c1, ca));
      const b = this.cellRect(Math.min(s.r2, rb), Math.min(s.c2, cb));
      ctx.fillStyle = this.theme.selectionFill;
      ctx.fillRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y);
    }

    // Текст.
    for (const cell of cells) {
      const text = textOf(cell.r, cell.c);
      if (!text) continue;
      const style = styleOf(cell.r, cell.c);
      // Длинный текст без переноса заходит на пустые соседние ячейки.
      let spillL = cell.x, spillR = cell.x + cell.w;
      const al = style?.alignment;
      let align = al?.horizontal ?? 'general';
      if (align === 'general') {
        const type = model.getCellType(this.sheet, cell.r, cell.c);
        align = type === 1 ? 'right' : type === 4 || type === 16 ? 'center' : 'left';
      }
      if (!al?.wrap_text && !cell.merged && align !== 'fill') {
        if (align === 'left' || align === 'center') {
          for (let c = cell.c + 1; c <= cb && !textOf(cell.r, c) && !mergeAt(cell.r, c); c++) spillR = this.sx(c) + this.width(c);
        }
        if (align === 'right' || align === 'center') {
          for (let c = cell.c - 1; c >= ca && !textOf(cell.r, c) && !mergeAt(cell.r, c); c--) spillL = this.sx(c);
        }
      }
      this.drawText(text, style, align, cell, spillL, spillR);
    }

    // Границы — поверх всего.
    for (const cell of cells) {
      const b = styleOf(cell.r, cell.c)?.border;
      if (b) this.drawBorders(b, cell);
    }
    ctx.restore();
    }
    if (this.focusCell.enabled) this.drawFocusCell(r0, rN, c0, cN);

    // Активная ячейка: толстая рамка; в неактивном окне — бледнее.
    const border = this.windowActive ? this.theme.activeBorder : this.theme.activeBorderInactive;
    if (shown(s.row, s.column)) {
      const { x, y, w, h } = this.activeRect();
      const lw = this.theme.activeWidth;
      ctx.strokeStyle = border;
      ctx.lineWidth = lw;
      ctx.strokeRect(x + lw / 2 - 1, y + lw / 2 - 1, w - lw + 2, h - lw + 2);
    }
    // Рамка диапазона (не для одной объединённой ячейки).
    const am = this.mergeOf(s.row, s.column);
    const onlyMerge = am && s.r1 === am.row && s.c1 === am.column && s.r2 === am.row + am.height - 1 && s.c2 === am.column + am.width - 1;
    if ((s.r1 !== s.r2 || s.c1 !== s.c2) && !onlyMerge) {
      const a = this.cellRect(s.r1, s.c1), b = this.cellRect(Math.min(s.r2, LAST_ROW), Math.min(s.c2, LAST_COLUMN));
      ctx.strokeStyle = border;
      ctx.lineWidth = 1.5;
      ctx.strokeRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y);
    }
    // Маркер заполнения — в правом нижнем углу выделения, как в Excel.
    this.handle = { x: -100, y: -100 };
    if (shown(s.r2, s.c2) && s.r2 < LAST_ROW && s.c2 < LAST_COLUMN) {
      const b = this.cellRect(s.r2, s.c2);
      this.handle = { x: b.x + b.w, y: b.y + b.h };
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(this.handle.x - 4, this.handle.y - 4, 8, 8);
      ctx.fillStyle = border;
      ctx.fillRect(this.handle.x - 3, this.handle.y - 3, 6, 6);
    }
    // Протягивание и скопированный диапазон — пунктир.
    const dashed = (r: { r1: number; c1: number; r2: number; c2: number }, color: string) => {
      const a = this.cellRect(r.r1, r.c1), b = this.cellRect(r.r2, r.c2);
      ctx.save();
      ctx.setLineDash([4, 3]);
      ctx.lineWidth = 2;
      ctx.strokeStyle = color;
      ctx.strokeRect(a.x + 1, a.y + 1, b.x + b.w - a.x - 2, b.y + b.h - a.y - 2);
      ctx.restore();
    };
    if (this.filling) dashed(this.filling, '#555');
    if (this.clipRange && this.clipRange.sheet === this.sheet) dashed(this.clipRange, border);
    for (const h of this.refHighlights) {
      if ((h.r2 < r0 && h.r1 > this.fr) || h.r1 > rN || (h.c2 < c0 && h.c1 > this.fc) || h.c1 > cN) continue;
      const a = this.cellRect(h.r1, h.c1), b = this.cellRect(h.r2, h.c2);
      ctx.save();
      ctx.globalAlpha = 0.1;
      ctx.fillStyle = h.color;
      ctx.fillRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y);
      ctx.globalAlpha = 1;
      ctx.lineWidth = 2;
      ctx.strokeStyle = h.color;
      ctx.strokeRect(a.x + 1, a.y + 1, b.x + b.w - a.x - 2, b.y + b.h - a.y - 2);
      ctx.restore();
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
    // Граница закреплённой области.
    ctx.strokeStyle = '#9a9a9a';
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    if (this.fr) { const y = Math.round(HEADER_H + this.frozenH) + 0.5; ctx.moveTo(this.headerW, y); ctx.lineTo(W, y); }
    if (this.fc) { const x = Math.round(this.headerW + this.frozenW) + 0.5; ctx.moveTo(x, HEADER_H); ctx.lineTo(x, H); }
    ctx.stroke();
    ctx.restore();

    this.drawHeaders(rowBands, colBands, W, H);
  }

  /** Текст ячейки: шрифт, выравнивание по горизонтали и вертикали, перенос по словам. */
  private drawText(text: string, style: CellStyle | null, align: string,
    cell: { x: number; y: number; w: number; h: number }, spillL: number, spillR: number): void {
    const ctx = this.ctx, model = this.model();
    const f = style?.font;
    const px = Math.round((f?.sz ?? 11) * 1.33);
    ctx.font = `${f?.i ? 'italic ' : ''}${f?.b ? 'bold ' : ''}${px}px ${fontStack(f?.name)}`;
    ctx.fillStyle = f?.color ? model.resolveColor(f.color) : '#000';
    const wrap = !!style?.alignment?.wrap_text;
    const lines = wrap ? wrapLines(ctx, text, cell.w - 6) : [text];
    const lineH = Math.round(px * 1.22);
    const vertical = style?.alignment?.vertical ?? 'bottom';
    const blockH = lines.length * lineH;
    let y0 = vertical === 'top' ? cell.y + 2
      : vertical === 'center' || vertical === 'justify' || vertical === 'distributed' ? cell.y + (cell.h - blockH) / 2
      : cell.y + cell.h - blockH - 2;
    ctx.save();
    ctx.beginPath(); ctx.rect(spillL + 1, cell.y, spillR - spillL - 2, cell.h); ctx.clip();
    ctx.textBaseline = 'middle';
    ctx.textAlign = align === 'right' ? 'right' : align === 'center' || align === 'centercontinuous' ? 'center' : 'left';
    const tx = ctx.textAlign === 'right' ? cell.x + cell.w - 3 : ctx.textAlign === 'center' ? cell.x + cell.w / 2 : cell.x + 3;
    for (const line of lines) {
      const ty = y0 + lineH / 2 + 1;
      if (align === 'fill' && line) {
        // «С заполнением»: текст повторяется на всю ширину.
        let t = line;
        while (ctx.measureText(t + line).width <= cell.w - 6) t += line;
        ctx.fillText(t, tx, ty);
      } else {
        ctx.fillText(line, tx, ty);
      }
      if (f?.u || f?.strike) {
        const tw = ctx.measureText(line).width;
        const ux = ctx.textAlign === 'right' ? tx - tw : ctx.textAlign === 'center' ? tx - tw / 2 : tx;
        if (f?.u) ctx.fillRect(ux, ty + px / 2 - 1, tw, 1);
        if (f?.strike) ctx.fillRect(ux, ty, tw, 1);
      }
      y0 += lineH;
    }
    ctx.restore();
  }

  /** Границы ячейки из файла: тонкая, средняя, толстая, двойная, пунктир. */
  private drawBorders(b: CellBorder, cell: { x: number; y: number; w: number; h: number }): void {
    const ctx = this.ctx, model = this.model();
    const x1 = Math.round(cell.x), y1 = Math.round(cell.y);
    const x2 = Math.round(cell.x + cell.w), y2 = Math.round(cell.y + cell.h);
    const side = (item: BorderItem | undefined, ax: number, ay: number, bx: number, by: number, vertical: boolean) => {
      if (!item?.style) return;
      const st = item.style;
      const width = st === 'thick' ? 3 : st.startsWith('medium') ? 2 : 1;
      ctx.save();
      ctx.strokeStyle = item.color ? model.resolveColor(item.color) || '#000' : '#000';
      ctx.lineWidth = width;
      ctx.setLineDash(st === 'dotted' ? [1, 2] : st === 'mediumdashed' ? [5, 2] : st.includes('dashdotdot') ? [5, 2, 1, 2, 1, 2]
        : st.includes('dashdot') ? [5, 2, 1, 2] : []);
      const off = width % 2 ? 0.5 : 0;
      ctx.beginPath();
      if (st === 'double') {
        for (const d of [-1, 1]) {
          ctx.moveTo(vertical ? ax + d + off : ax, vertical ? ay : ay + d + off);
          ctx.lineTo(vertical ? bx + d + off : bx, vertical ? by : by + d + off);
        }
      } else {
        ctx.moveTo(vertical ? ax + off : ax, vertical ? ay : ay + off);
        ctx.lineTo(vertical ? bx + off : bx, vertical ? by : by + off);
      }
      ctx.stroke();
      ctx.restore();
    };
    side(b.top, x1, y1, x2, y1, false);
    side(b.bottom, x1, y2, x2, y2, false);
    side(b.left, x1, y1, x1, y2, true);
    side(b.right, x2, y1, x2, y2, true);
  }

  /** Сколько строк займёт текст с переносом в ячейке шириной width (для автовысоты). */
  measureWrapped(text: string, style: CellStyle | null, width: number): number {
    const f = style?.font;
    const px = Math.round((f?.sz ?? 11) * 1.33);
    this.ctx.font = `${f?.i ? 'italic ' : ''}${f?.b ? 'bold ' : ''}${px}px ${fontStack(f?.name)}`;
    return wrapLines(this.ctx, text, width - 6).length * Math.round(px * 1.22) + 6;
  }

  /** Ширина столбца на экране, px. */
  columnPx(column: number): number { return this.width(column); }

  /** §2: полосы от заголовков к активной ячейке. */
  private drawFocusCell(r0: number, rN: number, c0: number, cN: number): void {
    const ctx = this.ctx, s = this.sel, fc = this.focusCell;
    const colRect = this.cellRect(this.fr ? 1 : r0, s.column);
    const rowRect = this.cellRect(s.row, this.fc ? 1 : c0);
    const rowEnd = fc.extent === 'full' ? rN : s.row - 1;
    const colEnd = fc.extent === 'full' ? cN : s.column - 1;
    const bandColH = rowEnd >= 1 ? this.sy(rowEnd) + this.height(rowEnd) - colRect.y : 0;
    const bandRowW = colEnd >= 1 ? this.sx(colEnd) + this.width(colEnd) - rowRect.x : 0;
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

  private drawHeaders(rowBands: [number, number][], colBands: [number, number][], W: number, H: number): void {
    const ctx = this.ctx, s = this.sel;
    ctx.fillStyle = '#f5f5f5';
    ctx.fillRect(0, 0, W, HEADER_H);
    ctx.fillRect(0, 0, this.headerW, H);
    ctx.font = '12px "Segoe UI", sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.strokeStyle = '#d4d4d4';
    ctx.lineWidth = 1;
    for (const [c0, cN] of colBands) {
    // Прокручиваемые заголовки не заходят под закреплённые.
    const left = c0 <= this.fc ? this.headerW : this.headerW + this.frozenW;
    ctx.save(); ctx.beginPath(); ctx.rect(left, 0, W - left, HEADER_H); ctx.clip();
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
    ctx.restore();
    }
    for (const [r0, rN] of rowBands) {
    const top = r0 <= this.fr ? HEADER_H : HEADER_H + this.frozenH;
    ctx.save(); ctx.beginPath(); ctx.rect(0, top, this.headerW, H - top); ctx.clip();
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
    ctx.restore();
    }
    const z = this.resizing;
    if (z) {
      // Размер при перетаскивании границы — подсказка, как в Excel.
      const label = z.kind === 'column' ? `Ширина: ${z.size} пикс.` : `Высота: ${z.size} пикс.`;
      const at = z.kind === 'column' ? this.cellRect(1, z.index) : this.cellRect(z.index, 1);
      ctx.font = '12px "Segoe UI", sans-serif';
      const tw = ctx.measureText(label).width + 12;
      const bx = z.kind === 'column' ? Math.min(W - tw - 2, at.x + at.w + 6) : this.headerW + 6;
      const by = z.kind === 'column' ? HEADER_H + 4 : at.y + at.h + 4;
      ctx.fillStyle = '#fffbe6'; ctx.strokeStyle = '#999';
      ctx.fillRect(bx, by, tw, 20); ctx.strokeRect(bx + 0.5, by + 0.5, tw, 20);
      ctx.fillStyle = '#222'; ctx.textAlign = 'left';
      ctx.fillText(label, bx + 6, by + 10);
      ctx.strokeStyle = 'rgba(0,0,0,0.45)'; ctx.setLineDash([3, 3]); ctx.beginPath();
      if (z.kind === 'column') { const lx = Math.round(at.x + at.w) + 0.5; ctx.moveTo(lx, 0); ctx.lineTo(lx, H); }
      else { const ly = Math.round(at.y + at.h) + 0.5; ctx.moveTo(0, ly); ctx.lineTo(W, ly); }
      ctx.stroke(); ctx.setLineDash([]);
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

/** Перенос по словам (и по символам, если слово длиннее ячейки); \n — новая строка. */
function wrapLines(ctx: CanvasRenderingContext2D, text: string, width: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    let line = '';
    for (const word of para.split(/(\s+)/)) {
      const next = line + word;
      if (!line || ctx.measureText(next).width <= width) { line = next; continue; }
      out.push(line.trimEnd());
      line = word.trimStart();
      while (line && ctx.measureText(line).width > width && line.length > 1) {
        let n = line.length - 1;
        while (n > 1 && ctx.measureText(line.slice(0, n)).width > width) n--;
        out.push(line.slice(0, n));
        line = line.slice(n);
      }
    }
    out.push(line);
  }
  return out;
}
