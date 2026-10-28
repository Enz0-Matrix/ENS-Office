// Сетка ENS-Table на Canvas: рисуются только видимые ячейки.
// Требования пользователя (docs/03-table-ux-requirements.md):
//  §1 активная ячейка видна всегда, в неактивном окне — бледнее, но заметно;
//  §2 «перекрестие» (фокус ячейки) до активной ячейки, настраиваемое.

import { columnName, type BorderItem, type CellBorder, type CellStyle, type Chart, type ExtendedStyle, type ImageInfo, type MergedCell, type Model, type RichRun } from './engine';
import { fontStack } from './fonts';
import { drawChart, parseRange, type ChartData } from './chart';
import { PALETTE } from './numfmt';

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
  /** Лист без заливки, линии сетки, заголовки строк и столбцов (темы оформления). */
  sheetBg: string;
  gridLine: string;
  headerBg: string;
  headerLine: string;
  headerText: string;
  cornerBg: string;
  cornerLine: string;
  /** Цвет текста и границ «авто» (в тёмных темах — светлый). */
  textColor?: string;
}

/** Классическая тема сетки (нефритовая). */
export const CLASSIC_GRID: GridTheme = {
  activeBorder: '#2b7a6b',
  activeBorderInactive: '#6fae9f',
  activeWidth: 3,
  selectionFill: 'rgba(43,122,107,0.12)',
  headerActiveFill: '#d3ebe5',
  headerActiveText: '#1d5a4f',
  sheetBg: '#ffffff',
  gridLine: '#e1e1e1',
  headerBg: '#f5f5f5',
  headerLine: '#d4d4d4',
  headerText: '#444',
  cornerBg: '#e8e8e8',
  cornerLine: '#c8c8c8',
};

const ROW_H = 22;
const HEADER_H = 22;
/** Высота строки в движке (IronCalc: 25 = стандартная) → пиксели на экране. */
const ROW_SCALE = ROW_H / 25;
/** Насколько близко к границе заголовка курсор превращается в «⟷». */
const GRIP = 4;
/** Ширина одного уровня групп в полосе структуры (слева и сверху). */
const OUTLINE_STEP = 14;

/** Полоса строк/столбцов при отрисовке: [первая, последняя, закреплена снизу/справа]. */
type Band = [number, number, boolean?];

/** Группа строк или столбцов: a..b, уровень; итоговая строка — b + 1 (как в Excel). */
/** `s` — итоговая строка/столбец с кнопкой «+/−»: после группы (b + 1) или перед ней (a − 1). */
export interface OutlineGroup { kind: 'row' | 'column'; a: number; b: number; level: number; collapsed: boolean; s: number }

/** Группы из уровней: каждая непрерывная полоса с уровнем ≥ L — группа уровня L. */
function groupsOf(kind: 'row' | 'column', levels: Map<number, { level: number; collapsed: boolean }>, after: boolean): OutlineGroup[] {
  const out: OutlineGroup[] = [];
  const idx = [...levels.keys()].filter((i) => levels.get(i)!.level > 0).sort((x, y) => x - y);
  const max = Math.max(0, ...idx.map((i) => levels.get(i)!.level));
  for (let L = 1; L <= max; L++) {
    let start = 0, prev = 0;
    const flush = () => {
      if (!start) return;
      const s = after || start <= 1 ? prev + 1 : start - 1;
      out.push({ kind, a: start, b: prev, level: L, collapsed: !!levels.get(s)?.collapsed, s });
      start = 0;
    };
    for (const i of idx) {
      if (levels.get(i)!.level < L) continue;
      if (start && i === prev + 1) { prev = i; continue; }
      flush();
      start = prev = i;
    }
    flush();
  }
  return out;
}
const LAST_ROW = 1_048_576;
const LAST_COLUMN = 16_384;

export const ZOOM_MIN = 0.25, ZOOM_MAX = 4;

/** Цвет «#rgb», «#rrggbb» или «rgb(…)» → [r, g, b] (0–255); null — не разобрать. */
function parseRgb(c: string): [number, number, number] | null {
  const hex = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(c.trim());
  if (hex) {
    const h = hex[1].length === 3 ? [...hex[1]].map((x) => x + x).join('') : hex[1];
    return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
  }
  const rgb = /^rgba?\(\s*(\d+)[ ,]+(\d+)[ ,]+(\d+)/i.exec(c.trim());
  return rgb ? [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])] : null;
}

/** Яркость цвета 0 (чёрный) … 1 (белый); неизвестный цвет считаем светлым. */
function luminance(c: string): number {
  const rgb = parseRgb(c);
  return rgb ? (0.299 * rgb[0] + 0.587 * rgb[1] + 0.114 * rgb[2]) / 255 : 1;
}
export const MARK_COLORS = { blue: '#2f6fd1', red: '#d93025', green: '#1e9e4a' };

export class Grid {
  sheet = 0;
  /** Масштаб таблицы (1 — 100%): меняется только лист, лента и панели остаются как есть. */
  zoom = 1;
  onZoom: (zoom: number) => void = () => {};
  /** Видимая область и прокрутка в координатах листа (без масштаба). */
  private get viewW(): number { return this.scroller.clientWidth / this.zoom; }
  private get viewH(): number { return this.scroller.clientHeight / this.zoom; }
  private get sTop(): number { return this.scroller.scrollTop / this.zoom; }
  private get sLeft(): number { return this.scroller.scrollLeft / this.zoom; }

  setZoom(z: number): void {
    const next = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(z * 100) / 100));
    if (next === this.zoom) return;
    // Верхний левый угол видимой области остаётся на месте.
    const top = this.sTop, left = this.sLeft;
    this.zoom = next;
    this.resize();
    this.scroller.scrollTop = top * next;
    this.scroller.scrollLeft = left * next;
    this.onZoom(next);
    this.render();
  }
  sel: Selection = { row: 1, column: 1, r1: 1, c1: 1, r2: 1, c2: 1 };
  /** Ещё выделенные диапазоны (Ctrl+щелчок, как в Excel); sel — последний, с активной ячейкой. */
  extraRanges: { r1: number; c1: number; r2: number; c2: number }[] = [];
  /** Идёт Ctrl+щелчок: новое выделение добавляется к прежним, а не заменяет их. */
  private ctrlPick = false;
  focusCell: FocusCellOptions = { enabled: false, mode: 'shade', extent: 'toCell', color: '#2b7a6b' };
  theme: GridTheme = { ...CLASSIC_GRID };
  /** Окно приложения в фокусе (влияет только на яркость выделения). */
  windowActive = true;
  onSelect: (sel: Selection) => void = () => {};
  /** Диапазон, выбираемый для формулы (пунктирная рамка, как в Excel). */
  pointRange: { r1: number; c1: number; r2: number; c2: number } | null = null;
  /** Ссылки редактируемой формулы на этом листе: у каждой свой цвет, как в Excel. */
  refHighlights: { r1: number; c1: number; r2: number; c2: number; color: string; id?: number }[] = [];
  /** Цветную рамку ссылки тянут мышью: id — номер ссылки в формуле; rect — новый прямоугольник. */
  onRefDrag: (id: number, rect: { r1: number; c1: number; r2: number; c2: number }) => void = () => {};
  private refDrag: { id: number; mode: 'move' | 'resize'; row: number; column: number; orig: { r1: number; c1: number; r2: number; c2: number } } | null = null;

  /** Рамка ссылки под курсором: за угол — растянуть, за край — переместить. */
  private refFrameAt(e: MouseEvent): { id: number; mode: 'move' | 'resize'; orig: { r1: number; c1: number; r2: number; c2: number } } | null {
    const cr = this.canvas.getBoundingClientRect();
    const px = (e.clientX - cr.left) / this.zoom, py = (e.clientY - cr.top) / this.zoom;
    for (const h of this.refHighlights) {
      if (h.id === undefined) continue;
      const a = this.lrect(h.r1, h.c1), b = this.lrect(h.r2, h.c2);
      const x1 = a.x, y1 = a.y, x2 = b.x + b.w, y2 = b.y + b.h, t = 4;
      const inside = px >= x1 - t && px <= x2 + t && py >= y1 - t && py <= y2 + t;
      if (!inside) continue;
      const orig = { r1: h.r1, c1: h.c1, r2: h.r2, c2: h.c2 };
      if (Math.abs(px - x2) <= 6 && Math.abs(py - y2) <= 6) return { id: h.id, mode: 'resize', orig };
      const onEdge = Math.abs(px - x1) <= t || Math.abs(px - x2) <= t || Math.abs(py - y1) <= t || Math.abs(py - y2) <= t;
      if (onEdge) return { id: h.id, mode: 'move', orig };
    }
    return null;
  }
  /** Щелчок или протяжка по ячейке во время ввода формулы. Вернуть true,
   *  если щелчок обработан (тогда выделение листа не меняется). */
  onPointer: (row: number, column: number, phase: 'down' | 'drag', shift: boolean) => boolean = () => false;
  private pointing = false;
  /** Перед выбором ячейки мышью: завершить ввод в прежней ячейке. false — остаться в ней. */
  beforeSelect: () => boolean = () => true;
  /** Двойной щелчок или F2: начать редактирование. */
  onEditRequest: (initial?: string) => void = () => {};
  /** Перетащили границу заголовка: новая ширина столбца / высота строки, px
   *  (null — двойной щелчок, подобрать по содержимому). */
  onResize: (kind: 'column' | 'row', index: number, px: number | null) => void = () => {};
  /** Правый щелчок по ячейке или заголовку. */
  /** Рисунки листа, выбранный рисунок (-1 — нет) и перенос или изменение размера мышью. */
  images: ImageInfo[] = [];
  charts: Chart[] = [];
  selObj = -1;
  private imgCache = new Map<string, HTMLImageElement | null>();
  private imgDrag: { mode: 'move' | 'resize'; i: number; sx: number; sy: number; orig: { x: number; y: number; w: number; h: number }; cur: { x: number; y: number; w: number; h: number } } | null = null;
  /** Рисунок перенесён или изменён: индекс, ячейка левого верхнего угла, сдвиг и размер (px). */
  onImagePlace: (kind: 'image' | 'chart', index: number, row: number, column: number, dx: number, dy: number, width: number, height: number) => void = () => {};
  onImageSelect: (index: number) => void = () => {};
  /** Двойной щелчок по выбранному рисунку или диаграмме. */
  onObjectDouble: () => void = () => {};
  /** Ячейки с примечаниями ("строка:столбец") — красный уголок. */
  commentCells = new Set<string>();
  /** Ручные разрывы страниц листа: [после строк, после столбцов] — штриховые линии. */
  pageBreaks: [number[], number[]] = [[], []];
  /** Указатель над ячейкой (row 0 — не над ячейкой); для всплывающего примечания. */
  onCellHover: (row: number, column: number, clientX: number, clientY: number) => void = () => {};
  onContextMenu: (x: number, y: number, header: 'row' | 'column' | 'corner' | null) => void = () => {};
  /** Протягивание маркера заполнения: куда дотянули. */
  onFill: (target: { r1: number; c1: number; r2: number; c2: number }) => void = () => {};
  /** Скопированный или вырезанный диапазон — пунктирная рамка, как в Excel. */
  clipRange: { r1: number; c1: number; r2: number; c2: number; sheet: number } | null = null;
  /** Автофильтр листа: строка заголовков r1, столбцы c1..c2; active — где выбран отбор. */
  /** Фильтры листа (таблиц может быть несколько): кнопки в строке заголовков каждой. */
  filters: { r1: number; c1: number; r2: number; c2: number; active: Set<number> }[] = [];
  /** Щелчок по кнопке фильтра в заголовке столбца. */
  onFilterButton: (row: number, column: number, x: number, y: number) => void = () => {};
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
  /** Полоса структуры (группы): слева от номеров строк и над буквами столбцов. */
  private outlineW = 0;
  private outlineH = 0;
  private get headerH(): number { return HEADER_H + this.outlineH; }
  rowGroups: OutlineGroup[] = [];
  colGroups: OutlineGroup[] = [];
  private rowLevels = 0;
  private colLevels = 0;
  /** «+/−» у группы: свернуть (true) или развернуть. */
  onOutlineToggle: (g: OutlineGroup, collapse: boolean) => void = () => {};
  /** Кнопки уровней «1 2 3» в углу. */
  onOutlineLevel: (kind: 'row' | 'column', level: number) => void = () => {};

  /** Группы листа из движка: [индекс, уровень, свёрнуто, …]. */
  setOutline(rows: ArrayLike<number>, cols: ArrayLike<number>, below = false, right = false): void {
    const parse = (a: ArrayLike<number>) => {
      const m = new Map<number, { level: number; collapsed: boolean }>();
      for (let i = 0; i + 2 < a.length; i += 3) m.set(a[i], { level: a[i + 1], collapsed: a[i + 2] === 1 });
      return m;
    };
    this.rowGroups = groupsOf('row', parse(rows), below);
    this.colGroups = groupsOf('column', parse(cols), right);
    this.rowLevels = Math.max(0, ...this.rowGroups.map((g) => g.level));
    this.colLevels = Math.max(0, ...this.colGroups.map((g) => g.level));
    const w = this.rowLevels ? (this.rowLevels + 1) * OUTLINE_STEP + 4 : 0;
    const h = this.colLevels ? (this.colLevels + 1) * OUTLINE_STEP + 4 : 0;
    if (w !== this.outlineW || h !== this.outlineH) {
      this.outlineW = w;
      this.outlineH = h;
      this.headerW = 0; // пересчитается при отрисовке
      this.updateExtent();
    }
    this.render();
  }

  /** Щелчок в полосе структуры: «+/−» группы или кнопка уровня. true — обработан. */
  private outlineClick(px: number, py: number): boolean {
    if (px < this.outlineW && py < this.headerH) {
      // Кнопки уровней строк — в углу над полосой.
      const level = Math.floor((px - 2) / OUTLINE_STEP) + 1;
      if (level >= 1 && level <= this.rowLevels + 1) { this.onOutlineLevel('row', level); return true; }
      return false;
    }
    if (py < this.outlineH && px < this.headerW) {
      const level = Math.floor((py - 2) / OUTLINE_STEP) + 1;
      if (level >= 1 && level <= this.colLevels + 1) { this.onOutlineLevel('column', level); return true; }
      return false;
    }
    if (px < this.outlineW && py >= this.headerH) {
      const level = Math.floor((px - 2) / OUTLINE_STEP) + 1;
      const row = this.rowAt(this.sheetY(py));
      const g = this.rowGroups.find((x) => x.level === level && (x.s === row || (!x.collapsed && row >= x.a && row <= x.b)));
      if (g) { this.onOutlineToggle(g, !g.collapsed); return true; }
      return true; // щелчок по полосе — не выделение
    }
    if (py < this.outlineH && px >= this.headerW) {
      const level = Math.floor((py - 2) / OUTLINE_STEP) + 1;
      const column = this.columnAt(this.sheetX(px));
      const g = this.colGroups.find((x) => x.level === level && (x.s === column || (!x.collapsed && column >= x.a && column <= x.b)));
      if (g) { this.onOutlineToggle(g, !g.collapsed); return true; }
      return true;
    }
    return false;
  }

  /** Полоса структуры: линии групп, «+/−», кнопки уровней. */
  private drawOutline(W: number, H: number): void {
    if (!this.outlineW && !this.outlineH) return;
    const ctx = this.ctx;
    ctx.save();
    const box = (x: number, y: number, label: string) => {
      ctx.fillStyle = this.theme.sheetBg;
      ctx.strokeStyle = this.theme.headerText;
      ctx.lineWidth = 1;
      ctx.fillRect(x, y, 11, 11);
      ctx.strokeRect(x + 0.5, y + 0.5, 10, 10);
      ctx.fillStyle = this.theme.headerText;
      ctx.fillRect(x + 3, y + 5, 5, 1);
      if (label === '+') ctx.fillRect(x + 5, y + 3, 1, 5);
    };
    const levelButton = (x: number, y: number, n: number) => {
      ctx.fillStyle = this.theme.headerBg;
      ctx.strokeStyle = this.theme.headerLine;
      ctx.fillRect(x, y, 12, 12);
      ctx.strokeRect(x + 0.5, y + 0.5, 11, 11);
      ctx.fillStyle = this.theme.headerText;
      ctx.font = '9px "Segoe UI", sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(String(n), x + 6, y + 6.5);
    };
    ctx.fillStyle = this.theme.headerBg;
    if (this.outlineW) ctx.fillRect(0, this.headerH, this.outlineW, H);
    if (this.outlineH) ctx.fillRect(this.headerW, 0, W, this.outlineH);
    ctx.strokeStyle = this.theme.headerText;
    // Строки.
    if (this.outlineW) {
      ctx.save(); ctx.beginPath(); ctx.rect(0, this.headerH, this.outlineW, H); ctx.clip();
      for (const g of this.rowGroups) {
        const x = 2 + (g.level - 1) * OUTLINE_STEP;
        const top = this.lrect(g.a, 1), bot = this.lrect(g.b, 1), sum = this.lrect(g.s, 1);
        if (!g.collapsed) {
          ctx.beginPath();
          if (g.s > g.b) { ctx.moveTo(x + 5.5, top.y + 2); ctx.lineTo(x + 5.5, bot.y + bot.h - 1); ctx.lineTo(x + 2, bot.y + bot.h - 1); }
          else { ctx.moveTo(x + 5.5, bot.y + bot.h - 2); ctx.lineTo(x + 5.5, top.y + 1); ctx.lineTo(x + 2, top.y + 1); }
          ctx.stroke();
        }
        if (sum.h > 0) box(x, sum.y + sum.h / 2 - 5, g.collapsed ? '+' : '-');
      }
      ctx.restore();
      for (let n = 1; n <= this.rowLevels + 1; n++) levelButton(2 + (n - 1) * OUTLINE_STEP, this.headerH - 15, n);
    }
    // Столбцы.
    if (this.outlineH) {
      ctx.save(); ctx.beginPath(); ctx.rect(this.headerW, 0, W, this.outlineH); ctx.clip();
      for (const g of this.colGroups) {
        const y = 2 + (g.level - 1) * OUTLINE_STEP;
        const left = this.lrect(1, g.a), right = this.lrect(1, g.b), sum = this.lrect(1, g.s);
        if (!g.collapsed) {
          ctx.beginPath();
          if (g.s > g.b) { ctx.moveTo(left.x + 2, y + 5.5); ctx.lineTo(right.x + right.w - 1, y + 5.5); ctx.lineTo(right.x + right.w - 1, y + 2); }
          else { ctx.moveTo(right.x + right.w - 2, y + 5.5); ctx.lineTo(left.x + 1, y + 5.5); ctx.lineTo(left.x + 1, y + 2); }
          ctx.stroke();
        }
        if (sum.w > 0) box(sum.x + sum.w / 2 - 5, y, g.collapsed ? '+' : '-');
      }
      ctx.restore();
      for (let n = 1; n <= this.colLevels + 1; n++) levelButton(Math.max(2, this.headerW - 15), 2 + (n - 1) * OUTLINE_STEP, n);
    }
    ctx.restore();
  }
  private colX: number[] = [0];
  private extentRows = 1000;
  private extentCols = 52;
  private dragging = false;
  /** Протяжка по заголовкам: какие линии и от какой. */
  private headerDrag: { kind: 'row' | 'column'; anchor: number } | null = null;
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
    // Ctrl + колесо: масштаб листа, а не всей страницы (лента и панели остаются как есть).
    this.scroller.addEventListener('wheel', (e) => {
      if (!e.ctrlKey && !e.metaKey) return;
      e.preventDefault();
      this.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }, { passive: false });
    new ResizeObserver(() => this.resize()).observe(this.scroller);
    this.canvas.addEventListener('mousedown', (e) => this.mouseDown(e));
    window.addEventListener('mousemove', (e) => this.mouseMove(e));
    window.addEventListener('mouseup', () => { this.dragging = false; this.pointing = false; this.headerDrag = null; this.refDrag = null; });
    window.addEventListener('mouseup', () => { this.endResize(); this.endFill(); this.endImageDrag(); });
    this.canvas.addEventListener('dblclick', (e) => {
      const g = this.grip(e);
      const h = this.hit(e);
      if (this.selObj >= 0) this.onObjectDouble();
      else if (g) this.onResize(g.kind, g.index, null);
      else if (!h.header) this.onEditRequest();
      else if (h.header === 'row' && h.row >= 1) this.toggleMark('row', h.row, this.markColor(e));
      else if (h.header === 'column' && h.column >= 1) this.toggleMark('column', h.column, this.markColor(e));
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

  // --- Метки строк и столбцов (закладки, как в Notepad++) ---

  private marks = new Map<number, { rows: Map<number, string>; cols: Map<number, string> }>();
  /** Метки действуют, пока открыта книга: двойной щелчок по заголовку ставит (синяя; Shift — красная,
   *  Ctrl — зелёная) и снимает (любую) метку. */
  private marksOf(sheet = this.sheet): { rows: Map<number, string>; cols: Map<number, string> } {
    let m = this.marks.get(sheet);
    if (!m) this.marks.set(sheet, m = { rows: new Map(), cols: new Map() });
    return m;
  }

  /** Цвет новой метки: Shift — красная, Ctrl — зелёная, иначе синяя. */
  private markColor(e: MouseEvent): string {
    return e.shiftKey ? MARK_COLORS.red : e.ctrlKey || e.metaKey ? MARK_COLORS.green : MARK_COLORS.blue;
  }

  toggleMark(kind: 'row' | 'column', index: number, color = MARK_COLORS.blue): void {
    const map = kind === 'row' ? this.marksOf().rows : this.marksOf().cols;
    if (!map.delete(index)) map.set(index, color);
    this.render();
  }

  hasMark(kind: 'row' | 'column', index: number): boolean {
    return (kind === 'row' ? this.marksOf().rows : this.marksOf().cols).has(index);
  }

  /** Следующая (dir = 1) или предыдущая метка по кругу: строки, а если их нет — столбцы. */
  nextMark(dir: 1 | -1): { kind: 'row' | 'column'; index: number } | null {
    const m = this.marksOf();
    const kind = m.rows.size ? 'row' : 'column';
    const list = [...(kind === 'row' ? m.rows : m.cols).keys()].sort((a, b) => a - b);
    if (!list.length) return null;
    const at = kind === 'row' ? this.sel.row : this.sel.column;
    const index = dir > 0 ? (list.find((i) => i > at) ?? list[0]) : ([...list].reverse().find((i) => i < at) ?? list[list.length - 1]);
    return { kind, index };
  }

  /** Метки сброшены (новая книга). */
  clearMarks(): void { this.marks.clear(); }

  // --- Геометрия ---

  private columnWidth(column: number): number {
    if (this.resizing?.kind === 'column' && this.resizing.index === column) return this.resizing.size;
    try {
      return Math.max(0, Math.round(this.model().getColumnWidth(this.sheet, column)));
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

  // Высоты строк: у большинства строк высота стандартная, поэтому храним только
  // нестандартные (номера по возрастанию) и считаем положение любой строки
  // двоичным поиском — без вопросов к движку о каждой из миллиона строк.
  private customRows: number[] = [];
  private customH: number[] = [];
  /** acc[k] — сумма отклонений от стандартной высоты у первых k нестандартных строк. */
  private customAcc: number[] = [0];
  private layoutLoaded = false;

  private loadRowLayout(): void {
    let flat: ArrayLike<number> = [];
    try { flat = this.model().getRowLayout(this.sheet); } catch { /* нет листа */ }
    const pairs: [number, number][] = [];
    for (let i = 0; i + 1 < flat.length; i += 2) {
      const h = Math.max(0, Math.round(flat[i + 1] * ROW_SCALE));
      if (h !== ROW_H) pairs.push([flat[i], h]);
    }
    pairs.sort((a, b) => a[0] - b[0]);
    this.customRows = pairs.map((p) => p[0]);
    this.customH = pairs.map((p) => p[1]);
    this.customAcc = [0];
    for (const [, h] of pairs) this.customAcc.push(this.customAcc[this.customAcc.length - 1] + h - ROW_H);
    this.layoutLoaded = true;
  }

  /** Сколько нестандартных строк выше строки row. */
  private customBefore(row: number): number {
    if (!this.layoutLoaded) this.loadRowLayout();
    let lo = 0, hi = this.customRows.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.customRows[mid] < row) lo = mid + 1; else hi = mid;
    }
    return lo;
  }

  private baseHeight(row: number): number {
    const k = this.customBefore(row);
    return this.customRows[k] === row ? this.customH[k] : ROW_H;
  }

  /** Верхняя граница строки в пикселях листа (без заголовка столбцов). */
  private y(row: number): number {
    let v = (row - 1) * ROW_H + this.customAcc[this.customBefore(row)];
    const z = this.resizing;
    if (z?.kind === 'row' && z.index < row) v += z.size - this.baseHeight(z.index);
    return v;
  }

  private height(row: number): number {
    const z = this.resizing;
    if (z?.kind === 'row' && z.index === row) return z.size;
    return this.baseHeight(row);
  }

  private rowAt(py: number): number {
    let lo = 1, hi = LAST_ROW;
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
    this.layoutLoaded = false;
    this.readFrozen();
    this.updateExtent();
    this.render();
  }

  private updateExtent(): void {
    // Выделен весь столбец — не растягиваем лист до миллионной строки.
    const last = this.sel.r1 === 1 && this.sel.r2 === LAST_ROW ? this.sel.row : this.sel.r2;
    this.extentRows = Math.min(LAST_ROW, Math.max(1000, last + 200, this.extentRows));
    const lastCol = this.sel.c1 === 1 && this.sel.c2 === LAST_COLUMN ? this.sel.column : this.sel.c2;
    this.extentCols = Math.min(LAST_COLUMN, Math.max(52, lastCol + 10, this.extentCols));
    this.spacer.style.height = `${(this.headerH + this.y(this.extentRows + 1)) * this.zoom}px`;
    this.spacer.style.width = `${(this.headerW + this.x(this.extentCols + 1)) * this.zoom}px`;
  }

  private resize(): void {
    const dpr = window.devicePixelRatio || 1;
    const w = this.scroller.clientWidth, h = this.scroller.clientHeight;
    this.canvas.width = Math.round(w * dpr);
    this.canvas.height = Math.round(h * dpr);
    this.canvas.style.width = `${w}px`;
    this.canvas.style.height = `${h}px`;
    this.ctx.setTransform(dpr * this.zoom, 0, 0, dpr * this.zoom, 0, 0);
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

  // --- Закрепление снизу и справа (ENS-Office; в Excel такого нет) ---
  // Строки pinRows всегда видны у нижнего края, столбцы pinCols — у правого
  // (итоги, суммы); середина листа прокручивается. Пока рисуется полоса
  // закреплённых, sy/sx переводят их строки/столбцы в эту полосу (pinMode*).
  pinRows: { a: number; b: number } | null = null;
  pinCols: { a: number; b: number } | null = null;
  private pinModeRow = false;
  private pinModeCol = false;

  setPins(rows: { a: number; b: number } | null, cols: { a: number; b: number } | null): void {
    const same = (x: { a: number; b: number } | null, y: { a: number; b: number } | null) => (x?.a ?? 0) === (y?.a ?? 0) && (x?.b ?? 0) === (y?.b ?? 0);
    if (same(rows, this.pinRows) && same(cols, this.pinCols)) return;
    this.pinRows = rows;
    this.pinCols = cols;
    this.render();
  }

  /** Высота полосы закреплённых снизу (не больше половины окна). */
  private get bottomH(): number {
    const p = this.pinRows;
    if (!p) return 0;
    return Math.min(this.y(p.b + 1) - this.y(p.a), Math.max(0, (this.viewH - this.headerH) / 2));
  }

  private get rightW(): number {
    const p = this.pinCols;
    if (!p) return 0;
    return Math.min(this.x(p.b + 1) - this.x(p.a), Math.max(0, (this.viewW - this.headerW) / 2));
  }

  /** Верх полосы закреплённых снизу / левый край полосы справа (px на canvas). */
  private get pinTop(): number { return this.viewH - this.bottomH; }
  private get pinLeft(): number { return this.viewW - this.rightW; }

  private inPinRows(row: number): boolean { return !!this.pinRows && row >= this.pinRows.a && row <= this.pinRows.b + 1; }
  private inPinCols(column: number): boolean { return !!this.pinCols && column >= this.pinCols.a && column <= this.pinCols.b + 1; }

  /** Первая строка прокручиваемой области. */
  private get firstRow(): number {
    return Math.max(this.fr + 1, this.rowAt(this.sTop + this.frozenH));
  }

  private get firstColumn(): number {
    return Math.max(this.fc + 1, this.columnAt(this.sLeft + this.frozenW));
  }

  private visibleRows(): number {
    const top = this.firstRow;
    let n = 0;
    while (top + n < LAST_ROW && this.y(top + n + 1) - this.sTop < this.viewH - this.headerH - this.bottomH) n++;
    return n;
  }

  /** Экранные координаты: закреплённые строки и столбцы не прокручиваются. */
  private sy(row: number): number {
    if (this.pinModeRow && this.inPinRows(row)) return this.pinTop + this.y(row) - this.y(this.pinRows!.a);
    return this.headerH + this.y(row) - (row <= this.fr ? 0 : this.sTop);
  }

  private sx(column: number): number {
    if (this.pinModeCol && this.inPinCols(column)) return this.pinLeft + this.x(column) - this.x(this.pinCols!.a);
    return this.headerW + this.x(column) - (column <= this.fc ? 0 : this.sLeft);
  }

  /** Координаты листа (px) под точкой экрана в области ячеек. */
  private sheetY(py: number): number {
    if (this.pinRows && py >= this.pinTop) return this.y(this.pinRows.a) + py - this.pinTop;
    const b = py - this.headerH;
    return b < this.frozenH ? b : b + this.sTop;
  }

  private sheetX(px: number): number {
    if (this.pinCols && px >= this.pinLeft) return this.x(this.pinCols.a) + px - this.pinLeft;
    const b = px - this.headerW;
    return b < this.frozenW ? b : b + this.sLeft;
  }

  /** Прямоугольник ячейки на экране (в координатах canvas). */
  private lrect(row: number, column: number): { x: number; y: number; w: number; h: number } {
    return { x: this.sx(column), y: this.sy(row), w: this.width(column), h: this.height(row) };
  }

  /** Прямоугольник ячейки в пикселях окна (с учётом масштаба): для полей ввода и подсказок. */
  cellRect(row: number, column: number): { x: number; y: number; w: number; h: number } {
    const r = this.lrect(row, column), z = this.zoom;
    return { x: r.x * z, y: r.y * z, w: r.w * z, h: r.h * z };
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
    if (!extend && !this.ctrlPick) this.extraRanges = [];
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
    if (!this.ctrlPick) this.extraRanges = [];
    this.sel = { row: r1, column: c1, r1, c1, r2, c2 };
    this.corner = { row: r2, column: c2 };
    this.ensureVisible(r1, c1);
    this.updateExtent();
    this.render();
    this.onSelect(this.sel);
  }

  /** Выделить целые столбцы (строки) от anchor до to, не прокручивая лист к началу. */
  selectLines(kind: 'row' | 'column', anchor: number, to: number): void {
    if (!this.ctrlPick && !this.headerDrag) this.extraRanges = [];
    const a = Math.min(anchor, to), b = Math.max(anchor, to);
    const s = this.sel;
    if (kind === 'column') {
      const row = s.r1 === 1 && s.r2 === LAST_ROW ? s.row : this.firstRow;
      this.sel = { row, column: anchor, r1: 1, c1: a, r2: LAST_ROW, c2: b };
      this.corner = { row: LAST_ROW, column: to };
      this.ensureVisible(row, to);
    } else {
      const column = s.c1 === 1 && s.c2 === LAST_COLUMN ? s.column : this.firstColumn;
      this.sel = { row: anchor, column, r1: a, c1: 1, r2: b, c2: LAST_COLUMN };
      this.corner = { row: to, column: LAST_COLUMN };
      this.ensureVisible(to, column);
    }
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
    // Скрытые (отфильтрованные) строки и столбцы пропускаем, как Excel.
    for (let i = 0; i < 100_000 && dRow && row > 1 && row < LAST_ROW && this.height(row) === 0; i++) row += Math.sign(dRow);
    for (let i = 0; i < 20_000 && dCol && column > 1 && column < LAST_COLUMN && this.width(column) === 0; i++) column += Math.sign(dCol);
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
    const r = this.lActiveRect(), z = this.zoom;
    return { x: r.x * z, y: r.y * z, w: r.w * z, h: r.h * z };
  }

  private lActiveRect(): { x: number; y: number; w: number; h: number } {
    // Активная ячейка в закреплённой снизу/справа полосе — рамка там.
    const pr = this.pinModeRow, pc = this.pinModeCol;
    this.pinModeRow = this.inPinRows(this.sel.row) && this.sel.row <= this.pinRows!.b;
    this.pinModeCol = this.inPinCols(this.sel.column) && this.sel.column <= this.pinCols!.b;
    try {
      const m = this.mergeOf(this.sel.row, this.sel.column);
      if (!m) return this.lrect(this.sel.row, this.sel.column);
      const a = this.lrect(m.row, m.column), b = this.lrect(m.row + m.height - 1, m.column + m.width - 1);
      return { x: a.x, y: a.y, w: b.x + b.w - a.x, h: b.y + b.h - a.y };
    } finally {
      this.pinModeRow = pr;
      this.pinModeCol = pc;
    }
  }

  pageRows(): number {
    return Math.max(1, this.visibleRows() - 1);
  }

  ensureVisible(row: number, column: number): void {
    const z = this.zoom;
    // Закреплённые снизу/справа и так видны — лист не прокручиваем.
    if (row > this.fr && !(this.pinRows && row >= this.pinRows.a && row <= this.pinRows.b)) {
      const top = this.y(row) - this.frozenH, rh = this.height(row);
      const viewH = this.viewH - this.headerH - this.frozenH - this.bottomH;
      if (top < this.sTop) this.scroller.scrollTop = top * z;
      else if (top + rh > this.sTop + viewH) this.scroller.scrollTop = (top + rh - viewH) * z;
    }
    if (column > this.fc && !(this.pinCols && column >= this.pinCols.a && column <= this.pinCols.b)) {
      const left = this.x(column) - this.frozenW, cw = this.width(column);
      const viewW = this.viewW - this.headerW - this.frozenW - this.rightW;
      if (left < this.sLeft) this.scroller.scrollLeft = left * z;
      else if (left + cw > this.sLeft + viewW) this.scroller.scrollLeft = (left + cw - viewW) * z;
    }
  }

  private hit(e: MouseEvent): { row: number; column: number; header: 'row' | 'column' | 'corner' | null } {
    const r = this.canvas.getBoundingClientRect();
    const px = (e.clientX - r.left) / this.zoom, py = (e.clientY - r.top) / this.zoom;
    const row = py < this.headerH ? 0 : this.rowAt(this.sheetY(py));
    const column = px < this.headerW ? 0 : this.columnAt(this.sheetX(px));
    // Полоса структуры — не заголовок (её щелчки обрабатывает outlineClick).
    const header = row === 0 && column === 0 ? 'corner' : row === 0 ? 'column' : column === 0 ? 'row' : null;
    return { row, column, header };
  }

  /** Курсор у границы заголовка: какой столбец/строку тянуть. */
  private grip(e: MouseEvent): { kind: 'column' | 'row'; index: number } | null {
    const r = this.canvas.getBoundingClientRect();
    const px = (e.clientX - r.left) / this.zoom, py = (e.clientY - r.top) / this.zoom;
    if (py < this.headerH && px > this.headerW) {
      const sx = this.sheetX(px);
      const c = this.columnAt(sx);
      if (this.x(c + 1) - sx <= GRIP) return { kind: 'column', index: c };
      if (sx - this.x(c) <= GRIP && c > 1) return { kind: 'column', index: c - 1 };
    } else if (px < this.headerW && py > this.headerH) {
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
    {
      const cr0 = this.canvas.getBoundingClientRect();
      if (this.outlineClick((e.clientX - cr0.left) / this.zoom, (e.clientY - cr0.top) / this.zoom)) { e.preventDefault(); return; }
    }
    if (this.images.length || this.charts.length) {
      const cr1 = this.canvas.getBoundingClientRect();
      const px = (e.clientX - cr1.left) / this.zoom, py = (e.clientY - cr1.top) / this.zoom;
      const sel = this.selObj >= 0 ? this.imageRect(this.selObj) : null;
      const onHandle = sel && Math.abs(px - (sel.x + sel.w)) <= 7 && Math.abs(py - (sel.y + sel.h)) <= 7;
      const hit = onHandle ? this.selObj : this.imageAt(px, py);
      if (hit >= 0 && px > this.headerW && py > this.headerH) {
        e.preventDefault();
        this.selObj = hit;
        const orig = { ...this.imageRect(hit) };
        this.imgDrag = { mode: onHandle ? 'resize' : 'move', i: hit, sx: px, sy: py, orig, cur: orig };
        this.onImageSelect(hit);
        this.render();
        return;
      }
      if (this.selObj >= 0) { this.selObj = -1; this.onImageSelect(-1); this.render(); }
    }
    const g = this.grip(e);
    if (g) {
      e.preventDefault();
      const size = g.kind === 'column' ? this.width(g.index) : this.height(g.index);
      this.resizing = { ...g, start: g.kind === 'column' ? e.clientX : e.clientY, startSize: size, size };
      document.body.style.cursor = g.kind === 'column' ? 'col-resize' : 'row-resize';
      return;
    }
    const cr = this.canvas.getBoundingClientRect();
    if (this.filters.length) {
      const h0 = this.hit(e);
      const f = this.filters.find((t) => h0.row === t.r1 && h0.column >= t.c1 && h0.column <= t.c2);
      if (f) {
        const rect = this.lrect(h0.row, h0.column);
        if ((e.clientX - cr.left) / this.zoom >= rect.x + rect.w - 18) {
          e.preventDefault();
          this.onFilterButton(f.r1, h0.column, cr.left + (rect.x + rect.w - 18) * this.zoom, cr.top + (rect.y + rect.h) * this.zoom);
          return;
        }
      }
    }
    if (Math.abs((e.clientX - cr.left) / this.zoom - this.handle.x) <= 5 && Math.abs((e.clientY - cr.top) / this.zoom - this.handle.y) <= 5) {
      // Маркер заполнения: протянуть формулу или ряд (1, 2 → 3, 4…).
      e.preventDefault();
      const s = this.sel;
      this.filling = { r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 };
      document.body.style.cursor = 'crosshair';
      return;
    }
    const frame = this.refFrameAt(e);
    if (frame) {
      e.preventDefault(); // фокус остаётся в поле ввода формулы
      const p = this.hit(e);
      this.refDrag = { ...frame, row: p.row, column: p.column };
      return;
    }
    const h = this.hit(e);
    if (!h.header && h.row >= 1 && h.column >= 1 && this.onPointer(h.row, h.column, 'down', e.shiftKey)) {
      e.preventDefault(); // фокус остаётся в поле ввода формулы
      this.pointing = true;
      return;
    }
    // Ввод завершаем сразу, а не после потери фокуса: иначе прежняя ячейка
    // выделяется заново и протяжка превращает щелчок в диапазон.
    if (!this.beforeSelect()) { e.preventDefault(); return; }
    this.scroller.focus();
    // Ctrl+щелчок: прежнее выделение остаётся, новое добавляется к нему.
    const ctrl = (e.ctrlKey || e.metaKey) && !e.shiftKey;
    if (ctrl) {
      const s = this.sel;
      this.extraRanges = [...this.extraRanges, { r1: s.r1, c1: s.c1, r2: s.r2, c2: s.c2 }];
      this.ctrlPick = true;
    }
    if (h.header === 'column' || h.header === 'row') {
      // Несколько столбцов (строк): протянуть по заголовкам или Shift+щелчок.
      const kind = h.header, s = this.sel;
      const at = kind === 'column' ? h.column : h.row;
      const anchor = e.shiftKey ? (kind === 'column' ? s.column : s.row) : at;
      this.selectLines(kind, anchor, at);
      this.headerDrag = { kind, anchor };
    } else if (h.header === 'corner') this.selectRange(1, 1, LAST_ROW, LAST_COLUMN);
    else {
      this.select(h.row, h.column, e.shiftKey);
      this.dragging = true;
    }
    this.ctrlPick = false;
  }

  private endImageDrag(): void {
    const d = this.imgDrag;
    if (!d) return;
    this.imgDrag = null;
    const c = d.cur, moved = Math.abs(c.x - d.orig.x) + Math.abs(c.y - d.orig.y) + Math.abs(c.w - d.orig.w) + Math.abs(c.h - d.orig.h) > 1;
    if (!moved) { this.render(); return; }
    const row = Math.max(1, this.rowAt(this.sheetY(Math.max(c.y, this.headerH + 1))));
    const column = Math.max(1, this.columnAt(this.sheetX(Math.max(c.x, this.headerW + 1))));
    const r = this.lrect(row, column);
    const o = this.objs()[d.i];
    if (o) this.onImagePlace(o.kind, o.index, row, column, Math.round(c.x - r.x), Math.round(c.y - r.y), Math.round(c.w), Math.round(c.h));
    this.render();
  }

  private mouseMove(e: MouseEvent): void {
    const d = this.imgDrag;
    if (d) {
      const cr = this.canvas.getBoundingClientRect();
      const px = (e.clientX - cr.left) / this.zoom, py = (e.clientY - cr.top) / this.zoom;
      d.cur = d.mode === 'move'
        ? { ...d.orig, x: d.orig.x + px - d.sx, y: d.orig.y + py - d.sy }
        : { ...d.orig, w: Math.max(12, d.orig.w + px - d.sx), h: Math.max(12, d.orig.h + py - d.sy) };
      this.render();
      return;
    }
    const z = this.resizing;
    if (z) {
      const pos = z.kind === 'column' ? e.clientX : e.clientY;
      // До нуля — строка (столбец) скрывается, как в Excel.
      z.size = Math.max(0, Math.round(z.startSize + (pos - z.start) / this.zoom));
      this.colX = [0];
      this.render();
      return;
    }
    if (this.refDrag) {
      const d = this.refDrag, p = this.hit(e);
      if (p.row < 1 || p.column < 1) return;
      let rect = d.orig;
      if (d.mode === 'resize') {
        rect = { ...d.orig, r2: Math.max(d.orig.r1, p.row), c2: Math.max(d.orig.c1, p.column) };
      } else {
        const dr = Math.max(1 - d.orig.r1, Math.min(LAST_ROW - d.orig.r2, p.row - d.row));
        const dc = Math.max(1 - d.orig.c1, Math.min(LAST_COLUMN - d.orig.c2, p.column - d.column));
        rect = { r1: d.orig.r1 + dr, c1: d.orig.c1 + dc, r2: d.orig.r2 + dr, c2: d.orig.c2 + dc };
      }
      this.onRefDrag(d.id, rect);
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
      const onHandle = Math.abs((e.clientX - cr.left) / this.zoom - this.handle.x) <= 5 && Math.abs((e.clientY - cr.top) / this.zoom - this.handle.y) <= 5;
      const frame = g || onHandle ? null : this.refFrameAt(e);
      this.canvas.style.cursor = g ? (g.kind === 'column' ? 'col-resize' : 'row-resize') : onHandle ? 'crosshair'
        : frame ? (frame.mode === 'resize' ? 'nwse-resize' : 'move') : '';
    }
    if (e.target === this.canvas && !this.dragging && !this.pointing && this.commentCells.size) {
      const p = this.hit(e);
      this.onCellHover(p.row >= 1 && p.column >= 1 ? p.row : 0, p.column, e.clientX, e.clientY);
    }
    if (this.pointing) {
      const p = this.hit(e);
      if (p.row >= 1 && p.column >= 1) this.onPointer(p.row, p.column, 'drag', true);
      return;
    }
    if (this.headerDrag) {
      const p = this.hit(e), d = this.headerDrag;
      const at = d.kind === 'column' ? p.column : p.row;
      if (at >= 1) this.selectLines(d.kind, d.anchor, at);
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
    const W = this.viewW, H = this.viewH;
    const sl = this.sLeft, st = this.sTop;
    this.readFrozen();
    this.readDecor(model);
    const r0 = this.firstRow;
    let rN = r0;
    const bottomH = this.bottomH, rightW = this.rightW;
    while (rN < LAST_ROW && this.headerH + this.y(rN + 1) - st < H - bottomH) rN++;
    const c0 = this.firstColumn;
    let cN = c0;
    while (cN < LAST_COLUMN && this.headerW + this.x(cN + 1) - sl < W - rightW) cN++;
    // Полосы: закреплённые сверху/слева, прокручиваемые и закреплённые снизу/справа (третий элемент — true).
    const rowBands: Band[] = this.fr ? [[1, this.fr], [r0, rN]] : [[r0, rN]];
    const colBands: Band[] = this.fc ? [[1, this.fc], [c0, cN]] : [[c0, cN]];
    if (this.pinRows && bottomH > 0) rowBands.push([this.pinRows.a, this.pinRows.b, true]);
    if (this.pinCols && rightW > 0) colBands.push([this.pinCols.a, this.pinCols.b, true]);
    const shown = (r: number, c: number) => (r <= this.fr || (r >= r0 && r <= rN)) && (c <= this.fc || (c >= c0 && c <= cN));
    const s = this.sel;
    const gridLines = (() => { try { return model.getShowGridLines(this.sheet); } catch { return true; } })();

    // Ширина заголовков строк — по количеству цифр.
    const hw = Math.max(40, 12 + 8 * String(rN).length) + this.outlineW;
    if (hw !== this.headerW) { this.headerW = hw; this.updateExtent(); }

    ctx.fillStyle = this.theme.sheetBg;
    ctx.fillRect(0, 0, W, H);
    ctx.save();
    ctx.beginPath();
    ctx.rect(this.headerW, this.headerH, W - this.headerW, H - this.headerH);
    ctx.clip();

    // Стили и объединения читаем один раз за кадр.
    const styles = new Map<number, ExtendedStyle | null>();
    const extOf = (r: number, c: number): ExtendedStyle | null => {
      const key = r * 16_385 + c;
      let st = styles.get(key);
      if (st === undefined) {
        try { st = model.getCellStyle(this.sheet, r, c); } catch { st = null; }
        styles.set(key, st);
      }
      return st;
    };
    const styleOf = (r: number, c: number): CellStyle | null => extOf(r, c)?.style ?? null;
    let merges: MergedCell[] = [];
    try { merges = model.getMergedCells(this.sheet) ?? []; } catch { merges = []; }
    // Объединения → по ячейкам видимых полос: поиск за один шаг (на листах с тысячами объединений
    // перебор списка для каждой ячейки тормозил отрисовку).
    const mergeIdx = new Map<number, MergedCell>();
    for (const m of merges) {
      const mr2 = m.row + m.height - 1, mc2 = m.column + m.width - 1;
      for (const [ra, rb] of rowBands) for (const [ca, cb] of colBands) {
        const r1 = Math.max(m.row, ra), r2 = Math.min(mr2, rb), c1 = Math.max(m.column, ca), c2 = Math.min(mc2, cb);
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) mergeIdx.set(r * 16_385 + c, m);
      }
    }
    const mergeAt = (r: number, c: number) => mergeIdx.get(r * 16_385 + c);
    const textOf = (r: number, c: number) => { try { return model.getFormattedCellValue(this.sheet, r, c); } catch { return ''; } };

    for (const [ra, rb, pinR] of rowBands) for (const [ca, cb, pinC] of colBands) {
    this.pinModeRow = !!pinR;
    this.pinModeCol = !!pinC;
    ctx.save();
    ctx.beginPath();
    // Прокручиваемая часть не заходит под закреплённые (сверху/слева и снизу/справа).
    const left = pinC ? this.pinLeft : ca <= this.fc ? this.sx(ca) : Math.max(this.sx(ca), this.headerW + this.frozenW);
    const top = pinR ? this.pinTop : ra <= this.fr ? this.sy(ra) : Math.max(this.sy(ra), this.headerH + this.frozenH);
    const right = pinC ? W : Math.min(this.sx(cb) + this.width(cb), W - rightW);
    const bottom = pinR ? H : Math.min(this.sy(rb) + this.height(rb), H - bottomH);
    ctx.rect(left, top, right - left, bottom - top);
    ctx.clip();

    // Линии сетки — под заливкой: залитые ячейки их скрывают, как в Excel.
    if (gridLines) {
      ctx.strokeStyle = this.theme.gridLine;
      ctx.lineWidth = 1;
      ctx.beginPath();
      for (let r = ra; r <= rb + 1; r++) {
        const y = Math.round(this.sy(r)) + 0.5;
        ctx.moveTo(this.headerW, y); ctx.lineTo(W, y);
      }
      for (let c = ca; c <= cb + 1; c++) {
        const x = Math.round(this.sx(c)) + 0.5;
        ctx.moveTo(x, this.headerH); ctx.lineTo(x, H);
      }
      ctx.stroke();
    }

    // Ячейки: обычные и объединённые (одна большая клетка с данными левой верхней).
    const cells: { r: number; c: number; x: number; y: number; w: number; h: number; merged: boolean }[] = [];
    for (let r = ra; r <= rb; r++) {
      for (let c = ca; c <= cb; c++) {
        if (mergeAt(r, c)) continue;
        cells.push({ r, c, ...this.lrect(r, c), merged: false });
      }
    }
    for (const m of merges) {
      const r2 = m.row + m.height - 1, c2 = m.column + m.width - 1;
      if (r2 < ra || m.row > rb || c2 < ca || m.column > cb) continue;
      const a = this.lrect(m.row, m.column), b = this.lrect(r2, c2);
      cells.push({ r: m.row, c: m.column, x: a.x, y: a.y, w: b.x + b.w - a.x, h: b.y + b.h - a.y, merged: true });
    }

    // Заливка (объединённая ячейка без заливки — белая: внутри нет линий сетки).
    for (const cell of cells) {
      const color = styleOf(cell.r, cell.c)?.fill?.color;
      const fill = color ? model.resolveColor(color) : cell.merged ? this.theme.sheetBg : '';
      if (fill) { ctx.fillStyle = fill; ctx.fillRect(cell.x, cell.y, cell.w, cell.h); }
    }

    // Ещё выделенные диапазоны (Ctrl+щелчок) — той же заливкой.
    for (const x of this.extraRanges) {
      if (x.r2 < ra || x.r1 > rb || x.c2 < ca || x.c1 > cb) continue;
      const a = this.lrect(Math.max(x.r1, ra), Math.max(x.c1, ca));
      const b = this.lrect(Math.min(x.r2, rb), Math.min(x.c2, cb));
      ctx.fillStyle = this.theme.selectionFill;
      ctx.fillRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y);
    }
    // Выделенный диапазон (кроме активной ячейки).
    if ((s.r1 !== s.r2 || s.c1 !== s.c2) && s.r2 >= ra && s.r1 <= rb && s.c2 >= ca && s.c1 <= cb) {
      const a = this.lrect(Math.max(s.r1, ra), Math.max(s.c1, ca));
      const b = this.lrect(Math.min(s.r2, rb), Math.min(s.c2, cb));
      ctx.fillStyle = this.theme.selectionFill;
      ctx.fillRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y);
    }

    // Гистограммы условного форматирования — под текстом.
    for (const cell of cells) {
      const bar = extOf(cell.r, cell.c)?.data_bar;
      if (bar) this.drawDataBar(bar, cell);
    }

    // Значки условного форматирования (наборы значков и оценки) — слева в ячейке.
    const iconW = new Map<number, number>();
    for (const cell of cells) {
      const ext = extOf(cell.r, cell.c);
      if (!ext?.icon && !ext?.rating) continue;
      const size = Math.max(8, Math.min(14, cell.h - 6));
      const cy = cell.y + cell.h / 2;
      if (ext.icon) {
        this.drawCfIcon(ext.icon.icon, model.resolveColor(ext.icon.color ?? '#808080') || '#808080', cell.x + 3, cy, size, true);
        iconW.set(cell.r * 16_385 + cell.c, size + 5);
      } else if (ext.rating) {
        const color = model.resolveColor(ext.rating.color ?? '#FFC000') || '#FFC000';
        for (let i = 0; i < ext.rating.max; i++) this.drawCfIcon(ext.rating.icon, color, cell.x + 3 + i * (size + 1), cy, size, i < ext.rating.count);
        iconW.set(cell.r * 16_385 + cell.c, ext.rating.max * (size + 1) + 4);
      }
    }

    // Текст.
    const richAt = new Map<number, RichRun[]>();
    try {
      const flat = model.richTextCellsIn(this.sheet, ra, ca, rb, cb);
      for (let i = 0; i < flat.length; i += 2) {
        const runs = model.getRichText(this.sheet, flat[i], flat[i + 1]);
        if (runs) richAt.set(flat[i] * 16_385 + flat[i + 1], runs);
      }
    } catch { /* без rich text */ }
    for (const cell of cells) {
      const ext = extOf(cell.r, cell.c);
      if (ext?.data_bar?.show_value === false || ext?.icon?.show_value === false || ext?.rating?.show_value === false) continue;
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
      const shift = iconW.get(cell.r * 16_385 + cell.c) ?? 0;
      const tc = shift ? { ...cell, x: cell.x + shift, w: Math.max(0, cell.w - shift) } : cell;
      // Цвет из числового формата: «#,##0;[Red]-#,##0» — отрицательные красным.
      let fc = 0;
      try { fc = model.getFormattedCellColor(this.sheet, cell.r, cell.c); } catch { fc = 0; }
      const shownStyle = fc > 0 && style ? { ...style, font: { ...style.font, color: PALETTE[fc] } } : style;
      const runs = richAt.get(cell.r * 16_385 + cell.c);
      if (runs && !style?.alignment?.text_rotation) this.drawRich(runs, shownStyle, align, tc, Math.max(spillL, tc.x), spillR);
      else this.drawText(text, shownStyle, align, tc, Math.max(spillL, tc.x), spillR);
    }

    // Границы — поверх всего.
    for (const cell of cells) {
      const b = styleOf(cell.r, cell.c)?.border;
      if (b) this.drawBorders(b, cell);
    }
    this.drawImages();
    // Красные уголки примечаний.
    if (this.commentCells.size) {
      ctx.save();
      ctx.fillStyle = '#d32f2f';
      for (const cell of cells) {
        if (!this.commentCells.has(`${cell.r}:${cell.c}`)) continue;
        const x = cell.x + cell.w, y = cell.y;
        ctx.beginPath(); ctx.moveTo(x - 7, y); ctx.lineTo(x, y); ctx.lineTo(x, y + 7); ctx.closePath(); ctx.fill();
      }
      ctx.restore();
    }
    // Ручные разрывы страниц: штриховые линии после строки (слева от столбца).
    if (this.pageBreaks[0].length || this.pageBreaks[1].length) {
      ctx.save();
      ctx.strokeStyle = '#1f4e9c';
      ctx.lineWidth = 1;
      ctx.setLineDash([5, 3]);
      const a = this.lrect(ra, ca), b = this.lrect(rb, cb);
      for (const n of this.pageBreaks[0]) {
        if (n < ra || n > rb) continue;
        const r = this.lrect(n, ca), y = Math.round(r.y + r.h) - 0.5;
        ctx.beginPath(); ctx.moveTo(a.x, y); ctx.lineTo(b.x + b.w, y); ctx.stroke();
      }
      for (const n of this.pageBreaks[1]) {
        if (n < ca || n > cb) continue;
        const r = this.lrect(ra, n), x = Math.round(r.x + r.w) - 0.5;
        ctx.beginPath(); ctx.moveTo(x, a.y); ctx.lineTo(x, b.y + b.h); ctx.stroke();
      }
      ctx.restore();
    }
    // Кнопки автофильтра в строке заголовков.
    for (const flt of this.filters) {
    if (flt.r1 >= ra && flt.r1 <= rb) {
      for (let c = Math.max(ca, flt.c1); c <= Math.min(cb, flt.c2); c++) {
        const r = this.lrect(flt.r1, c);
        const bx = r.x + r.w - 17, by = r.y + r.h - 17;
        ctx.fillStyle = flt.active.has(c) ? '#d3ebe5' : '#f3f3f3';
        ctx.strokeStyle = '#9a9a9a';
        ctx.lineWidth = 1;
        ctx.fillRect(bx, by, 16, 16);
        ctx.strokeRect(bx + 0.5, by + 0.5, 15, 15);
        ctx.fillStyle = '#333';
        ctx.beginPath();
        if (flt.active.has(c)) {
          // воронка: отбор включён
          ctx.moveTo(bx + 3, by + 4); ctx.lineTo(bx + 13, by + 4); ctx.lineTo(bx + 9, by + 9);
          ctx.lineTo(bx + 9, by + 13); ctx.lineTo(bx + 7, by + 12); ctx.lineTo(bx + 7, by + 9); ctx.closePath();
        } else {
          ctx.moveTo(bx + 4, by + 6); ctx.lineTo(bx + 12, by + 6); ctx.lineTo(bx + 8, by + 11); ctx.closePath();
        }
        ctx.fill();
      }
    }
    }
    ctx.restore();
    }
    this.pinModeRow = this.pinModeCol = false;
    if (this.focusCell.enabled) this.drawFocusCell(r0, rN, c0, cN);

    // Активная ячейка: толстая рамка; в неактивном окне — бледнее.
    const border = this.windowActive ? this.theme.activeBorder : this.theme.activeBorderInactive;
    if (shown(s.row, s.column)) {
      const { x, y, w, h } = this.lActiveRect();
      const lw = this.theme.activeWidth;
      ctx.strokeStyle = border;
      ctx.lineWidth = lw;
      ctx.strokeRect(x + lw / 2 - 1, y + lw / 2 - 1, w - lw + 2, h - lw + 2);
    }
    // Рамка диапазона (не для одной объединённой ячейки).
    const am = this.mergeOf(s.row, s.column);
    const onlyMerge = am && s.r1 === am.row && s.c1 === am.column && s.r2 === am.row + am.height - 1 && s.c2 === am.column + am.width - 1;
    if ((s.r1 !== s.r2 || s.c1 !== s.c2) && !onlyMerge) {
      const a = this.lrect(s.r1, s.c1), b = this.lrect(Math.min(s.r2, LAST_ROW), Math.min(s.c2, LAST_COLUMN));
      ctx.strokeStyle = border;
      ctx.lineWidth = 1.5;
      ctx.strokeRect(a.x, a.y, b.x + b.w - a.x, b.y + b.h - a.y);
    }
    // Маркер заполнения — в правом нижнем углу выделения, как в Excel.
    this.handle = { x: -100, y: -100 };
    if (shown(s.r2, s.c2) && s.r2 < LAST_ROW && s.c2 < LAST_COLUMN) {
      const b = this.lrect(s.r2, s.c2);
      this.handle = { x: b.x + b.w, y: b.y + b.h };
      ctx.fillStyle = this.theme.sheetBg;
      ctx.fillRect(this.handle.x - 4, this.handle.y - 4, 8, 8);
      ctx.fillStyle = border;
      ctx.fillRect(this.handle.x - 3, this.handle.y - 3, 6, 6);
    }
    // Протягивание и скопированный диапазон — пунктир.
    const dashed = (r: { r1: number; c1: number; r2: number; c2: number }, color: string) => {
      const a = this.lrect(r.r1, r.c1), b = this.lrect(r.r2, r.c2);
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
      const a = this.lrect(h.r1, h.c1), b = this.lrect(h.r2, h.c2);
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
      const a = this.lrect(p.r1, p.c1), b = this.lrect(p.r2, p.c2);
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
    if (this.fr) { const y = Math.round(this.headerH + this.frozenH) + 0.5; ctx.moveTo(this.headerW, y); ctx.lineTo(W, y); }
    if (this.fc) { const x = Math.round(this.headerW + this.frozenW) + 0.5; ctx.moveTo(x, this.headerH); ctx.lineTo(x, H); }
    ctx.stroke();
    // Граница закреплённых снизу/справа — потолще и цветом темы, чтобы было видно, где «итоги».
    if (bottomH || rightW) {
      ctx.strokeStyle = this.theme.activeBorderInactive;
      ctx.lineWidth = 2;
      ctx.beginPath();
      if (bottomH) { const y = Math.round(H - bottomH); ctx.moveTo(0, y); ctx.lineTo(W, y); }
      if (rightW) { const x = Math.round(W - rightW); ctx.moveTo(x, 0); ctx.lineTo(x, H); }
      ctx.stroke();
    }
    ctx.restore();

    this.drawHeaders(rowBands, colBands, W, H);
  }

  /** Текст ячейки: шрифт, выравнивание по горизонтали и вертикали, перенос по словам. */
  private drawText(text: string, style: CellStyle | null, align: string,
    cell: { x: number; y: number; w: number; h: number }, spillL: number, spillR: number): void {
    const ctx = this.ctx, model = this.model();
    const f = style?.font;
    const px = Math.round((f?.sz ?? 11) * 1.33);
    // Надстрочный и подстрочный текст — мельче и со сдвигом от базовой линии.
    const vert = (f as { vert_align?: string } | undefined)?.vert_align;
    const drawPx = vert === 'superscript' || vert === 'subscript' ? Math.round(px * 0.7) : px;
    const shift = vert === 'superscript' ? -Math.round(px * 0.3) : vert === 'subscript' ? Math.round(px * 0.2) : 0;
    ctx.font = `${f?.i ? 'italic ' : ''}${f?.b ? 'bold ' : ''}${drawPx}px ${fontStack(f?.name)}`;
    let textFill = f?.color ? model.resolveColor(f.color) || '#000' : this.theme.textColor ?? '#000';
    if (this.theme.textColor) {
      // Тёмная тема листа: «чёрный» из файла нечитаем — берём светлый; на светлой заливке ячейки — тёмный.
      const bg = style?.fill?.color ? model.resolveColor(style.fill.color) || this.theme.sheetBg : this.theme.sheetBg;
      if (Math.abs(luminance(bg) - luminance(textFill)) < 0.35) textFill = luminance(bg) < 0.5 ? this.theme.textColor : '#000';
    }
    ctx.fillStyle = textFill;
    // Направление текста (Excel textRotation): 1..90 — против часовой стрелки,
    // 91..180 — по часовой (90 + угол), 255 — буквы столбиком.
    const rot = style?.alignment?.text_rotation ?? 0;
    if (rot) {
      ctx.save();
      ctx.beginPath(); ctx.rect(cell.x + 1, cell.y + 1, cell.w - 2, cell.h - 2); ctx.clip();
      ctx.textBaseline = 'middle';
      ctx.textAlign = 'center';
      if (rot === 255) {
        const chars = [...text.replace(/\n/g, '')];
        const lh = Math.round(px * 1.1);
        const vertical = style?.alignment?.vertical ?? 'bottom';
        const total = chars.length * lh;
        let y = vertical === 'top' ? cell.y + 2 : vertical === 'center' ? cell.y + (cell.h - total) / 2 : cell.y + cell.h - total - 2;
        const cx = align === 'left' ? cell.x + px / 2 + 3 : align === 'right' ? cell.x + cell.w - px / 2 - 3 : cell.x + cell.w / 2;
        for (const ch of chars) { ctx.fillText(ch, cx, y + lh / 2); y += lh; }
      } else {
        const deg = rot <= 90 ? rot : -(rot - 90);
        const tw = ctx.measureText(text).width;
        const rad = (-deg * Math.PI) / 180;
        // Повёрнутая строка: по горизонтали — по выравниванию ячейки, по вертикали — по центру.
        const span = Math.abs(Math.cos(rad)) * tw;
        const cx = align === 'left' ? cell.x + 3 + span / 2 : align === 'right' ? cell.x + cell.w - 3 - span / 2 : cell.x + cell.w / 2;
        ctx.translate(cx, cell.y + cell.h / 2);
        ctx.rotate(rad);
        ctx.fillText(text.replace(/\n/g, ' '), 0, 0);
        if (f?.u) ctx.fillRect(-tw / 2, px / 2 - 1, tw, 1);
        if (f?.strike) ctx.fillRect(-tw / 2, 0, tw, 1);
      }
      ctx.restore();
      return;
    }
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
    lines.forEach((line, n) => {
      const ty = y0 + lineH / 2 + 1 + shift;
      if (align === 'justify' && wrap && n < lines.length - 1 && line.includes(' ')) {
        // «По ширине»: пробелы растягиваются так, чтобы строка заняла всю ширину (кроме последней).
        const words = line.split(' ');
        const free = cell.w - 6 - words.reduce((sum, w) => sum + ctx.measureText(w).width, 0);
        const gap = free / (words.length - 1);
        let wx = cell.x + 3;
        ctx.textAlign = 'left';
        for (const w of words) { ctx.fillText(w, wx, ty); wx += ctx.measureText(w).width + gap; }
      } else if (align === 'fill' && line) {
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
    });
    ctx.restore();
  }

  /** Текст ячейки с разным оформлением частей (rich text): слова с переносом, у каждой части свой шрифт. */
  private drawRich(runs: RichRun[], style: CellStyle | null, align: string,
    cell: { x: number; y: number; w: number; h: number }, spillL: number, spillR: number): void {
    const ctx = this.ctx, model = this.model();
    const f = style?.font;
    const basePx = Math.round((f?.sz ?? 11) * 1.33);
    const base = f?.color ? model.resolveColor(f.color) || '#000' : this.theme.textColor ?? '#000';
    interface Seg { text: string; font: string; color: string; u: boolean; strike: boolean; px: number; shift: number; w: number }
    const wrap = !!style?.alignment?.wrap_text;
    const lines: Seg[][] = [[]];
    const widths = [0];
    for (const run of runs) {
      const px = Math.round((run.sz ?? f?.sz ?? 11) * 1.33);
      const drawPx = run.vert ? Math.round(px * 0.7) : px;
      const font = `${run.i ? 'italic ' : ''}${run.b ? 'bold ' : ''}${drawPx}px ${fontStack(run.name ?? f?.name)}`;
      ctx.font = font;
      const shift = run.vert === 1 ? -Math.round(px * 0.3) : run.vert === 2 ? Math.round(px * 0.2) : 0;
      for (const part of run.text.split(/(\n|\s+|[^\s]+)/).filter((t) => t)) {
        if (part === '\n') { lines.push([]); widths.push(0); continue; }
        const w = ctx.measureText(part).width;
        const last = lines.length - 1;
        if (wrap && lines[last].length && widths[last] + w > cell.w - 6 && part.trim()) { lines.push([]); widths.push(0); }
        const cur = lines.length - 1;
        if (wrap && !lines[cur].length && !part.trim()) continue;
        lines[cur].push({ text: part, font, color: run.color ?? base, u: !!run.u, strike: !!run.strike, px: drawPx, shift, w });
        widths[cur] += w;
      }
    }
    const lh = lines.map((l) => Math.round(Math.max(basePx, ...l.map((s) => s.px)) * 1.22));
    const blockH = lh.reduce((a, b) => a + b, 0);
    const vertical = style?.alignment?.vertical ?? 'bottom';
    let y = vertical === 'top' ? cell.y + 2 : vertical === 'center' || vertical === 'justify' || vertical === 'distributed' ? cell.y + (cell.h - blockH) / 2
      : cell.y + cell.h - blockH - 2;
    ctx.save();
    ctx.beginPath(); ctx.rect(spillL + 1, cell.y, spillR - spillL - 2, cell.h); ctx.clip();
    ctx.textBaseline = 'middle';
    ctx.textAlign = 'left';
    lines.forEach((line, n) => {
      let x = align === 'right' ? cell.x + cell.w - 3 - widths[n] : align === 'center' || align === 'centercontinuous' ? cell.x + (cell.w - widths[n]) / 2 : cell.x + 3;
      const ty = y + lh[n] / 2 + 1;
      for (const s of line) {
        ctx.font = s.font; ctx.fillStyle = s.color;
        ctx.fillText(s.text, x, ty + s.shift);
        if (s.u) ctx.fillRect(x, ty + s.shift + s.px / 2 - 1, s.w, 1);
        if (s.strike) ctx.fillRect(x, ty + s.shift, s.w, 1);
        x += s.w;
      }
      y += lh[n];
    });
    ctx.restore();
  }

  /** Гистограмма условного форматирования (как в Excel: от оси, градиентом). */
  private drawDataBar(bar: NonNullable<ExtendedStyle['data_bar']>, cell: { x: number; y: number; w: number; h: number }): void {
    const ctx = this.ctx, model = this.model();
    const inner = cell.w - 4;
    const axis = cell.x + 2 + inner * (bar.axis_position ?? 0);
    const len = inner * Math.min(1, Math.abs(bar.value));
    const negative = bar.value < 0;
    const color = model.resolveColor((negative ? bar.negative_color : bar.positive_color) ?? '#638EC6') || '#638EC6';
    const x0 = negative ? axis - len : axis;
    const y0 = cell.y + 2, h = cell.h - 4;
    if (bar.is_gradient) {
      const g = ctx.createLinearGradient(x0, 0, x0 + len, 0);
      g.addColorStop(negative ? 1 : 0, color);
      g.addColorStop(negative ? 0 : 1, '#ffffff');
      ctx.fillStyle = g;
    } else {
      ctx.fillStyle = color;
    }
    ctx.fillRect(x0, y0, len, h);
    ctx.strokeStyle = color;
    ctx.lineWidth = 1;
    ctx.strokeRect(Math.round(x0) + 0.5, Math.round(y0) + 0.5, Math.max(0, Math.round(len) - 1), h - 1);
  }

  /** Метка «здесь скрыты строки/столбцы» на заголовке. */
  private hiddenMark(x: number, y: number, w: number, h: number): void {
    const ctx = this.ctx;
    ctx.save();
    ctx.fillStyle = this.theme.activeBorder;
    if (w > h) { ctx.fillRect(x, y, w, 1); ctx.fillRect(x, y + 2, w, 1); } else { ctx.fillRect(x - 1, y, 1, h); ctx.fillRect(x + 1, y, 1, h); }
    ctx.restore();
  }

  /** Значок условного форматирования (как в Excel, своим рисунком): x — левый край, cy — середина по высоте. */
  private drawCfIcon(name: string, color: string, x: number, cy: number, s: number, filled: boolean): void {
    const ctx = this.ctx;
    const h = s / 2, cx = x + h, t = cy - h, b = cy + h;
    ctx.save();
    ctx.fillStyle = color;
    ctx.strokeStyle = color;
    ctx.lineWidth = Math.max(1.5, s / 7);
    ctx.lineJoin = 'round';
    ctx.lineCap = 'round';
    ctx.beginPath();
    const arrow = (angle: number) => {
      ctx.translate(cx, cy);
      ctx.rotate(angle);
      ctx.moveTo(0, -h); ctx.lineTo(h, 0); ctx.lineTo(h * 0.35, 0); ctx.lineTo(h * 0.35, h);
      ctx.lineTo(-h * 0.35, h); ctx.lineTo(-h * 0.35, 0); ctx.lineTo(-h, 0); ctx.closePath();
      ctx.fill();
    };
    const fillOrRing = () => { if (filled) ctx.fill(); else { ctx.lineWidth = 1; ctx.stroke(); } };
    switch (name) {
      case 'ArrowUp': arrow(0); break;
      case 'ArrowAngleUp': arrow(Math.PI / 4); break;
      case 'ArrowRight': arrow(Math.PI / 2); break;
      case 'ArrowAngleDown': arrow((3 * Math.PI) / 4); break;
      case 'ArrowDown': arrow(Math.PI); break;
      case 'TriangleUp': case 'TriangleUpFilled':
        ctx.moveTo(cx, t + 1); ctx.lineTo(x + s, b - 2); ctx.lineTo(x, b - 2); ctx.closePath(); ctx.fill(); break;
      case 'TriangleDown': case 'TriangleDownFilled':
        ctx.moveTo(x, t + 2); ctx.lineTo(x + s, t + 2); ctx.lineTo(cx, b - 1); ctx.closePath(); ctx.fill(); break;
      case 'FlatRectangle': ctx.fillRect(x, cy - s / 6, s, s / 3); break;
      case 'Rhombus': ctx.moveTo(cx, t); ctx.lineTo(x + s, cy); ctx.lineTo(cx, b); ctx.lineTo(x, cy); ctx.closePath(); ctx.fill(); break;
      case 'Flag':
        ctx.moveTo(x + 2, b); ctx.lineTo(x + 2, t); ctx.stroke();
        ctx.beginPath(); ctx.moveTo(x + 2, t); ctx.lineTo(x + s, t + s * 0.28); ctx.lineTo(x + 2, t + s * 0.56); ctx.closePath(); ctx.fill(); break;
      case 'Check':
        ctx.moveTo(x + 1, cy); ctx.lineTo(x + s * 0.4, b - 2); ctx.lineTo(x + s - 1, t + 2); ctx.stroke(); break;
      case 'Cross':
        ctx.moveTo(x + 2, t + 2); ctx.lineTo(x + s - 2, b - 2); ctx.moveTo(x + s - 2, t + 2); ctx.lineTo(x + 2, b - 2); ctx.stroke(); break;
      case 'Exclamation':
        ctx.moveTo(cx, t + 1); ctx.lineTo(cx, cy + s * 0.15); ctx.stroke();
        ctx.beginPath(); ctx.arc(cx, b - 1.5, ctx.lineWidth * 0.7, 0, Math.PI * 2); ctx.fill(); break;
      case 'Star':
        for (let i = 0; i < 10; i++) {
          const r = i % 2 ? h * 0.45 : h, a = -Math.PI / 2 + (i * Math.PI) / 5;
          const px = cx + r * Math.cos(a), py = cy + r * Math.sin(a);
          if (i) ctx.lineTo(px, py); else ctx.moveTo(px, py);
        }
        ctx.closePath();
        fillOrRing();
        break;
      case 'Heart':
        ctx.moveTo(cx, b - 1);
        ctx.bezierCurveTo(x - 1, cy, x + 1, t, cx, t + s * 0.3);
        ctx.bezierCurveTo(x + s - 1, t, x + s + 1, cy, cx, b - 1);
        fillOrRing();
        break;
      case 'ThumbsUp': case 'ThumbsDown':
        ctx.translate(cx, cy);
        if (name === 'ThumbsDown') ctx.rotate(Math.PI);
        ctx.fillRect(-h, -h * 0.1, h * 0.5, h * 1.1);
        ctx.fillRect(-h * 0.4, -h * 0.1, h * 1.3, h * 1.1);
        ctx.fillRect(-h * 0.3, -h, h * 0.55, h);
        break;
      default: // Circle и прочие — кружок
        ctx.arc(cx, cy, h - 0.5, 0, Math.PI * 2);
        fillOrRing();
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
      const auto = this.theme.textColor ?? '#000';
      let color = item.color ? model.resolveColor(item.color) || auto : auto;
      ctx.lineWidth = width;
      // Тонкие тёмные линии (другие программы ставят их по всей таблице) — мягче: не давят на данные.
      const dark = !item.color || (parseRgb(color) !== null && luminance(color) < 0.3);
      if (dark && this.theme.textColor) color = this.theme.textColor; // тёмная тема листа: «чёрные» границы светлые
      ctx.strokeStyle = color;
      if (dark && width === 1) ctx.globalAlpha = st === 'hair' ? 0.2 : 0.3; // грань соседних ячеек рисуется дважды
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
    const colRect = this.lrect(this.fr ? 1 : r0, s.column);
    const rowRect = this.lrect(s.row, this.fc ? 1 : c0);
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

  /** Положение рисунка на экране (в логических пикселях) с учётом переноса мышью. */
  /** Рисунки, затем диаграммы — единый список объектов поверх листа. */
  private objs(): { kind: 'image' | 'chart'; index: number; row: number; column: number; dx: number; dy: number; width: number; height: number }[] {
    return [
      ...this.images.map((o, index) => ({ ...o, kind: 'image' as const, index })),
      ...this.charts.map((o, index) => ({ ...o, kind: 'chart' as const, index })),
    ];
  }

  /** Вид объекта по номеру в общем списке (для главного окна). */
  objAt(i: number): { kind: 'image' | 'chart'; index: number } | null {
    const o = this.objs()[i];
    return o ? { kind: o.kind, index: o.index } : null;
  }

  private chartData(c: Chart): ChartData {
    const empty: ChartData = { cats: [], names: [], series: [] };
    const rg = parseRange(c.range);
    if (!rg) return empty;
    const [r1, c1, r2, c2] = rg, m = this.model();
    const cats: string[] = [];
    for (let r = r1 + 1; r <= r2; r++) cats.push(m.getFormattedCellValue(this.sheet, r, c1));
    const names: string[] = [], series: number[][] = [];
    const cols = c.kind === 2 ? 1 : c2 - c1;
    for (let k = 1; k <= cols; k++) {
      names.push(m.getFormattedCellValue(this.sheet, r1, c1 + k));
      series.push(Array.from(m.getRangeNumbers(this.sheet, r1 + 1, c1 + k, r2, c1 + k)));
    }
    return { cats, names, series };
  }

  private imageRect(i: number): { x: number; y: number; w: number; h: number } {
    if (this.imgDrag?.i === i) return this.imgDrag.cur;
    const im = this.objs()[i], c = this.lrect(im.row, im.column);
    return { x: c.x + im.dx, y: c.y + im.dy, w: im.width, h: im.height };
  }

  private imageAt(px: number, py: number): number {
    for (let i = this.objs().length - 1; i >= 0; i--) {
      const r = this.imageRect(i);
      if (px >= r.x && px <= r.x + r.w && py >= r.y && py <= r.y + r.h) return i;
    }
    return -1;
  }

  private drawImages(): void {
    const ctx = this.ctx, model = this.model();
    ctx.save();
    ctx.beginPath(); ctx.rect(this.headerW, this.headerH, this.viewW - this.headerW, this.viewH - this.headerH); ctx.clip();
    const all = this.objs();
    all.forEach((o, i) => {
      const r = this.imageRect(i);
      if (r.x > this.viewW || r.y > this.viewH || r.x + r.w < this.headerW || r.y + r.h < this.headerH) return;
      const drawFrame = () => {
        if (i !== this.selObj) return;
        ctx.strokeStyle = '#1f4e9c'; ctx.lineWidth = 2; ctx.strokeRect(r.x, r.y, r.w, r.h);
        ctx.fillStyle = '#fff'; ctx.lineWidth = 1;
        ctx.fillRect(r.x + r.w - 5, r.y + r.h - 5, 10, 10); ctx.strokeRect(r.x + r.w - 5.5, r.y + r.h - 5.5, 10, 10);
      };
      if (o.kind === 'chart') {
        try { drawChart(ctx, r, this.charts[o.index].kind, this.charts[o.index].title, this.chartData(this.charts[o.index]), this.charts[o.index]); } catch { /* диаграмма без данных */ }
        drawFrame();
        return;
      }
      const im = this.images[o.index];
      const key = `${this.sheet}:${o.index}:${im.sig}`;
      let el = this.imgCache.get(key);
      if (el === undefined) {
        this.imgCache.set(key, null);
        try {
          const img = new Image();
          img.onload = () => { this.imgCache.set(key, img); this.render(); };
          img.src = URL.createObjectURL(new Blob([model.getImageData(this.sheet, o.index) as BlobPart], { type: `image/${im.ext}` }));
        } catch { /* битый рисунок */ }
        el = null;
      }
      if (el) ctx.drawImage(el, r.x, r.y, r.w, r.h);
      else { ctx.strokeStyle = '#999'; ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w, r.h); }
      drawFrame();
    });
    ctx.restore();
  }

  /** Примечания и разрывы страниц листа — из модели (их мало, чтение дешёвое). */
  private readDecor(model: Model): void {
    this.commentCells.clear();
    try {
      for (const c of model.getComments(this.sheet)) {
        const m = /^([A-Z]+)(\d+)$/.exec(c.ref);
        if (m) this.commentCells.add(`${m[2]}:${m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)}`);
      }
      this.pageBreaks = model.getPageBreaks(this.sheet);
    } catch { this.pageBreaks = [[], []]; }
    try { this.images = model.getImages(this.sheet); } catch { this.images = []; }
    try { this.charts = model.getCharts(this.sheet); } catch { this.charts = []; }
    if (this.selObj >= this.images.length + this.charts.length) this.selObj = -1;
  }

  /** Значок метки: синий кружок. */
  private drawMark(x: number, y: number, color: string): void {
    const ctx = this.ctx;
    ctx.save();
    ctx.fillStyle = color;
    ctx.beginPath(); ctx.arc(x, y, 3.5, 0, Math.PI * 2); ctx.fill();
    ctx.restore();
  }

  private drawHeaders(rowBands: Band[], colBands: Band[], W: number, H: number): void {
    const ctx = this.ctx, s = this.sel;
    ctx.fillStyle = this.theme.headerBg;
    ctx.fillRect(0, 0, W, this.headerH);
    ctx.fillRect(0, 0, this.headerW, H);
    ctx.font = '12px "Segoe UI", sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.strokeStyle = this.theme.headerLine;
    ctx.lineWidth = 1;
    for (const [c0, cN, pinC] of colBands) {
    this.pinModeCol = !!pinC;
    // Прокручиваемые заголовки не заходят под закреплённые.
    const left = pinC ? this.pinLeft : c0 <= this.fc ? this.headerW : this.headerW + this.frozenW;
    const right = pinC ? W : W - this.rightW;
    ctx.save(); ctx.beginPath(); ctx.rect(left, 0, right - left, this.headerH); ctx.clip();
    for (let c = c0; c <= cN; c++) {
      const { x, w } = this.lrect(1, c);
      if (w === 0) continue; // скрытый столбец
      const inSel = (c >= s.c1 && c <= s.c2) || this.extraRanges.some((x) => c >= x.c1 && c <= x.c2);
      if (inSel) {
        ctx.fillStyle = this.theme.headerActiveFill;
        ctx.fillRect(x, this.outlineH, w, HEADER_H);
        ctx.fillStyle = this.windowActive ? this.theme.activeBorder : this.theme.activeBorderInactive;
        ctx.fillRect(x, this.headerH - 2, w, 2);
      }
      ctx.fillStyle = inSel ? this.theme.headerActiveText : this.theme.headerText;
      ctx.font = `${c === s.column ? 'bold ' : ''}12px "Segoe UI", sans-serif`;
      ctx.fillText(columnName(c), x + w / 2, this.outlineH + HEADER_H / 2);
      const mc = this.marksOf().cols.get(c);
      if (mc) this.drawMark(x + 7, this.outlineH + 5, mc);
      ctx.beginPath(); ctx.moveTo(Math.round(x + w) + 0.5, this.outlineH); ctx.lineTo(Math.round(x + w) + 0.5, this.headerH); ctx.stroke();
      // Перед этим столбцом скрытые — метка в заголовке (двойная черта, как в Excel).
      if (c > 1 && this.width(c - 1) === 0) this.hiddenMark(Math.round(x), this.outlineH, 3, HEADER_H);
    }
    ctx.restore();
    }
    this.pinModeCol = false;
    for (const [r0, rN, pinR] of rowBands) {
    this.pinModeRow = !!pinR;
    const top = pinR ? this.pinTop : r0 <= this.fr ? this.headerH : this.headerH + this.frozenH;
    const bottom = pinR ? H : H - this.bottomH;
    ctx.save(); ctx.beginPath(); ctx.rect(0, top, this.headerW, bottom - top); ctx.clip();
    for (let r = r0; r <= rN; r++) {
      const { y, h } = this.lrect(r, 1);
      if (h === 0) continue; // скрытая (отфильтрованная) строка
      const inSel = (r >= s.r1 && r <= s.r2) || this.extraRanges.some((x) => r >= x.r1 && r <= x.r2);
      if (inSel) {
        ctx.fillStyle = this.theme.headerActiveFill;
        ctx.fillRect(this.outlineW, y, this.headerW - this.outlineW, h);
        ctx.fillStyle = this.windowActive ? this.theme.activeBorder : this.theme.activeBorderInactive;
        ctx.fillRect(this.headerW - 2, y, 2, h);
      }
      // Как в Excel: номера строк под действующим фильтром — синие.
      const filtered = this.filters.some((f) => f.active.size > 0 && r > f.r1 && r <= f.r2);
      ctx.fillStyle = filtered ? '#2f6fd1' : inSel ? this.theme.headerActiveText : this.theme.headerText;
      ctx.font = `${r === s.row ? 'bold ' : ''}12px "Segoe UI", sans-serif`;
      ctx.fillText(String(r), this.outlineW + (this.headerW - this.outlineW) / 2, y + h / 2);
      const mr = this.marksOf().rows.get(r);
      if (mr) this.drawMark(this.outlineW + 6, y + h / 2, mr);
      if (r > 1 && this.height(r - 1) === 0) this.hiddenMark(this.outlineW, Math.round(y) - 1, this.headerW - this.outlineW, 3);
      ctx.beginPath(); ctx.moveTo(this.outlineW, Math.round(y + h) + 0.5); ctx.lineTo(this.headerW, Math.round(y + h) + 0.5); ctx.stroke();
    }
    ctx.restore();
    }
    this.pinModeRow = false;
    const z = this.resizing;
    if (z) {
      // Размер при перетаскивании границы — подсказка, как в Excel.
      const label = z.kind === 'column' ? `Ширина: ${z.size} пикс.` : `Высота: ${z.size} пикс.`;
      const at = z.kind === 'column' ? this.lrect(1, z.index) : this.lrect(z.index, 1);
      ctx.font = '12px "Segoe UI", sans-serif';
      const tw = ctx.measureText(label).width + 12;
      const bx = z.kind === 'column' ? Math.min(W - tw - 2, at.x + at.w + 6) : this.headerW + 6;
      const by = z.kind === 'column' ? this.headerH + 4 : at.y + at.h + 4;
      ctx.fillStyle = '#fffbe6'; ctx.strokeStyle = '#999';
      ctx.fillRect(bx, by, tw, 20); ctx.strokeRect(bx + 0.5, by + 0.5, tw, 20);
      ctx.fillStyle = '#222'; ctx.textAlign = 'left';
      ctx.fillText(label, bx + 6, by + 10);
      ctx.strokeStyle = 'rgba(0,0,0,0.45)'; ctx.setLineDash([3, 3]); ctx.beginPath();
      if (z.kind === 'column') { const lx = Math.round(at.x + at.w) + 0.5; ctx.moveTo(lx, 0); ctx.lineTo(lx, H); }
      else { const ly = Math.round(at.y + at.h) + 0.5; ctx.moveTo(0, ly); ctx.lineTo(W, ly); }
      ctx.stroke(); ctx.setLineDash([]);
    }
    ctx.fillStyle = this.theme.cornerBg;
    ctx.fillRect(0, 0, this.headerW, this.headerH);
    ctx.strokeStyle = this.theme.cornerLine;
    ctx.beginPath();
    ctx.moveTo(0, this.headerH + 0.5); ctx.lineTo(W, this.headerH + 0.5);
    ctx.moveTo(this.headerW + 0.5, 0); ctx.lineTo(this.headerW + 0.5, H);
    ctx.stroke();
    this.drawOutline(W, H);
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
