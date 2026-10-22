// Обёртка над движком ENS-Table (форк IronCalc, собранный в WebAssembly).

export interface Area {
  sheet: number;
  row: number;
  column: number;
  width: number;
  height: number;
}

export type Color = string | [number, number] | undefined;

export interface BorderItem { style: string; color?: Color }

export interface CellBorder { top?: BorderItem; bottom?: BorderItem; left?: BorderItem; right?: BorderItem; diagonal?: BorderItem }

export interface CellStyle {
  num_fmt: string;
  fill: { color?: Color };
  font: { b: boolean; i: boolean; u: boolean; strike: boolean; sz: number; color?: Color; name: string };
  alignment?: { horizontal?: string; vertical?: string; wrap_text?: boolean };
  border?: CellBorder;
}

/** Стиль ячейки с учётом условного форматирования. */
export interface ExtendedStyle {
  style: CellStyle;
  data_bar?: { positive_color?: Color; negative_color?: Color; is_gradient: boolean; value: number; axis_position: number; show_value: boolean } | null;
  /** Набор значков: какой значок и каким цветом. */
  icon?: { icon: string; color?: Color; show_value: boolean } | null;
  /** Оценка: count закрашенных значков из max. */
  rating?: { icon: string; count: number; max: number; color?: Color; show_value: boolean } | null;
}

/** Формат правила условного форматирования (разница к стилю ячейки). */
export interface Dxf {
  font?: { b?: boolean; i?: boolean; u?: boolean; strike?: boolean; color?: Color } | null;
  fill?: { color?: Color } | null;
  border?: unknown;
  num_fmt?: unknown;
  alignment?: unknown;
}

/** Правило условного форматирования, как его хранит движок (поле type — вид правила). */
export type CfRule = { type: string } & Record<string, unknown>;

export interface CfView { index: number; range: string; cf_rule: CfRule; priority: number }

export interface MergedCell { row: number; column: number; width: number; height: number }

/** Какие границы ставить на выделение. */
export type BorderType = 'All' | 'Inner' | 'Outer' | 'Top' | 'Right' | 'Bottom' | 'Left' | 'CenterH' | 'CenterV' | 'None';

export interface SheetProps {
  name: string;
  state: string;
  sheet_id: number;
  color?: string;
}

/** Скопированные ячейки во внутреннем виде движка: формулы, стили, ссылки. */
export interface EngineClipboard {
  csv: string;
  data: unknown;
  sheet: number;
  range: [number, number, number, number];
}

/** Методы движка, которыми пользуется интерфейс. */
export interface Model {
  setUserInput(sheet: number, row: number, column: number, input: string): void;
  getFormattedCellValue(sheet: number, row: number, column: number): string;
  getCellContent(sheet: number, row: number, column: number): string;
  getCellType(sheet: number, row: number, column: number): number;
  getColumnWidth(sheet: number, column: number): number;
  getRowHeight(sheet: number, row: number): number;
  getMergedCells(sheet: number): MergedCell[];
  /** Занятая часть листа: [первая строка, первый столбец, последняя строка, последний столбец]. */
  getSheetDimension(sheet: number): number[];
  /** Строки с нестандартной высотой: [строка, высота, …] (скрытые — 0). */
  getRowLayout(sheet: number): Float64Array;
  /** Сортировка по нескольким столбцам: keys = [столбец, по убыванию 0/1, …]; blanks: 0 — в конце, 1 — в начале, 2 — как ноль;
   *  visibleOnly — скрытые строки (отобранные фильтром) остаются на местах. */
  sortRangeBy(area: Area, keys: number[], blanks: number, visibleOnly: boolean): void;
  getAutoFilter(sheet: number): number[];
  setAutoFilter(sheet: number, range: number[]): void;
  sortRange(area: Area, keyColumn: number, descending: boolean): void;
  /** Условное форматирование: правила листа, сначала главные; index — для изменения и удаления. */
  getConditionalFormattingList(sheet: number): CfView[];
  getDxfForConditionalFormatting(sheet: number, index: number): Dxf | null;
  addConditionalFormatting(sheet: number, range: string, rule: CfRule): void;
  updateConditionalFormatting(sheet: number, index: number, range: string, rule: CfRule): void;
  deleteConditionalFormatting(sheet: number, index: number): void;
  raiseConditionalFormattingPriority(sheet: number, index: number): void;
  lowerConditionalFormattingPriority(sheet: number, index: number): void;
  findAll(text: string, matchCase: boolean, wholeCell: boolean, inFormulas: boolean, sheet: number): { sheet: number; row: number; column: number }[];
  replaceAll(text: string, replacement: string, matchCase: boolean, wholeCell: boolean, sheet: number, onlyRow: number, onlyColumn: number): number;
  pasteValues(sourceSheet: number, source: [number, number, number, number], transpose: boolean): void;
  onPasteStyles(styles: CellStyle[][]): void;
  setRowsHidden(sheet: number, rowStart: number, rowEnd: number, hidden: boolean): void;
  mergeCells(range: Area): void;
  mergeCellsCenter(range: Area): void;
  mergeCellsAcross(range: Area): void;
  unmergeCells(range: Area): void;
  setAreaWithBorder(area: Area, border: { item: BorderItem; type: BorderType }): void;
  setSelectedCell(row: number, column: number): void;
  setSelectedRange(startRow: number, startColumn: number, endRow: number, endColumn: number): void;
  copyToClipboard(): EngineClipboard;
  pasteFromClipboard(sourceSheet: number, sourceRange: [number, number, number, number], data: unknown, isCut: boolean): void;
  autoFillRows(source: Area, toRow: number): void;
  autoFillColumns(source: Area, toColumn: number): void;
  getFrozenRowsCount(sheet: number): number;
  getFrozenColumnsCount(sheet: number): number;
  setFrozenRowsCount(sheet: number, count: number): void;
  setFrozenColumnsCount(sheet: number, count: number): void;
  setColumnsWidth(sheet: number, columnStart: number, columnEnd: number, width: number): void;
  setRowsHeight(sheet: number, rowStart: number, rowEnd: number, height: number): void;
  insertRows(sheet: number, row: number, count: number): void;
  /** Склейка шагов отмены: всё после метки отменяется одним Ctrl+Z. */
  undoMark(): number;
  mergeUndoSince(mark: number): void;
  /** Цвет из числового формата ([Red]…): номер палитры Excel 1..56, 0 — нет. */
  getFormattedCellColor(sheet: number, row: number, column: number): number;
  /** Значение ячейки в формате code, книга не меняется (образец). */
  formatCellValueAs(sheet: number, row: number, column: number, code: string): string;
  /** Переставить строки (столбцы) ровно на delta: с высотой, шириной и скрытием; одна отмена. */
  moveRowsExact(sheet: number, row: number, count: number, delta: number): void;
  moveColumnsExact(sheet: number, column: number, count: number, delta: number): void;
  deleteRows(sheet: number, row: number, count: number): void;
  insertColumns(sheet: number, column: number, count: number): void;
  deleteColumns(sheet: number, column: number, count: number): void;
  getCellStyle(sheet: number, row: number, column: number): ExtendedStyle;
  updateRangeStyle(range: Area, path: string, value: string): void;
  rangeClearContents(sheet: number, r1: number, c1: number, r2: number, c2: number): void;
  pasteCsvText(area: Area, csv: string): void;
  undo(): void;
  redo(): void;
  canUndo(): boolean;
  canRedo(): boolean;
  newSheet(): void;
  setSelectedSheet(sheet: number): void;
  getRowsWithData(sheet: number, column: number): Int32Array;
  getColumnsWithData(sheet: number, row: number): Int32Array;
  setRecognizeDates(on: boolean): void;
  getRecognizeDates(): boolean;
  setLocale(locale: string): void;
  renameSheet(sheet: number, name: string): void;
  deleteSheet(sheet: number): void;
  duplicateSheet(sheet: number): void;
  moveSheet(sheet: number, newIndex: number): void;
  hideSheet(sheet: number): void;
  unhideSheet(sheet: number): void;
  setSheetColor(sheet: number, color: Color): void;
  getWorksheetsProperties(): SheetProps[];
  getShowGridLines(sheet: number): boolean;
  setShowGridLines(sheet: number, show: boolean): void;
  resolveColor(color: Color): string;
  toXlsx(): Uint8Array;
  toBytes(): Uint8Array;
  flushSendQueue(): Uint8Array;
  applyExternalDiffs(diffs: Uint8Array): void;
  evaluate(): void;
  free(): void;
}

// Собирается scripts/build-engine.sh (в git не хранится).
import init, { Model as EngineModel } from './engine-pkg/engine.js';

interface EngineStatic {
  new (name: string, locale: string, timezone: string, language: string): Model;
  fromXlsx(bytes: Uint8Array, name: string, locale: string, timezone: string, language: string): Model;
  fromBytes(bytes: Uint8Array, language: string): Model;
}

const Engine = EngineModel as unknown as EngineStatic;

export async function loadEngine(): Promise<void> {
  // Версия «одним файлом» (для браузера на старых системах): движок
  // вложен в страницу как base64 и не загружается отдельным файлом.
  const embedded = (globalThis as { __ENS_WASM_B64__?: string }).__ENS_WASM_B64__;
  if (embedded) {
    const bin = atob(embedded);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    await init({ module_or_path: bytes });
    return;
  }
  await init();
}

const TZ = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';

// Локаль: 'ru' — запятая в числах, «;» между аргументами, даты дд.мм.гггг;
// 'en' — точка. Язык функций — русский (=СУММ); английские имена (=SUM)
// движок тоже понимает и показывает по-русски, как русский Excel.
const LANGUAGE = 'ru';

export function newWorkbook(locale: string, name = 'Книга1'): Model {
  const model = new Engine(name, locale, TZ, LANGUAGE);
  model.renameSheet(0, 'Лист1');
  model.flushSendQueue(); // переименование — не правка пользователя
  return model;
}

/** Свободное имя для нового листа: Лист2, Лист3… */
export function nextSheetName(model: Model): string {
  const names = new Set(model.getWorksheetsProperties().map((p) => p.name.toLowerCase()));
  for (let n = 1; ; n++) if (!names.has(`лист${n}`)) return `Лист${n}`;
}

/** Книга из внутреннего формата движка (снимок для восстановления). */
export function fromBytes(bytes: Uint8Array): Model {
  return Engine.fromBytes(bytes, LANGUAGE);
}

export function openXlsx(bytes: Uint8Array, name: string, locale: string): Model {
  return Engine.fromXlsx(bytes, name, locale, TZ, LANGUAGE);
}

/** Номер столбца → буквы (1 → A, 27 → AA). */
export function columnName(column: number): string {
  let name = '';
  for (let n = column; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}

/** Буквы → номер столбца. */
export function columnNumber(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
