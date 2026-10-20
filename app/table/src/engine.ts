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
  deleteRows(sheet: number, row: number, count: number): void;
  insertColumns(sheet: number, column: number, count: number): void;
  deleteColumns(sheet: number, column: number, count: number): void;
  getCellStyle(sheet: number, row: number, column: number): { style: CellStyle };
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
