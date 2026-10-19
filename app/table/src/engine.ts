// Обёртка над движком ENS-Table (форк IronCalc, собранный в WebAssembly).

export interface Area {
  sheet: number;
  row: number;
  column: number;
  width: number;
  height: number;
}

export type Color = string | [number, number] | undefined;

export interface CellStyle {
  num_fmt: string;
  fill: { color?: Color };
  font: { b: boolean; i: boolean; u: boolean; strike: boolean; sz: number; color?: Color; name: string };
  alignment?: { horizontal?: string; vertical?: string; wrap_text?: boolean };
}

export interface SheetProps {
  name: string;
  state: string;
  sheet_id: number;
}

/** Методы движка, которыми пользуется интерфейс. */
export interface Model {
  setUserInput(sheet: number, row: number, column: number, input: string): void;
  getFormattedCellValue(sheet: number, row: number, column: number): string;
  getCellContent(sheet: number, row: number, column: number): string;
  getCellType(sheet: number, row: number, column: number): number;
  getColumnWidth(sheet: number, column: number): number;
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
  getWorksheetsProperties(): SheetProps[];
  getShowGridLines(sheet: number): boolean;
  setShowGridLines(sheet: number, show: boolean): void;
  resolveColor(color: Color): string;
  toXlsx(): Uint8Array;
  free(): void;
}

// Собирается scripts/build-engine.sh (в git не хранится).
import init, { Model as EngineModel } from './engine-pkg/engine.js';

interface EngineStatic {
  new (name: string, locale: string, timezone: string, language: string): Model;
  fromXlsx(bytes: Uint8Array, name: string, locale: string, timezone: string, language: string): Model;
}

const Engine = EngineModel as unknown as EngineStatic;

export async function loadEngine(): Promise<void> {
  await init();
}

const TZ = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';

// Локаль: 'ru' — запятая в числах, «;» между аргументами, даты дд.мм.гггг;
// 'en' — точка. Имена функций пока английские (задание 6 для помощников).
export function newWorkbook(locale: string, name = 'Книга1'): Model {
  return new Engine(name, locale, TZ, 'en');
}

export function openXlsx(bytes: Uint8Array, name: string, locale: string): Model {
  return Engine.fromXlsx(bytes, name, locale, TZ, 'en');
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
