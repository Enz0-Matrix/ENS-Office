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

/** Часть текста ячейки со своим шрифтом; vert: 0 — обычный, 1 — надстрочный, 2 — подстрочный. */
export interface RichRun { text: string; b?: boolean; i?: boolean; u?: boolean; strike?: boolean; sz?: number | null; color?: string | null; name?: string | null; vert?: number }

/** Рисунок на листе: левый верхний угол — в ячейке (row, column) со сдвигом dx, dy (px). sig меняется вместе с содержимым. */
export interface ImageInfo { row: number; column: number; dx: number; dy: number; width: number; height: number; ext: string; name: string; sig: number }

/** Диаграмма: kind 0 — столбцы, 1 — линии, 2 — круговая; range — «A1:C5» (первая строка — имена рядов, первый столбец — подписи). */
export interface Chart { kind: number; title: string; range: string; row: number; column: number; dx: number; dy: number; width: number; height: number; colors: string[]; legend: number; labels: boolean;
  /** Подписи осей, границы шкалы значений (null — авто), без сетки, линии: без маркеров, сглаженные, толстые. */
  x_title: string; y_title: string; y_min: number | null; y_max: number | null; no_grid: boolean; no_markers: boolean; smooth: boolean; thick: boolean }

/** Фигура или надпись: kind 0 — прямоугольник, 1 — скруглённый, 2 — эллипс, 3 — линия, 4 — стрелка, 5 — надпись. Цвета «#RRGGBB», пусто — нет. */
export interface Shape {
  kind: number; text: string; fill: string; line_color: string; line_width: number; text_color: string; font_size: number;
  bold: boolean; align: number; flip: boolean; row: number; column: number; dx: number; dy: number; width: number; height: number;
}

/** Раскрывающийся список: ячейки (как в xlsx, «B2:B10 D4») и источник (`"Да,Нет"` или диапазон `Лист1!$A$1:$A$5`). */
export interface Validation { sqref: string; source: string; allow_blank: boolean; no_arrow: boolean; error: string }

export interface CellStyle {
  num_fmt: string;
  fill: { color?: Color };
  font: { b: boolean; i: boolean; u: boolean; strike: boolean; sz: number; color?: Color; name: string };
  alignment?: { horizontal?: string; vertical?: string; wrap_text?: boolean; text_rotation?: number };
  border?: CellBorder;
  /** Защита ячейки: нет поля — защищаемая, формулы видны. */
  protection?: { locked: boolean; hidden: boolean };
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
  setColumnsHidden(sheet: number, columnStart: number, columnEnd: number, hidden: boolean): void;
  /** Отбор фильтров листа (JSON), сохраняется в xlsx. */
  getFilterCriteria(sheet: number): string;
  setFilterCriteria(sheet: number, json: string): void;
  /** Комментарии (примечания) листа; ref — адрес ячейки A1. */
  getComments(sheet: number): { ref: string; text: string; author: string }[];
  setComment(sheet: number, row: number, column: number, text: string, author: string): void;
  deleteComment(sheet: number, row: number, column: number): void;
  /** Ручные разрывы страниц: [после строк, после столбцов] (номера с 1). */
  getPageBreaks(sheet: number): [number[], number[]];
  setPageBreak(sheet: number, isRow: boolean, index: number, on: boolean): void;
  clearPageBreaks(sheet: number): void;
  /** Части текста ячейки с собственным оформлением (rich text) или null. */
  getRichText(sheet: number, row: number, column: number): RichRun[] | null;
  setRichText(sheet: number, row: number, column: number, runs: RichRun[]): void;
  /** Ячейки с rich text в прямоугольнике: [строка, столбец, …]. */
  richTextCellsIn(sheet: number, r1: number, c1: number, r2: number, c2: number): number[];
  getImages(sheet: number): ImageInfo[];
  getImageData(sheet: number, index: number): Uint8Array;
  addImage(sheet: number, row: number, column: number, width: number, height: number, ext: string, name: string, data: Uint8Array): number;
  placeImage(sheet: number, index: number, row: number, column: number, dx: number, dy: number, width: number, height: number): void;
  deleteImage(sheet: number, index: number): void;
  getValidations(sheet: number): Validation[];
  /** Заменяет все списки листа (одно действие отмены). */
  setValidations(sheet: number, list: Validation[]): void;
  /** У ячейки есть список со стрелкой. */
  hasDropdown(sheet: number, row: number, column: number): boolean;
  getDropdownItems(sheet: number, row: number, column: number): string[];
  getShapes(sheet: number): Shape[];
  /** Заменяет все фигуры листа (одно действие отмены). */
  setShapes(sheet: number, shapes: Shape[]): void;
  getCharts(sheet: number): Chart[];
  /** Заменяет все диаграммы листа (одно действие отмены). */
  setCharts(sheet: number, charts: Chart[]): void;
  getRangeNumbers(sheet: number, r1: number, c1: number, r2: number, c2: number): number[];
  isSheetProtected(sheet: number): boolean;
  /** Защита снабжена паролем (задан в другой программе): снять её здесь нельзя. */
  isSheetPasswordProtected(sheet: number): boolean;
  setSheetProtected(sheet: number, on: boolean): void;
  /** Защита листа с паролем (пустой пароль — без пароля); salt — случайные байты. */
  /** allowed — что остаётся разрешённым: formatCells, formatColumns, formatRows, insertColumns, insertRows, deleteColumns, deleteRows, sort. */
  protectSheet(sheet: number, password: string, salt: Uint8Array, allowed: string[]): void;
  /** Снять защиту листа; с паролем — нужен верный пароль (иначе ошибка «Неверный пароль»). */
  unprotectSheet(sheet: number, password: string): void;
  isStructurePasswordProtected(): boolean;
  lockStructure(password: string, salt: Uint8Array): void;
  unlockStructure(password: string): void;
  isStructureLocked(): boolean;
  setStructureLocked(on: boolean): void;
  getPageSetup(sheet: number): string;
  getPrintArea(sheet: number): number[];
  /** Сквозные строки и столбцы печати: [строка1, строка2, столбец1, столбец2], 0 — нет. */
  getPrintTitles(sheet: number): number[];
  setPrintTitles(sheet: number, titles: number[]): void;
  setPrintArea(sheet: number, area: number[]): void;
  getEnsChanges(): string;
  clearEnsChanges(): void;
  sheetContentHash(sheet: number): string;
  getChangeLog(): string;
  setChangeLog(json: string): void;
  getLastModifiedBy(): string;
  setLastModifiedBy(name: string): void;
  setPageSetup(sheet: number, json: string): void;
  /** Имена (используются для второго и следующих фильтров листа). */
  getDefinedNameList(): { name: string; scope?: number | null; formula: string }[];
  newDefinedName(name: string, scope: number | null | undefined, formula: string): void;
  updateDefinedName(name: string, scope: number | null | undefined, newName: string, newScope: number | null | undefined, newFormula: string): void;
  deleteDefinedName(name: string, scope?: number | null): void;
  /** Группы (структура): [индекс, уровень, свёрнуто 0/1, …]. */
  getRowOutline(sheet: number): Int32Array;
  getColumnOutline(sheet: number): Int32Array;
  /** Где «+/−» групп: [под строками (1) или над, справа от столбцов (1) или слева]. */
  getOutlineSummary(sheet: number): Uint8Array;
  setOutlineSummary(sheet: number, below: boolean, right: boolean): void;
  setRowsOutline(sheet: number, r1: number, r2: number, level: number): void;
  setRowCollapsed(sheet: number, row: number, collapsed: boolean): void;
  setColumnsOutline(sheet: number, c1: number, c2: number, level: number): void;
  setColumnCollapsed(sheet: number, column: number, collapsed: boolean): void;
  /** Новые книги — по центру по вертикали. */
  setDefaultVerticalCenter(): void;
  setDefaultFont(name: string, size: number): void;
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
  /** Отмена по листам: mode 0 — только этот лист, 1 — все листы, кроме этого. Ответ: 0 — сделано, 1 — нечего, 2 — мешает структурное изменение. */
  undoScoped(sheet: number, mode: number): number;
  redoScoped(sheet: number, mode: number): number;
  canUndoScoped(sheet: number, mode: number): boolean;
  canRedoScoped(sheet: number, mode: number): boolean;
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

/** Шрифт новых книг (Параметры → Шрифт по умолчанию). */
export const newBookFont = { name: 'Times New Roman', size: 12 };

export function newWorkbook(locale: string, name = 'Книга1'): Model {
  const model = new Engine(name, locale, TZ, LANGUAGE);
  model.renameSheet(0, 'Лист1');
  // По желанию пользователя: в новых книгах — по центру по вертикали.
  try { model.setDefaultVerticalCenter(); } catch { /* старый движок */ }
  try { model.setDefaultFont(newBookFont.name, newBookFont.size); } catch { /* старый движок */ }
  model.flushSendQueue(); // переименование — не правка пользователя
  try { model.clearEnsChanges(); } catch { /* старый движок */ }
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
