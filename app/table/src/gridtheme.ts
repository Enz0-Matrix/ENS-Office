// Цвета сетки (темы оформления). Вынесено из grid.ts, чтобы темы можно было подключать без движка таблиц (ENS-Tekst).

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
