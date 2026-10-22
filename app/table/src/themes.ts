// Темы оформления: классическая (по умолчанию) и весёлые. Тема меняет только
// вид программы: оформление ячеек в книге и печать остаются как есть.
// Интерфейс красит CSS ([data-theme] в styles.css), сетку — GridTheme.

import { CLASSIC_GRID, type GridTheme } from './grid';

export interface Theme {
  id: string;
  name: string;
  /** Короткая подпись в меню. */
  note: string;
  grid: Partial<GridTheme>;
}

export const THEMES: Theme[] = [
  { id: 'classic', name: 'Классическая', note: 'нефритовая, спокойная', grid: {} },
  {
    id: 'wasteland',
    name: 'Пустошь',
    note: 'постапокалипсис: ржавчина, пыль, старая бумага',
    grid: {
      // Лист — пожелтевшая бумага из бункера, заголовки — потёртый металл.
      sheetBg: '#f2e8cf',
      gridLine: '#dccfa9',
      headerBg: '#34302a',
      headerLine: '#4d4637',
      headerText: '#c9b98f',
      cornerBg: '#24211c',
      cornerLine: '#4d4637',
      headerActiveFill: '#6b3a17',
      headerActiveText: '#ffcf70',
      activeBorder: '#c2571a',
      activeBorderInactive: '#9a6a48',
      selectionFill: 'rgba(194,87,26,0.16)',
    },
  },
];

export function themeById(id: string): Theme {
  return THEMES.find((t) => t.id === id) ?? THEMES[0];
}

/** Включить тему: атрибут для CSS и цвета сетки. */
export function applyTheme(id: string, grid: { theme: GridTheme; render: () => void }, motion: boolean): void {
  const t = themeById(id);
  const root = document.documentElement;
  if (t.id === 'classic') delete root.dataset.theme; else root.dataset.theme = t.id;
  root.toggleAttribute('data-theme-still', !motion);
  grid.theme = { ...CLASSIC_GRID, ...t.grid };
  grid.render();
}
