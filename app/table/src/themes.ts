// Темы оформления: классическая (по умолчанию) и весёлые. Тема меняет только
// вид программы: оформление ячеек в книге и печать остаются как есть.
// Интерфейс красит CSS ([data-theme] в themes.css), сетку — GridTheme.

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
  {
    id: 'neon',
    name: 'Диско-неон',
    note: 'ночной клуб: розовый и голубой неон, переливы',
    grid: {
      // Лист светлый (чтобы читался чёрный текст из файлов), чуть сиреневый.
      sheetBg: '#fbf7ff',
      gridLine: '#eadcff',
      headerBg: '#1a1033',
      headerLine: '#3a2466',
      headerText: '#c9a8ff',
      cornerBg: '#120a26',
      cornerLine: '#3a2466',
      headerActiveFill: '#ff2bd6',
      headerActiveText: '#ffffff',
      activeBorder: '#ff2bd6',
      activeBorderInactive: '#b07ad6',
      selectionFill: 'rgba(0,240,255,0.16)',
    },
  },
  {
    id: 'matrix',
    name: 'Матрица',
    note: 'зелёный код на чёрном, цифровой дождь',
    grid: {
      sheetBg: '#f2fff4',
      gridLine: '#cdeed6',
      headerBg: '#000000',
      headerLine: '#0b3d17',
      headerText: '#00c03a',
      cornerBg: '#000000',
      cornerLine: '#0b3d17',
      headerActiveFill: '#003b12',
      headerActiveText: '#7dff9b',
      activeBorder: '#00b336',
      activeBorderInactive: '#5f9f6f',
      selectionFill: 'rgba(0,200,60,0.14)',
    },
  },
];

export function themeById(id: string): Theme {
  return THEMES.find((t) => t.id === id) ?? THEMES[0];
}

/** Включить тему: атрибут для CSS и цвета сетки. */
export const systemReducesMotion = () => matchMedia('(prefers-reduced-motion: reduce)').matches;

/** motion — анимация темы включена; forced — включена вопреки «уменьшить движение» в системе. */
export function applyTheme(id: string, grid: { theme: GridTheme; render: () => void }, motion: boolean, forced = false): void {
  const t = themeById(id);
  const root = document.documentElement;
  if (t.id === 'classic') delete root.dataset.theme; else root.dataset.theme = t.id;
  root.toggleAttribute('data-theme-still', !motion);
  root.toggleAttribute('data-theme-forced', motion && forced);
  grid.theme = { ...CLASSIC_GRID, ...t.grid };
  grid.render();
  digitalRain(t.id === 'matrix' && motion && (forced || !systemReducesMotion()));
}

// --- «Матрица»: цифровой дождь в заголовке окна (лёгкий: ~15 кадров в секунду) ---

let rain: { canvas: HTMLCanvasElement; timer: number } | null = null;

function digitalRain(on: boolean): void {
  if (!on) {
    if (rain) { clearInterval(rain.timer); rain.canvas.remove(); rain = null; }
    return;
  }
  if (rain) return;
  const bar = document.querySelector<HTMLElement>('.titlebar');
  if (!bar) return;
  const canvas = document.createElement('canvas');
  canvas.className = 'matrix-rain';
  canvas.setAttribute('aria-hidden', 'true');
  bar.prepend(canvas);
  const ctx = canvas.getContext('2d')!;
  const glyphs = 'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝ0123456789ЕНС=+-*';
  const step = 10;
  let drops: number[] = [];
  const resize = () => {
    canvas.width = bar.clientWidth;
    canvas.height = bar.clientHeight;
    drops = Array.from({ length: Math.ceil(canvas.width / step) }, () => Math.random() * -20);
  };
  resize();
  const timer = window.setInterval(() => {
    if (canvas.width !== bar.clientWidth) resize();
    ctx.fillStyle = 'rgba(0,0,0,0.18)';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.font = `${step}px monospace`;
    drops.forEach((y, i) => {
      const ch = glyphs[(Math.random() * glyphs.length) | 0];
      ctx.fillStyle = Math.random() < 0.06 ? '#d8ffe0' : '#00e04a';
      ctx.fillText(ch, i * step, y * step);
      drops[i] = y * step > canvas.height && Math.random() > 0.9 ? 0 : y + 1;
    });
  }, 66);
  rain = { canvas, timer };
}
