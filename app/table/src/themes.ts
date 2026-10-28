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
  /** Тёмная или светлая основа (общие правила в themes2.css); у первых тем — свои правила. */
  tone?: 'light' | 'dark';
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
  {
    id: 'clinic', name: 'Детская больница', note: 'мятный и небесный, крестики и пластыри', tone: 'light',
    grid: {
      sheetBg: '#ffffff', gridLine: '#dcefed', headerBg: '#eaf7f5', headerLine: '#c8e6e3', headerText: '#4b7a76',
      cornerBg: '#d5efec', cornerLine: '#c8e6e3', headerActiveFill: '#ffd6e0', headerActiveText: '#c2185b',
      activeBorder: '#2a9d97', activeBorderInactive: '#8fd3cd', selectionFill: 'rgba(92,195,224,0.15)',
    },
  },
  {
    id: 'transport', name: 'Транспорт', note: 'асфальт с разметкой, указатели, светофор', tone: 'light',
    grid: {
      sheetBg: '#ffffff', gridLine: '#e1e6eb', headerBg: '#3a3f44', headerLine: '#555c63', headerText: '#e8e8e8',
      cornerBg: '#2b2f33', cornerLine: '#555c63', headerActiveFill: '#f2b705', headerActiveText: '#1d2733',
      activeBorder: '#0b5cad', activeBorderInactive: '#7fa6cf', selectionFill: 'rgba(11,92,173,0.12)',
    },
  },
  {
    id: 'shell', name: 'Панцири', note: 'зелёные панцири, четыре цветные повязки', tone: 'dark',
    grid: {
      sheetBg: '#f4f8ee', gridLine: '#dbe8cf', headerBg: '#24302a', headerLine: '#3a4a3e', headerText: '#a3d18a',
      cornerBg: '#1c2620', cornerLine: '#3a4a3e', headerActiveFill: '#f28c28', headerActiveText: '#1c2620',
      activeBorder: '#3f8f2f', activeBorderInactive: '#8bb57c', selectionFill: 'rgba(63,143,47,0.15)',
    },
  },
  {
    id: 'teacher', name: 'Учительская', note: 'доска, тетрадь в клетку, красная ручка', tone: 'light',
    grid: {
      sheetBg: '#fffdf6', gridLine: '#d6e2f0', headerBg: '#2f4a3a', headerLine: '#466a55', headerText: '#f1efe3',
      cornerBg: '#243a2d', cornerLine: '#466a55', headerActiveFill: '#c62828', headerActiveText: '#ffffff',
      activeBorder: '#c62828', activeBorderInactive: '#d98b8b', selectionFill: 'rgba(198,40,40,0.10)',
    },
  },
  {
    id: 'armor', name: 'Броня', note: 'красный металл, золотая кромка, свет реактора', tone: 'dark',
    grid: {
      sheetBg: '#fbf6ee', gridLine: '#eadfcf', headerBg: '#1d1012', headerLine: '#4a2a23', headerText: '#e0b23a',
      cornerBg: '#12090b', cornerLine: '#4a2a23', headerActiveFill: '#0f4a57', headerActiveText: '#6fe8ff',
      activeBorder: '#b3121f', activeBorderInactive: '#c9787e', selectionFill: 'rgba(111,232,255,0.16)',
    },
  },
  {
    id: 'bubbles', name: 'Круглые друзья', note: 'пастельные шарики, всё круглое и мягкое', tone: 'light',
    grid: {
      sheetBg: '#ffffff', gridLine: '#f6e3ee', headerBg: '#fff0f7', headerLine: '#f3d6e6', headerText: '#8a6f9e',
      cornerBg: '#ffe3f0', cornerLine: '#f3d6e6', headerActiveFill: '#ffd36e', headerActiveText: '#6b4a00',
      activeBorder: '#ff7aa8', activeBorderInactive: '#f7b7d2', selectionFill: 'rgba(143,211,255,0.20)',
    },
  },
  {
    id: 'dark', name: 'Тёмная', note: 'спокойная, для работы', tone: 'dark',
    grid: {
      sheetBg: '#1f2023', gridLine: '#34363b', headerBg: '#2b2d31', headerLine: '#3c3f45', headerText: '#b8bcc4',
      cornerBg: '#232427', cornerLine: '#3c3f45', headerActiveFill: '#33507a', headerActiveText: '#eaf2ff',
      activeBorder: '#5a9be0', activeBorderInactive: '#5b6270', selectionFill: 'rgba(90,155,224,0.20)', textColor: '#e6e8ec',
    },
  },
  {
    id: 'finance', name: 'Финансовая', note: 'строгая, синяя с золотой линией', tone: 'light',
    grid: {
      sheetBg: '#ffffff', gridLine: '#d9dee6', headerBg: '#e8edf3', headerLine: '#c3ccd8', headerText: '#2d3e56',
      cornerBg: '#d8dfe9', cornerLine: '#c3ccd8', headerActiveFill: '#c9d8ee', headerActiveText: '#14305c',
      activeBorder: '#1f3a68', activeBorderInactive: '#8da2c4', selectionFill: 'rgba(31,58,104,0.12)',
    },
  },
  {
    id: 'blueprint', name: 'Инженерная', note: 'синяя «синька», белые линии', tone: 'dark',
    grid: {
      sheetBg: '#14407a', gridLine: '#2c5b97', headerBg: '#0f3566', headerLine: '#2c5b97', headerText: '#cfe3ff',
      cornerBg: '#0c2b54', cornerLine: '#2c5b97', headerActiveFill: '#2f6fbf', headerActiveText: '#ffffff',
      activeBorder: '#ffffff', activeBorderInactive: '#9fc2ee', selectionFill: 'rgba(255,255,255,0.14)', textColor: '#f2f7ff',
    },
  },
];

// --- Свои темы из файла (без переустановки программы) ---
// Файл темы — JSON: имя, светлая/тёмная основа, цвета интерфейса (CSS-переменные),
// цвета сетки и, по желанию, свои CSS-правила. Хранится в настройках программы.
// Ссылки на интернет (url(http…), @import) вырезаются: программа работает без сети.

export interface CustomTheme {
  id: string;
  name: string;
  note: string;
  tone?: 'light' | 'dark';
  vars: Record<string, string>;
  grid: Record<string, string>;
  css?: string;
}

let custom: CustomTheme[] = [];

/** Без сети: никаких внешних ссылок и импортов в CSS темы. */
function offlineCss(css: string): string {
  return css
    .replace(/@import[^;]*;?/gi, '')
    .replace(/url\(\s*(['"]?)\s*(https?:|\/\/|ftp:)[^)]*\)/gi, 'none')
    .replace(/<\/?style[^>]*>/gi, '');
}

const cssValue = (v: string) => offlineCss(String(v)).replace(/[{};]/g, '');

/** Свои темы — в стиль страницы. */
export function setCustomThemes(list: CustomTheme[]): void {
  custom = list;
  let el = document.getElementById('ens-user-themes') as HTMLStyleElement | null;
  if (!el) {
    el = document.createElement('style');
    el.id = 'ens-user-themes';
    document.head.append(el);
  }
  el.textContent = list.map((t) => {
    const vars = Object.entries(t.vars).filter(([k]) => /^--[\w-]+$/.test(k)).map(([k, v]) => `${k}: ${cssValue(v)};`).join(' ');
    return `:root[data-theme="${t.id}"] { ${vars} }\n${offlineCss(t.css ?? '')}`;
  }).join('\n');
}

/** Какие встроенные темы скрыты из меню (настройка «Скрыть тему»); задаётся из main. */
let hiddenIds: string[] = [];
export const setHiddenThemes = (ids: string[]) => { hiddenIds = ids; };
const hiddenBuiltin = () => hiddenIds;

/** Все темы для меню: встроенные и свои. */
export function allThemes(): Theme[] {
  return [
    ...THEMES.filter((x) => !hiddenBuiltin().includes(x.id)),
    ...custom.map((t) => ({ id: t.id, name: t.name, note: t.note || 'своя тема', tone: t.tone, grid: t.grid as Partial<GridTheme> })),
  ];
}

/** Файл темы → тема (или текст ошибки). */
export function parseThemeFile(text: string): CustomTheme | string {
  let j: Record<string, unknown>;
  try { j = JSON.parse(text.replace(/^﻿/, '')); } catch { return 'Файл темы — не JSON. Возьмите за основу «Сохранить образец темы».'; }
  const name = typeof j.name === 'string' ? j.name.trim().slice(0, 40) : '';
  if (!name) return 'В файле темы нет имени («name»).';
  const obj = (v: unknown) => (v && typeof v === 'object' ? Object.fromEntries(Object.entries(v as object).map(([k, x]) => [k, String(x)])) : {});
  const slug = name.toLowerCase().replace(/[^\p{L}\p{N}]+/gu, '-').replace(/^-|-$/g, '') || 'theme';
  return {
    id: `user-${slug}`,
    name,
    note: typeof j.note === 'string' ? j.note.slice(0, 80) : '',
    tone: j.tone === 'dark' ? 'dark' : j.tone === 'light' ? 'light' : undefined,
    vars: obj(j.vars),
    grid: obj(j.grid),
    css: typeof j.css === 'string' ? j.css.slice(0, 50_000) : undefined,
  };
}

/** Образец файла темы — от текущей темы: цвета, которые можно менять. */
export function themeTemplate(id: string): string {
  const t = themeById(id);
  const cs = getComputedStyle(document.documentElement);
  const names = ['--accent', '--accent-dark', '--bg', '--panel', '--border', '--text', '--muted', '--hover', '--on-bg', '--status-bg', '--brand-bg'];
  const vars = Object.fromEntries(names.map((n) => [n, cs.getPropertyValue(n).trim()]).filter(([, v]) => v));
  return JSON.stringify({
    name: `Моя тема (на основе «${t.name}»)`,
    note: 'своя тема',
    tone: t.tone ?? 'light',
    vars,
    grid: { ...CLASSIC_GRID, ...t.grid },
    css: '/* Свои правила, например: [data-theme] .app-mark { color: gold; } */',
  }, null, 2);
}

export function themeById(id: string): Theme {
  return allThemes().find((t) => t.id === id) ?? THEMES[0];
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
  root.toggleAttribute('data-theme-light', t.tone === 'light');
  root.toggleAttribute('data-theme-dark', t.tone === 'dark');
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
