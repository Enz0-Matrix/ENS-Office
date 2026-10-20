// Лента в стиле Office 2021–2024. Описана данными (ADR-003): элемент ленты
// ссылается на команду по command_id; поле module — для модульной установки.

import { command, keyLabel, run } from './commands';
import { ICONS } from './icons';
import { OFFICE_FONTS } from './fonts';

export function icon(name: string, size: 20 | 24 = 20): string {
  return ICONS[`${name}_${size}`] ?? ICONS[`${name}_${size === 20 ? 24 : 20}`] ?? '';
}

export type RibbonItem =
  | { type: 'large' | 'small'; command: string; icon: string; label?: string; module?: string }
  /** Кнопка с палитрой: основная часть применяет последний цвет, стрелка открывает палитру. */
  | { type: 'color'; command: string; icon: string; kind: 'fill' | 'font' }
  | { type: 'select'; command: string; options: { value: string; label: string }[]; width?: number }
  | { type: 'separator' };

export interface RibbonGroup {
  name: string;
  items: RibbonItem[];
}

export interface RibbonTab {
  id: string;
  label: string;
  groups: RibbonGroup[];
}

export const NUMBER_FORMATS = [
  { value: 'general', label: 'Общий' },
  { value: '#,##0.00', label: 'Числовой' },
  { value: '#,##0.00 "₽"', label: 'Денежный' },
  { value: 'dd.mm.yyyy', label: 'Краткая дата' },
  { value: '0.00%', label: 'Процентный' },
  { value: '@', label: 'Текстовый' },
];

export const FONTS = OFFICE_FONTS.map((f) => ({ value: f, label: f }));

export const FONT_SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
  .map((n) => ({ value: String(n), label: String(n) }));

/** Цвета темы Office (основной ряд) — оттенки строятся от них, как в палитре Excel. */
const THEME = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const STANDARD = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];

function mixColor(hex: string, toward: number, amount: number): string {
  const n = parseInt(hex.slice(1), 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((c) => Math.round(c + (toward - c) * amount));
  return '#' + ch.map((c) => c.toString(16).padStart(2, '0')).join('').toUpperCase();
}

/** Последний выбранный цвет: основная часть кнопки применяет его. */
export const lastColor = { fill: '#FFFF00', font: '#FF0000' };

const palette = document.createElement('div');
palette.className = 'palette';
palette.hidden = true;
document.body.append(palette);
document.addEventListener('mousedown', (e) => {
  const t = e.target as HTMLElement;
  if (!palette.hidden && !palette.contains(t) && !t.closest('.split-arrow')) palette.hidden = true;
});

function openPalette(anchor: HTMLElement, kind: 'fill' | 'font', pick: (color: string) => void): void {
  const swatch = (color: string) => {
    const b = document.createElement('button');
    b.className = 'swatch';
    b.style.background = color;
    b.title = color;
    b.addEventListener('mousedown', (e) => { e.preventDefault(); palette.hidden = true; pick(color); });
    return b;
  };
  palette.replaceChildren();
  const none = document.createElement('button');
  none.className = 'palette-none';
  none.textContent = kind === 'fill' ? 'Нет заливки' : 'Авто';
  none.addEventListener('mousedown', (e) => { e.preventDefault(); palette.hidden = true; pick(''); });
  const title = (t: string) => { const d = document.createElement('div'); d.className = 'palette-title'; d.textContent = t; return d; };
  const grid = document.createElement('div');
  grid.className = 'palette-grid';
  for (const c of THEME) grid.append(swatch(c));
  for (const amount of [0.8, 0.6, 0.4]) for (const c of THEME) grid.append(swatch(mixColor(c, 255, amount)));
  for (const amount of [0.25, 0.5]) for (const c of THEME) grid.append(swatch(mixColor(c, 0, amount)));
  const std = document.createElement('div');
  std.className = 'palette-grid';
  for (const c of STANDARD) std.append(swatch(c));
  palette.append(none, title('Цвета темы'), grid, title('Стандартные цвета'), std);
  const r = anchor.getBoundingClientRect();
  palette.style.left = `${r.left}px`;
  palette.style.top = `${r.bottom + 2}px`;
  palette.hidden = false;
}

export const TABS: RibbonTab[] = [
  {
    id: 'home',
    label: 'Главная',
    groups: [
      {
        name: 'Буфер обмена',
        items: [
          { type: 'large', command: 'clipboard.paste', icon: 'clipboard_paste', label: 'Вставить' },
          { type: 'small', command: 'clipboard.cut', icon: 'cut' },
          { type: 'small', command: 'clipboard.copy', icon: 'copy' },
        ],
      },
      {
        name: 'Шрифт',
        items: [
          { type: 'select', command: 'format.fontName', options: FONTS, width: 140 },
          { type: 'select', command: 'format.fontSize', options: FONT_SIZES, width: 56 },
          { type: 'small', command: 'format.bold', icon: 'text_bold' },
          { type: 'small', command: 'format.italic', icon: 'text_italic' },
          { type: 'small', command: 'format.underline', icon: 'text_underline' },
          { type: 'color', command: 'format.fill', icon: 'paint_bucket', kind: 'fill' },
          { type: 'color', command: 'format.fontColor', icon: 'text_color', kind: 'font' },
        ],
      },
      {
        name: 'Выравнивание',
        items: [
          { type: 'small', command: 'align.left', icon: 'text_align_left' },
          { type: 'small', command: 'align.center', icon: 'text_align_center' },
          { type: 'small', command: 'align.right', icon: 'text_align_right' },
          { type: 'small', command: 'align.top', icon: 'align_top' },
          { type: 'small', command: 'align.middle', icon: 'align_middle' },
          { type: 'small', command: 'align.bottom', icon: 'align_bottom' },
          { type: 'small', command: 'format.wrap', icon: 'text_wrap' },
          { type: 'small', command: 'format.merge', icon: 'merge' },
          { type: 'small', command: 'format.borders', icon: 'border_all' },
        ],
      },
      {
        name: 'Число',
        items: [{ type: 'select', command: 'format.number', options: NUMBER_FORMATS, width: 130 }],
      },
      {
        name: 'Редактирование',
        items: [
          { type: 'small', command: 'edit.autosum', icon: 'autosum' },
          { type: 'small', command: 'edit.insertFunction', icon: 'math_formula' },
          { type: 'small', command: 'edit.clear', icon: 'eraser' },
        ],
      },
    ],
  },
  {
    id: 'insert',
    label: 'Вставка',
    groups: [
      { name: 'Листы', items: [{ type: 'large', command: 'sheet.new', icon: 'table_add', label: 'Новый лист' }] },
    ],
  },
  {
    id: 'formulas',
    label: 'Формулы',
    groups: [
      {
        name: 'Библиотека функций',
        items: [
          { type: 'large', command: 'edit.insertFunction', icon: 'math_formula', label: 'Вставить функцию' },
          { type: 'large', command: 'edit.autosum', icon: 'autosum', label: 'Автосумма' },
        ],
      },
      {
        name: 'Зависимости формул',
        items: [
          { type: 'large', command: 'nav.precedents', icon: 'arrow_previous', label: 'Влияющие ячейки' },
          { type: 'large', command: 'nav.back', icon: 'arrow_hook_up_left', label: 'Вернуться к ячейке' },
        ],
      },
    ],
  },
  {
    id: 'view',
    label: 'Вид',
    groups: [
      {
        name: 'Показ',
        items: [
          { type: 'large', command: 'view.focusCell', icon: 'target', label: 'Фокус ячейки' },
          { type: 'large', command: 'view.gridlines', icon: 'grid', label: 'Сетка' },
        ],
      },
      {
        name: 'Окно',
        items: [{ type: 'large', command: 'view.freeze', icon: 'table_freeze', label: 'Закрепить области' }],
      },
      {
        name: 'Ввод',
        items: [
          { type: 'large', command: 'options.recognizeDates', icon: 'calendar_ltr', label: 'Распознавать даты' },
          { type: 'large', command: 'options.expandEditor', icon: 'arrow_autofit_width', label: 'Расширять поле ввода' },
          {
            type: 'select', command: 'options.decimal', width: 150,
            options: [{ value: 'comma', label: 'Дробная часть: 0,5' }, { value: 'point', label: 'Дробная часть: 0.5' }],
          },
        ],
      },
    ],
  },
];

/** Рисует ленту и строку вкладок. Возвращает функцию обновления состояний кнопок. */
export function renderRibbon(tabsEl: HTMLElement, bodyEl: HTMLElement, onFile: () => void): () => void {
  let active = 'home';
  const file = document.createElement('button');
  file.className = 'tab file';
  file.textContent = 'Файл';
  file.addEventListener('click', onFile);
  tabsEl.append(file);
  const tabButtons = new Map<string, HTMLButtonElement>();
  for (const tab of TABS) {
    const b = document.createElement('button');
    b.className = 'tab';
    b.textContent = tab.label;
    b.addEventListener('click', () => { active = tab.id; draw(); });
    tabsEl.append(b);
    tabButtons.set(tab.id, b);
  }

  const toggles: { el: HTMLElement; id: string }[] = [];
  function draw(): void {
    for (const [id, b] of tabButtons) b.classList.toggle('active', id === active);
    bodyEl.replaceChildren();
    toggles.length = 0;
    const tab = TABS.find((t) => t.id === active)!;
    for (const group of tab.groups) {
      const g = document.createElement('div');
      g.className = 'group';
      const items = document.createElement('div');
      items.className = 'items';
      for (const item of group.items) items.append(renderItem(item));
      const name = document.createElement('div');
      name.className = 'group-name';
      name.textContent = group.name;
      g.append(items, name);
      bodyEl.append(g);
    }
    refresh();
  }

  function renderItem(item: RibbonItem): HTMLElement {
    if (item.type === 'separator') {
      const s = document.createElement('div');
      s.className = 'sep';
      return s;
    }
    const cmd = command(item.command);
    if (item.type === 'color') {
      const wrap = document.createElement('div');
      wrap.className = 'split';
      const main = document.createElement('button');
      main.className = 'btn small color-main';
      main.title = cmd?.label ?? '';
      main.innerHTML = `${icon(item.icon, 20)}<i class="color-bar"></i>`;
      const bar = main.querySelector<HTMLElement>('.color-bar')!;
      bar.style.background = lastColor[item.kind];
      main.addEventListener('mousedown', (e) => e.preventDefault());
      main.addEventListener('click', () => run(item.command, lastColor[item.kind]));
      const arrow = document.createElement('button');
      arrow.className = 'btn split-arrow';
      arrow.title = `${cmd?.label ?? ''}: выбрать цвет`;
      arrow.textContent = '▾';
      arrow.addEventListener('mousedown', (e) => {
        e.preventDefault();
        openPalette(wrap, item.kind, (color) => {
          if (color) { lastColor[item.kind] = color; bar.style.background = color; }
          run(item.command, color);
        });
      });
      wrap.append(main, arrow);
      return wrap;
    }
    if (item.type === 'select') {
      const sel = document.createElement('select');
      sel.className = 'ribbon-select';
      sel.style.width = `${item.width ?? 120}px`;
      sel.title = cmd?.label ?? '';
      for (const o of item.options) sel.add(new Option(o.label, o.value));
      sel.addEventListener('change', () => run(item.command, sel.value));
      sel.dataset.command = item.command;
      return sel;
    }
    const b = document.createElement('button');
    b.className = item.type === 'large' ? 'btn large' : 'btn small';
    const keys = cmd?.keys?.length ? ` (${keyLabel(cmd.keys[0])})` : '';
    b.title = `${cmd?.label ?? item.command}${keys}`;
    b.innerHTML = icon(item.icon, item.type === 'large' ? 24 : 20) +
      (item.type === 'large' ? `<span>${item.label ?? cmd?.label ?? ''}</span>` : '');
    b.addEventListener('mousedown', (e) => e.preventDefault()); // не отнимать фокус у сетки
    b.addEventListener('click', () => run(item.command));
    if (cmd?.checked) toggles.push({ el: b, id: item.command });
    return b;
  }

  function refresh(): void {
    for (const t of toggles) t.el.classList.toggle('on', !!command(t.id)?.checked?.());
  }

  draw();
  return refresh;
}
