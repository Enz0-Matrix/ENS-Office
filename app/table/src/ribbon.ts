// Лента в стиле Office 2021–2024. Описана данными (ADR-003): элемент ленты
// ссылается на команду по command_id; поле module — для модульной установки.

import { command, keyLabel, run } from './commands';
import { ICONS } from './icons';
import { OFFICE_FONTS } from './fonts';
import { saveSettings, settings } from './settings';

/** Вызывается после перерисовки вкладки: кнопки с живым видом (например, «Границы») обновляют значок. */
let afterDraw: () => void = () => {};
export function onRibbonDrawn(f: () => void): void { afterDraw = f; f(); }

export function icon(name: string, size: 20 | 24 = 20): string {
  return ICONS[`${name}_${size}`] ?? ICONS[`${name}_${size === 20 ? 24 : 20}`] ?? '';
}

export type RibbonItem =
  | { type: 'large' | 'medium' | 'small'; command: string; icon: string; label?: string; module?: string }
  /** Кнопка с памятью: основная часть повторяет последний выбор, стрелка — меню (command + ':menu'). */
  | { type: 'split'; command: string; icon: string }
  /** Кнопка с палитрой: основная часть применяет последний цвет, стрелка открывает палитру. */
  | { type: 'color'; command: string; icon: string; kind: 'fill' | 'font' }
  /** Ползунок «Правила отмены»: что отменяют кнопки отмены и повтора (вся книга, один лист, остальные листы). */
  | { type: 'undoScope' }
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
  { value: '#,##0.00;[Red]-#,##0.00', label: 'Числовой' },
  { value: '#,##0.00 "₽";[Red]-#,##0.00 "₽"', label: 'Денежный' },
  { value: 'dd.mm.yyyy', label: 'Краткая дата' },
  { value: 'd mmmm yyyy', label: 'Длинная дата' },
  { value: 'hh:mm:ss', label: 'Время' },
  { value: '0.00%', label: 'Процентный' },
  { value: '0.00E+00', label: 'Экспоненциальный' },
  { value: '@', label: 'Текстовый' },
  /** Не формат, а окно «Формат ячеек». */
  { value: ':more', label: 'Другие форматы…' },
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

/** «Другие цвета…»: окно выбора любого цвета системы; результат — #RRGGBB. */
export function pickColor(initial: string, done: (color: string) => void): void {
  const input = document.createElement('input');
  input.type = 'color';
  input.value = /^#[0-9a-f]{6}$/i.test(initial) ? initial : '#2b7a6b';
  input.style.cssText = 'position:fixed;left:-100px;top:-100px;opacity:0';
  document.body.append(input);
  input.addEventListener('change', () => { done(input.value.toUpperCase()); input.remove(); });
  input.addEventListener('blur', () => setTimeout(() => input.remove(), 500));
  input.click();
}

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
  const custom = document.createElement('button');
  custom.className = 'palette-none';
  custom.textContent = 'Другие цвета…';
  custom.addEventListener('mousedown', (e) => {
    e.preventDefault();
    palette.hidden = true;
    pickColor(kind === 'fill' ? lastColor.fill : lastColor.font, pick);
  });
  palette.append(none, title('Цвета темы'), grid, title('Стандартные цвета'), std, custom);
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
        name: 'Отмена',
        items: [
          { type: 'medium', command: 'edit.undo', icon: 'arrow_undo' },
          { type: 'medium', command: 'edit.redo', icon: 'arrow_redo' },
          { type: 'undoScope' },
        ],
      },
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
          { type: 'small', command: 'format.strike', icon: 'text_strike' },
          { type: 'small', command: 'format.superscript', icon: 'text_sup' },
          { type: 'small', command: 'format.subscript', icon: 'text_sub' },
          { type: 'color', command: 'format.fill', icon: 'paint_bucket', kind: 'fill' },
          { type: 'color', command: 'format.fontColor', icon: 'text_color', kind: 'font' },
        ],
      },
      {
        name: 'Выравнивание',
        items: [
          { type: 'medium', command: 'align.top', icon: 'align_top' },
          { type: 'medium', command: 'align.middle', icon: 'align_middle' },
          { type: 'medium', command: 'align.bottom', icon: 'align_bottom' },
          { type: 'medium', command: 'align.left', icon: 'text_align_left' },
          { type: 'medium', command: 'align.center', icon: 'text_align_center' },
          { type: 'medium', command: 'align.right', icon: 'text_align_right' },
          { type: 'medium', command: 'align.justify', icon: 'text_align_justify' },
          { type: 'medium', command: 'format.wrap', icon: 'text_wrap' },
          { type: 'medium', command: 'format.rotate', icon: 'text_direction_rotate_90_left' },
          { type: 'medium', command: 'format.merge', icon: 'merge' },
          { type: 'split', command: 'format.borders', icon: 'border_all' },
        ],
      },
      {
        name: 'Число',
        items: [
          { type: 'select', command: 'format.number', options: NUMBER_FORMATS, width: 150 },
          { type: 'medium', command: 'format.currency', icon: 'num_currency' },
          { type: 'medium', command: 'format.percent', icon: 'num_percent' },
          { type: 'medium', command: 'format.grouping', icon: 'num_grouping' },
          { type: 'medium', command: 'format.decimalsMore', icon: 'num_dec_more' },
          { type: 'medium', command: 'format.decimalsLess', icon: 'num_dec_less' },
          { type: 'medium', command: 'format.cells', icon: 'num_format' },
        ],
      },
      {
        name: 'Стили',
        items: [{ type: 'large', command: 'format.conditional', icon: 'color', label: 'Условное форматирование' }],
      },
      {
        name: 'Редактирование',
        items: [
          { type: 'medium', command: 'edit.autosum', icon: 'autosum' },
          { type: 'medium', command: 'edit.insertFunction', icon: 'math_formula' },
          { type: 'medium', command: 'edit.clear', icon: 'eraser' },
          { type: 'medium', command: 'data.sortAsc', icon: 'sort_asc' },
          { type: 'medium', command: 'data.filter', icon: 'filter' },
          { type: 'medium', command: 'edit.find', icon: 'search' },
        ],
      },
    ],
  },
  {
    id: 'insert',
    label: 'Вставка',
    groups: [
      { name: 'Листы', items: [{ type: 'large', command: 'sheet.new', icon: 'table_add', label: 'Новый лист' }] },
      { name: 'Диаграммы', items: [
        { type: 'large', command: 'chart.bar', icon: 'chart_bar', label: 'Столбцы' },
        { type: 'large', command: 'chart.line', icon: 'chart_line', label: 'Линии' },
        { type: 'large', command: 'chart.pie', icon: 'chart_pie', label: 'Круговая' },
        { type: 'large', command: 'chart.edit', icon: 'chart_bar', label: 'Изменить' },
      ] },
      { name: 'Фигуры', items: [
        { type: 'large', command: 'shape.textbox', icon: 'text_box', label: 'Надпись' },
        { type: 'large', command: 'shape.rect', icon: 'shape_rect', label: 'Прямоугольник' },
        { type: 'large', command: 'shape.round', icon: 'shape_rect', label: 'Скруглённый' },
        { type: 'large', command: 'shape.ellipse', icon: 'shape_ellipse', label: 'Эллипс' },
        { type: 'large', command: 'shape.line', icon: 'shape_line', label: 'Линия' },
        { type: 'large', command: 'shape.arrow', icon: 'shape_arrow', label: 'Стрелка' },
      ] },
      { name: 'Ссылки', items: [{ type: 'large', command: 'link.edit', icon: 'link', label: 'Ссылка' }] },
      { name: 'Расположение фигуры', items: [
        { type: 'large', command: 'shape.front', icon: 'arrange_front', label: 'На передний план' },
        { type: 'large', command: 'shape.back', icon: 'arrange_back', label: 'На задний план' },
        { type: 'large', command: 'shape.rotateRight', icon: 'rotate_cw', label: 'Повернуть 90°' },
      ] },
      { name: 'Иллюстрации', items: [{ type: 'large', command: 'insert.image', icon: 'image', label: 'Рисунок' }] },
      { name: 'Примечания', items: [
        { type: 'large', command: 'comment.edit', icon: 'comment_add', label: 'Примечание' },
        { type: 'large', command: 'comment.delete', icon: 'comment_dismiss', label: 'Удалить примечание' },
      ] },
      { name: 'Разрывы страниц', items: [
        { type: 'large', command: 'page.breakInsert', icon: 'page_break', label: 'Вставить разрыв' },
        { type: 'large', command: 'page.breakRemove', icon: 'page_break', label: 'Убрать разрыв' },
        { type: 'large', command: 'page.breakReset', icon: 'page_break', label: 'Сбросить разрывы' },
      ] },
      { name: 'Символы', items: [{ type: 'large', command: 'insert.symbol', icon: 'symbols', label: 'Символ' }] },
    ],
  },
  {
    id: 'review',
    label: 'Рецензирование',
    groups: [
      { name: 'Защита', items: [
        { type: 'large', command: 'protect.sheet', icon: 'lock_closed', label: 'Защитить лист' },
        { type: 'large', command: 'protect.book', icon: 'lock_closed', label: 'Защитить структуру книги' },
        { type: 'large', command: 'protect.cells', icon: 'lock_closed', label: 'Защита ячеек' },
      ] },
      { name: 'Примечания', items: [
        { type: 'large', command: 'comment.edit', icon: 'comment_add', label: 'Примечание' },
        { type: 'large', command: 'comment.delete', icon: 'comment_dismiss', label: 'Удалить примечание' },
      ] },
    ],
  },
  {
    id: 'data',
    label: 'Данные',
    groups: [
      {
        name: 'Сортировка и фильтр',
        items: [
          { type: 'large', command: 'data.sortAsc', icon: 'sort_asc', label: 'От А до Я' },
          { type: 'large', command: 'data.sortDesc', icon: 'sort_desc', label: 'От Я до А' },
          { type: 'large', command: 'data.sortCustom', icon: 'arrow_sort', label: 'Сортировка…' },
          { type: 'large', command: 'data.filter', icon: 'filter', label: 'Фильтр' },
        ],
      },
      { name: 'Проверка данных', items: [{ type: 'large', command: 'data.validation', icon: 'dropdown_list', label: 'Проверка' }] },
      {
        name: 'Структура',
        items: [
          { type: 'large', command: 'data.group', icon: 'group', label: 'Сгруппировать' },
          { type: 'large', command: 'data.ungroup', icon: 'group_list', label: 'Разгруппировать' },
          { type: 'large', command: 'data.summary', icon: 'group', label: 'Положение «+»' },
        ],
      },
      {
        name: 'Поиск',
        items: [
          { type: 'large', command: 'edit.find', icon: 'search', label: 'Найти' },
          { type: 'large', command: 'edit.replace', icon: 'arrow_sort', label: 'Заменить' },
        ],
      },
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
          { type: 'large', command: 'fn.finance', icon: 'money', label: 'Финансовые' },
          { type: 'large', command: 'fn.logic', icon: 'branch_fork', label: 'Логические' },
          { type: 'large', command: 'fn.text', icon: 'text_font', label: 'Текстовые' },
          { type: 'large', command: 'fn.date', icon: 'calendar_ltr', label: 'Дата и время' },
          { type: 'large', command: 'fn.lookup', icon: 'search', label: 'Ссылки и массивы' },
          { type: 'large', command: 'fn.math', icon: 'calculator', label: 'Математические' },
          { type: 'large', command: 'fn.other', icon: 'more_horizontal', label: 'Другие функции' },
        ],
      },
      {
        name: 'Определённые имена',
        items: [{ type: 'large', command: 'formulas.names', icon: 'book_open', label: 'Диспетчер имён' }],
      },
      {
        name: 'Зависимости формул',
        items: [
          { type: 'large', command: 'nav.precedents', icon: 'arrow_previous', label: 'Влияющие ячейки' },
          { type: 'large', command: 'nav.back', icon: 'location', label: 'Вернуться к ячейке' },
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
        name: 'Оформление',
        items: [{ type: 'large', command: 'view.theme', icon: 'paint_brush', label: 'Тема' }],
      },
      {
        name: 'Скрытое',
        items: [{
          type: 'select', command: 'view.hidden', width: 190,
          options: [
            { value: 'ask', label: 'Скрытое: спрашивать' },
            { value: 'excel', label: 'Скрытое: как в Excel' },
            { value: 'all', label: 'Скрытое: всегда со всеми' },
            { value: 'visible', label: 'Скрытое: только видимые' },
          ],
        }],
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
/** tabs — свои вкладки (ENS-Text); по умолчанию вкладки ENS-Table. */
export function renderRibbon(tabsEl: HTMLElement, bodyEl: HTMLElement, onFile: () => void, tabs: RibbonTab[] = TABS): () => void {
  let active = 'home';
  const file = document.createElement('button');
  file.className = 'tab file';
  file.textContent = 'Файл';
  file.addEventListener('click', onFile);
  tabsEl.append(file);
  const tabButtons = new Map<string, HTMLButtonElement>();
  for (const tab of tabs) {
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
    const tab = tabs.find((t) => t.id === active)!;
    for (const group of tab.groups) {
      const g = document.createElement('div');
      g.className = 'group';
      const items = document.createElement('div');
      items.className = 'items';
      // Группа без мелких кнопок — в один ряд: над и под большими кнопками ничего нет.
      if (group.items.every((i) => i.type !== 'small' && i.type !== 'medium' && i.type !== 'color')) items.classList.add('row');
      for (const item of group.items) items.append(renderItem(item));
      const name = document.createElement('div');
      name.className = 'group-name';
      name.textContent = group.name;
      g.append(items, name);
      bodyEl.append(g);
    }
    refresh();
    afterDraw();
  }

  /** Три положения: вся книга / только этот лист / остальные листы; под ползунком — действующее правило. */
  const RULES: { key: 'book' | 'sheet' | 'others'; short: string; name: string; hint: string }[] = [
    { key: 'book', short: 'Книга', name: 'Вся книга', hint: 'Кнопки отмены и повтора идут по всем действиям в книге, на любых листах, как в Excel.' },
    { key: 'sheet', short: 'Лист', name: 'Только этот лист', hint: 'Отменяются и повторяются только действия на открытом листе; остальное не затрагивается.' },
    { key: 'others', short: 'Другие', name: 'Только другие листы, не этот', hint: 'Отменяются действия на других листах, а на открытом листе всё остаётся как есть.' },
  ];
  let scopeAbort: AbortController | null = null;
  function renderUndoScope(): HTMLElement {
    scopeAbort?.abort();
    scopeAbort = new AbortController();
    const wrap = document.createElement('div');
    wrap.className = 'undo-scope';
    wrap.innerHTML = `<div class="us-caption">Правила отмены</div>
      <div class="us-track" role="radiogroup" aria-label="Правила отмены"><span class="us-thumb"></span>${RULES.map((r) => `<button class="us-stop" role="radio" data-k="${r.key}" tabindex="-1">${r.short}</button>`).join('')}</div>
      <div class="us-rule"></div>`;
    const thumb = wrap.querySelector<HTMLElement>('.us-thumb')!;
    const rule = wrap.querySelector<HTMLElement>('.us-rule')!;
    const show = () => {
      const i = Math.max(0, RULES.findIndex((r) => r.key === settings.undoScope));
      thumb.style.transform = `translateY(${i * 100}%)`;
      wrap.querySelectorAll<HTMLElement>('.us-stop').forEach((b, k) => { b.classList.toggle('on', k === i); b.setAttribute('aria-checked', String(k === i)); });
      rule.textContent = RULES[i].name;
      wrap.title = RULES[i].hint;
    };
    wrap.querySelectorAll<HTMLButtonElement>('.us-stop').forEach((b) => {
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', () => {
        settings.undoScope = b.dataset.k as typeof settings.undoScope;
        saveSettings();
        document.dispatchEvent(new Event('ens-undo-scope-set')); // главное окно обновит стрелку у кнопок отмены
      });
    });
    document.addEventListener('ens-undo-scope', show, { signal: scopeAbort.signal });
    show();
    return wrap;
  }

  function renderItem(item: RibbonItem): HTMLElement {
    if (item.type === 'separator') {
      const s = document.createElement('div');
      s.className = 'sep';
      return s;
    }
    if (item.type === 'undoScope') return renderUndoScope();
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
    if (item.type === 'split') {
      const wrap = document.createElement('div');
      wrap.className = 'split medium';
      const main = document.createElement('button');
      main.className = 'btn medium';
      main.title = cmd?.label ?? '';
      main.innerHTML = icon(item.icon, 24);
      main.dataset.command = item.command;
      main.addEventListener('mousedown', (e) => e.preventDefault());
      main.addEventListener('click', () => run(item.command));
      const arrow = document.createElement('button');
      arrow.className = 'btn split-arrow';
      arrow.title = `${cmd?.label ?? ''}: выбрать`;
      arrow.textContent = '▾';
      // Меню — по click: на mousedown его тут же закрывает общий обработчик «щелчок мимо меню».
      arrow.addEventListener('mousedown', (e) => e.preventDefault());
      arrow.addEventListener('click', () => {
        const r = arrow.getBoundingClientRect();
        run(`${item.command}:menu`, `${Math.round(r.left)},${Math.round(r.bottom + 2)}`);
      });
      wrap.append(main, arrow);
      return wrap;
    }
    const b = document.createElement('button');
    b.className = `btn ${item.type}`;
    b.dataset.command = item.command;
    const keys = cmd?.keys?.length ? ` (${keyLabel(cmd.keys[0])})` : '';
    b.title = `${cmd?.label ?? item.command}${keys}`;
    b.innerHTML = icon(item.icon, item.type === 'small' ? 20 : 24) +
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
