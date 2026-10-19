// Лента в стиле Office 2021–2024. Описана данными (ADR-003): элемент ленты
// ссылается на команду по command_id; поле module — для модульной установки.

import { command, keyLabel, run } from './commands';
import { ICONS } from './icons';

export function icon(name: string, size: 20 | 24 = 20): string {
  return ICONS[`${name}_${size}`] ?? ICONS[`${name}_${size === 20 ? 24 : 20}`] ?? '';
}

export type RibbonItem =
  | { type: 'large' | 'small'; command: string; icon: string; label?: string; module?: string }
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
          { type: 'small', command: 'format.bold', icon: 'text_bold' },
          { type: 'small', command: 'format.italic', icon: 'text_italic' },
          { type: 'small', command: 'format.underline', icon: 'text_underline' },
          { type: 'small', command: 'format.fill', icon: 'paint_bucket' },
          { type: 'small', command: 'format.fontColor', icon: 'text_color' },
        ],
      },
      {
        name: 'Выравнивание',
        items: [
          { type: 'small', command: 'align.left', icon: 'text_align_left' },
          { type: 'small', command: 'align.center', icon: 'text_align_center' },
          { type: 'small', command: 'align.right', icon: 'text_align_right' },
        ],
      },
      {
        name: 'Число',
        items: [{ type: 'select', command: 'format.number', options: NUMBER_FORMATS, width: 130 }],
      },
      {
        name: 'Редактирование',
        items: [
          { type: 'small', command: 'edit.autosum', icon: 'math_formula' },
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
        name: 'Зависимости формул',
        items: [
          { type: 'large', command: 'nav.precedents', icon: 'arrow_previous', label: 'Влияющие ячейки' },
          { type: 'large', command: 'nav.back', icon: 'arrow_undo', label: 'Назад' },
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
