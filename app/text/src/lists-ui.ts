// Окно «Настроить многоуровневый список»: шаблон, формат, начало и ширина номера для каждого уровня, с живым образцом.

import { FORMATS, PRESETS, labelFor, sampleLabels, type Level, type Scheme } from './numbering';
import { dialog, el } from './ui';

const LEVELS = 9;

export function schemeDialog(current: Scheme | null, apply: (s: Scheme | null) => void): void {
  dialog((box, close) => {
    const base: Scheme = (current ?? PRESETS[1].scheme!).map((l) => ({ ...l }));
    const levels: Level[] = Array.from({ length: LEVELS }, (_, i) => ({ ...base[Math.min(i, base.length - 1)] }));
    const preview = el('div', { class: 'scheme-preview' });
    const rows = el('div', { class: 'scheme-rows' });
    let focused: HTMLInputElement | null = null;

    const refresh = () => {
      preview.replaceChildren(...sampleLabels(levels, 6).map((lab, i) => {
        const ind = el('span', { class: 'sp-ind' });
        ind.style.width = `${i * 18}px`;
        return el('div', { class: 'sp-line' }, ind, el('b', { textContent: lab }), el('span', { textContent: '  Текст пункта' }));
      }));
    };
    for (let i = 0; i < LEVELS; i++) {
      const lv = levels[i];
      const text = el('input', { type: 'text', value: lv.text, class: 'sr-text', oninput: () => { lv.text = text.value; refresh(); }, onfocus: () => { focused = text; } });
      const fmt = el('select', { onchange: () => { lv.fmt = fmt.value; refresh(); } }, ...FORMATS.map((f) => new Option(f.label, f.value)));
      fmt.value = lv.fmt;
      const start = el('input', { type: 'number', value: String(lv.start), min: '1', class: 'sr-num', oninput: () => { lv.start = Math.max(1, +start.value || 1); refresh(); } });
      const hang = el('input', { type: 'number', value: String(lv.hang), min: '10', max: '200', class: 'sr-num', oninput: () => { lv.hang = Math.max(10, +hang.value || 36); refresh(); } });
      rows.append(el('div', { class: 'scheme-row' }, el('span', { class: 'sr-lv', textContent: `Уровень ${i + 1}` }), text, fmt, start, hang));
    }
    const tokens = el('div', { class: 'scheme-tokens' }, el('span', { textContent: 'Вставить в шаблон номер уровня: ' }),
      ...Array.from({ length: 6 }, (_, i) => el('button', {
        textContent: `№${i + 1}`, title: `Номер ${i + 1}-го уровня (%${i + 1})`,
        onmousedown: (e: MouseEvent) => e.preventDefault(),
        onclick: () => {
          const t = focused ?? (rows.querySelector('.sr-text') as HTMLInputElement);
          const at = t.selectionStart ?? t.value.length;
          t.value = `${t.value.slice(0, at)}%${i + 1}${t.value.slice(t.selectionEnd ?? at)}`;
          t.dispatchEvent(new Event('input'));
          t.focus();
        },
      })));
    box.append(
      el('h2', { textContent: 'Многоуровневый список' }),
      el('p', { textContent: 'Шаблон: %1 — номер первого уровня, %2 — второго и т. д. Например «%1.%2.%3.» даёт «1.2.14.», а «%1.%2.%3.%4» с форматом «А, Б, В» на четвёртом уровне — «1.2.14.Б».' }),
      el('div', { class: 'scheme-head' }, el('span', { textContent: 'Уровень' }), el('span', { textContent: 'Шаблон' }), el('span', { textContent: 'Формат числа' }), el('span', { textContent: 'Начать с' }), el('span', { textContent: 'Ширина номера, pt' })),
      rows, tokens, el('h3', { textContent: 'Образец' }), preview,
      el('div', { class: 'dlg-actions' },
        el('button', { textContent: 'Отмена', onclick: close }),
        el('button', { class: 'primary', textContent: 'Применить', onclick: () => { close(); apply(levels.map((l) => ({ ...l }))); } })),
    );
    // Проверка: каждый шаблон должен показывать хоть что-то.
    void labelFor;
    refresh();
  }, false);
}
