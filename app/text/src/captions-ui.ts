// SPDX-License-Identifier: AGPL-3.0-or-later
// Окно «Вставить название» (Ссылки → Название): подпись к рисунку, таблице или формуле с автоматическим номером.

import type { EditorView } from 'prosemirror-view';
import { CAPTION_LABELS, CAPTION_SEP, captionTarget, insertCaption, nextNumber } from './captions';
import { dialog, el } from './ui';

export function captionDialog(view: EditorView): void {
  dialog((box, close) => {
    box.classList.add('caption-dialog');
    const target = captionTarget(view.state);
    const initial = target.kind === 'table' ? 'Таблица' : 'Рисунок';
    const label = el('select', { ariaLabel: 'Подпись' }, ...CAPTION_LABELS.map((l) => new Option(l, l)), new Option('Своя…', ''));
    label.value = initial;
    const custom = el('input', { type: 'text', ariaLabel: 'Своя подпись', placeholder: 'Например: Схема', hidden: true, maxLength: 30 });
    const text = el('input', { type: 'text', ariaLabel: 'Текст подписи', placeholder: 'Название' });
    const pos = el('select', { ariaLabel: 'Положение' }, new Option('Над выделенным объектом', 'above'), new Option('Под выделенным объектом', 'below'));
    // ГОСТ 7.32: название таблицы — над таблицей, подпись рисунка — под рисунком.
    pos.value = target.kind === 'table' ? 'above' : 'below';
    const preview = el('p', { class: 'caption-preview' });
    const name = () => (label.value || custom.value.trim() || 'Рисунок').replace(/\s+/g, '_');
    const sync = () => {
      custom.hidden = label.value !== '';
      const t = text.value.trim();
      preview.textContent = `${name().replace(/_/g, ' ')} ${nextNumber(view.state, name(), pos.value === 'above')}${t ? CAPTION_SEP + t : ''}`;
    };
    for (const c of [label, custom, text, pos]) { c.addEventListener('input', sync); c.addEventListener('change', sync); }
    const ok = el('button', { class: 'primary', textContent: 'Вставить', onclick: () => {
      close();
      view.dispatch(insertCaption(view.state, { name: name(), text: text.value, above: pos.value === 'above' }));
      view.focus();
    } });
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter' && e.target instanceof HTMLInputElement) ok.click(); });
    box.append(el('h2', { textContent: 'Название' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Подпись' }), label, custom),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Текст' }), text),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Положение' }), pos),
      preview,
      el('p', { class: 'hint', textContent: 'Номера обновляются сами при вставке, удалении и перестановке подписей — как поле SEQ в Word, но без «Обновить поле».' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), ok));
    sync();
    text.focus();
  }, false);
}
