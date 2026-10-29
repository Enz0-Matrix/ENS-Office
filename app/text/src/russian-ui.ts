// Окна «по-русски»: сумма прописью, ФИО и должность по падежам, дата.

import type { EditorView } from 'prosemirror-view';
import { CASES, CASE_LABELS, CURRENCIES, declineFullName, declinePosition, sumWords, type Case, type NameGender } from './russian';
import { dialog, el, toast } from './ui';

type Host = () => EditorView;

function insertText(host: Host, text: string): void {
  const v = host();
  v.dispatch(v.state.tr.insertText(text).scrollIntoView());
  v.focus();
}

/** Текст выделения (одного абзаца) — как подсказка в поле ввода. */
function selected(host: Host): string {
  const { state } = host();
  const s = state.selection;
  return !s.empty && s.$from.sameParent(s.$to) ? state.doc.textBetween(s.from, s.to) : '';
}

export function sumDialog(host: Host): void {
  dialog((box, close) => {
    const amount = el('input', { type: 'text', value: selected(host).replace(/[^\d\s,. -]/g, '').trim(), placeholder: '12 345,67' });
    const cur = el('select', {}, ...Object.entries(CURRENCIES).map(([k, c]) => el('option', { value: k, textContent: `${k} (${c.sign})` })));
    const style = el('select', {}, el('option', { value: 'words', textContent: 'Только прописью' }), el('option', { value: 'both', textContent: '12 345 (прописью) рублей 67 копеек' }));
    const kw = el('input', { type: 'checkbox' });
    const out = el('div', { class: 'ru-out' });
    let text = '';
    const calc = () => {
      try { text = sumWords(amount.value, { currency: cur.value, style: style.value as 'words' | 'both', kopWords: kw.checked }); out.textContent = text; out.classList.remove('bad'); }
      catch (e) { text = ''; out.textContent = amount.value.trim() ? (e as Error).message : ''; out.classList.add('bad'); }
    };
    for (const c of [amount, cur, style, kw]) { c.addEventListener('input', calc); c.addEventListener('change', calc); }
    const ok = el('button', { class: 'primary', textContent: 'Вставить', onclick: () => { if (!text) return; close(); insertText(host, text); } });
    box.addEventListener('keydown', (e) => { if (e.key === 'Enter' && text) ok.click(); });
    box.append(
      el('h2', { textContent: 'Сумма прописью' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Сумма' }), amount),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Валюта' }), cur),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Вид' }), style),
      el('div', { class: 'dlg-row' }, el('label', {}, kw, ' Копейки словами')),
      out,
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), ok),
    );
    calc();
    amount.focus();
  });
}

export function nameDialog(host: Host): void {
  dialog((box, close) => {
    const fio = el('input', { type: 'text', value: selected(host), placeholder: 'Иванов Иван Иванович' });
    const pos = el('input', { type: 'text', placeholder: 'Должность (необязательно): Генеральный директор' });
    const gender = el('select', {}, el('option', { value: '', textContent: 'Пол: определить сам' }), el('option', { value: 'm', textContent: 'Мужской' }), el('option', { value: 'f', textContent: 'Женский' }));
    const ini = el('input', { type: 'checkbox' });
    const list = el('div', { class: 'ru-cases' });
    const render = () => {
      list.replaceChildren();
      if (!fio.value.trim() && !pos.value.trim()) return;
      try {
        const name = fio.value.trim() ? declineFullName(fio.value, { gender: (gender.value || undefined) as NameGender | undefined, initials: ini.checked }) : null;
        for (const c of CASES) {
          const text = [pos.value.trim() ? declinePosition(pos.value, c) : '', name?.forms[c] ?? ''].filter(Boolean).join(' ');
          list.append(el('button', { class: 'ru-case', title: 'Вставить', onclick: () => { close(); insertText(host, text); } }, el('span', { class: 'ru-case-label', textContent: CASE_LABELS[c as Case] }), el('span', { textContent: text })));
        }
      } catch (e) { list.textContent = (e as Error).message; }
    };
    for (const c of [fio, pos, gender, ini]) { c.addEventListener('input', render); c.addEventListener('change', render); }
    box.append(
      el('h2', { textContent: 'ФИО и должность по падежам' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'ФИО' }), fio),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Должность' }), pos),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Пол' }), gender),
      el('div', { class: 'dlg-row' }, el('label', {}, ini, ' Имя и отчество инициалами (Иванову И. И.)')),
      list,
      el('p', { textContent: 'Склонение — по правилам, без словаря: редкие и иностранные фамилии проверяйте глазами. Щелчок по строке вставляет её в текст.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Закрыть', onclick: close })),
    );
    render();
    fio.focus();
    void toast;
  });
}
