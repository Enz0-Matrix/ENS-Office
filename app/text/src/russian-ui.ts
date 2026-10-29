// Окна «по-русски»: сумма прописью, ФИО и должность по падежам, дата.

import type { EditorView } from 'prosemirror-view';
import { CASES, CASE_LABELS, CURRENCIES, declineFullName, declinePosition, sumWords, type Case, type NameGender } from './russian';
import { addCalendarDays, addWorkdays, fromYmd, parseCorrections, toYmd, workdaysBetween } from './calendar';
import { dateRu } from './russian';
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
    // На «change» текстовых полей список не перестраиваем: уход фокуса при щелчке по строке съедал бы сам щелчок.
    for (const c of [fio, pos]) c.addEventListener('input', render);
    for (const c of [gender, ini]) c.addEventListener('change', render);
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

const CAL_KEY = 'ens-text.calendar';
const WEEKDAYS = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];

function loadCorrections(): string {
  try { return localStorage.getItem(CAL_KEY) ?? ''; } catch { return ''; }
}
function saveCorrections(v: string): void {
  try { localStorage.setItem(CAL_KEY, v); } catch { /* без хранилища поправки живут до закрытия окна */ }
}

/** Срок в рабочих или календарных днях с учётом праздников РФ; сколько рабочих дней между датами. */
export function deadlineDialog(host: Host): void {
  dialog((box, close) => {
    const start = el('input', { type: 'date', value: toYmd(new Date()) });
    const mode = el('select', {},
      el('option', { value: 'work', textContent: 'через N рабочих дней' }),
      el('option', { value: 'cal', textContent: 'через N календарных дней' }),
      el('option', { value: 'between', textContent: 'рабочих дней до даты' }));
    const n = el('input', { type: 'number', value: '10', step: '1' });
    const end = el('input', { type: 'date', value: toYmd(new Date()) });
    const roll = el('input', { type: 'checkbox', checked: true });
    const corr = el('textarea', { rows: 3, placeholder: '+2026-05-04 — выходной; =2026-02-21 — рабочая суббота', value: loadCorrections() });
    const out = el('div', { class: 'ru-out' });
    let text = '';
    const row = (label: string, ...kids: Node[]) => el('div', { class: 'dlg-row' }, el('label', { textContent: label }), ...kids);
    const nRow = row('Сколько', n);
    const endRow = row('До даты', end);
    const rollRow = el('div', { class: 'dlg-row' }, el('label', {}, roll, ' Если срок выпал на нерабочий день — перенести на рабочий'));
    const calc = () => {
      nRow.hidden = mode.value === 'between';
      endRow.hidden = mode.value !== 'between';
      rollRow.hidden = mode.value !== 'cal';
      text = '';
      try {
        const { corr: c, errors } = parseCorrections(corr.value);
        const a = fromYmd(start.value);
        if (mode.value === 'between') {
          const b = fromYmd(end.value);
          const k = workdaysBetween(a, b, c);
          out.textContent = `Рабочих дней: ${k}`;
          text = String(k);
        } else {
          const k = parseInt(n.value, 10);
          if (!Number.isFinite(k)) throw new Error('Введите число дней');
          const r = mode.value === 'work' ? addWorkdays(a, k, c) : addCalendarDays(a, k, c, roll.checked);
          text = dateRu(r);
          out.textContent = `${text} (${WEEKDAYS[r.getDay()]})`;
        }
        if (errors.length) out.textContent += `  ·  не понял поправки: ${errors.join(', ')}`;
        out.classList.remove('bad');
      } catch (e) { out.textContent = (e as Error).message; out.classList.add('bad'); }
    };
    for (const c of [start, mode, n, end, roll, corr]) { c.addEventListener('input', calc); c.addEventListener('change', calc); }
    corr.addEventListener('change', () => saveCorrections(corr.value));
    const ok = el('button', { class: 'primary', textContent: 'Вставить', onclick: () => { if (!text) return; saveCorrections(corr.value); close(); insertText(host, text); } });
    box.append(
      el('h2', { textContent: 'Срок в рабочих днях' }),
      row('Начиная с', start), row('Вид', mode), nRow, endRow, rollRow,
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Поправки' }), corr),
      out,
      el('p', { textContent: 'Календарь расчётный: праздники РФ, выпавшие на выходные, переносятся по ст. 112 ТК РФ. Ежегодные постановления Правительства переносят дни иначе — для важных сроков сверьтесь с постановлением и внесите отличия в «Поправки» (+дата — выходной, =дата — рабочий день, через «;» или с новой строки). Поправки запоминаются.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Закрыть', onclick: close }), ok),
    );
    calc();
    n.focus();
    void toast;
  });
}
