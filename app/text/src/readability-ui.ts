// Окно «Читаемость».

import { Selection, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { analyze, minutesText, type Issue } from './readability';
import { dialog, el } from './ui';

const KIND: Record<Issue['kind'], string> = { long: 'Длинное предложение', repeat: 'Повтор слова', clerical: 'Канцеляризм' };

export function readabilityDialog(host: { view: () => EditorView }): void {
  dialog((box, close) => {
    const v = host.view();
    const sel = v.state.selection;
    const part = !sel.empty;
    const r = analyze(v.state.doc, part ? sel.from : 0, part ? sel.to : v.state.doc.content.size);
    const stat = (label: string, value: string) => el('div', { class: 'rd-stat' }, el('div', { class: 'rd-val', textContent: value }), el('div', { class: 'rd-lab', textContent: label }));
    const list = el('div', { class: 'rd-list' });
    const show = (kind: 'all' | Issue['kind']) => {
      list.replaceChildren();
      const rows = r.issues.filter((i) => kind === 'all' || i.kind === kind);
      if (!rows.length) list.append(el('p', { textContent: 'Здесь замечаний нет.' }));
      for (const i of rows) {
        list.append(el('div', { class: 'rd-row' },
          el('span', { class: `rd-kind ${i.kind}`, textContent: KIND[i.kind] }),
          el('span', { class: 'rd-text', textContent: i.text }),
          el('span', { class: 'rd-note', textContent: i.note }),
          el('button', { textContent: 'Показать', onclick: () => {
            const max = v.state.doc.content.size;
            const from = Math.min(i.pos, max), to = Math.min(i.pos + i.len, max);
            v.dispatch(v.state.tr.setSelection(from < to ? TextSelection.create(v.state.doc, from, to) : Selection.near(v.state.doc.resolve(from))).scrollIntoView());
            close();
            v.focus();
          } })));
      }
    };
    const tabs = el('div', { class: 'rd-tabs' });
    const tab = (label: string, kind: 'all' | Issue['kind']) => {
      const n = kind === 'all' ? r.issues.length : r.issues.filter((i) => i.kind === kind).length;
      const b = el('button', { textContent: `${label} (${n})`, onclick: () => { for (const x of tabs.children) x.classList.remove('on'); b.classList.add('on'); show(kind); } });
      return b;
    };
    tabs.append(tab('Все', 'all'), tab('Длинные', 'long'), tab('Повторы', 'repeat'), tab('Канцеляризмы', 'clerical'));
    tabs.firstElementChild!.classList.add('on');
    show('all');
    box.append(
      el('h2', { textContent: part ? 'Читаемость выделенного' : 'Читаемость' }),
      r.words
        ? el('div', { class: 'rd-stats' },
          stat('слов', String(r.words)), stat('предложений', String(r.sentences)), stat('время чтения', minutesText(r.minutes)),
          stat('слов в предложении (в среднем)', String(r.avgSentence)), stat('индекс Флеша', `${r.flesch}`))
        : el('p', { textContent: 'Текста для проверки нет.' }),
      r.words ? el('p', { class: 'rd-level', textContent: `Оценка: ${r.level}.` }) : el('span'),
      tabs, list,
      el('p', { textContent: 'Подсказки по правилам, без интернета: это повод перечитать, а не приговор. Индекс Флеша для русского языка (Оборнева): ниже 40 — трудно, 60 и выше — легко.' }),
      el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: close })),
    );
  }, false);
}
