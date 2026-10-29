// Окна мастера служебных документов и журнала регистрации.

import type { Node as PMNode } from 'prosemirror-model';
import type { DocMeta } from './docx';
import { buildForm, FORMS, loadJournal, nextNumber, numberLabel, saveJournal, type FormType } from './forms';
import { dialog, el, toast } from './ui';

const today = (): string => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };

/** Мастер: выбрать вид, заполнить поля, присвоить номер из журнала и создать новый документ. */
export function wizardDialog(create: (doc: PMNode, meta: DocMeta, name: string) => void, initial: FormType = 'letter'): void {
  dialog((box, close) => {
    const journal = loadJournal();
    const kind = el('select', {}, ...FORMS.map((f) => el('option', { value: f.type, textContent: f.title })));
    kind.value = initial;
    const form = el('div', { class: 'wiz-form' });
    const prefix = el('input', { type: 'text', placeholder: 'Индекс, например 01-08 (необязательно)', title: 'Индекс дела: номер будет «01-08/12»' });
    const reg = el('input', { type: 'checkbox', checked: true });
    const numInfo = el('div', { class: 'wiz-num' });
    let inputs: Record<string, HTMLInputElement | HTMLTextAreaElement> = {};

    const spec = () => FORMS.find((f) => f.type === kind.value)!;
    const year = () => parseInt((inputs.date?.value || today()).slice(0, 4), 10);
    const next = () => nextNumber(journal, spec().type, year());
    const shownNo = () => numberLabel(prefix.value, next());

    const refreshNum = () => { numInfo.textContent = reg.checked ? `Номер из журнала «${spec().journal}»: № ${shownNo()}` : 'Номер не присваивается: в документе будет «№ ___»'; };

    const build = () => {
      form.replaceChildren();
      inputs = {};
      prefix.value = journal.prefix[spec().type] ?? '';
      for (const f of spec().fields) {
        const inp = f.kind === 'area'
          ? el('textarea', { rows: f.rows ?? 3, placeholder: f.hint ?? '', value: f.value ?? '' })
          : el('input', { type: f.kind === 'date' ? 'date' : 'text', placeholder: f.hint ?? '', value: f.kind === 'date' ? today() : f.value ?? '' });
        inp.addEventListener('input', refreshNum);
        inputs[f.key] = inp;
        form.append(el('div', { class: 'dlg-row' }, el('label', { textContent: f.label }), inp));
      }
      refreshNum();
    };
    kind.addEventListener('change', build);
    prefix.addEventListener('input', refreshNum);
    reg.addEventListener('change', refreshNum);

    const go = () => {
      const values: Record<string, string> = {};
      for (const [k, i] of Object.entries(inputs)) values[k] = i.value.trim();
      let label: string | null = null;
      if (reg.checked) {
        const no = next();
        label = numberLabel(prefix.value, no);
        journal.prefix[spec().type] = prefix.value.trim();
        journal.entries.push({ id: `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, type: spec().type, year: year(), no, label, date: values.date, subject: values.subject || values.meeting || '', at: Date.now() });
        saveJournal(journal);
      }
      const built = buildForm(spec().type, { values, numberLabel: label });
      close();
      create(built.doc, built.meta, built.name);
      if (label) toast(`Зарегистрировано: ${spec().title} № ${label}`);
    };

    box.append(
      el('h2', { textContent: 'Служебный документ' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Вид' }), kind),
      form,
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Индекс' }), prefix),
      el('div', { class: 'dlg-row' }, el('label', {}, reg, ' Присвоить номер и записать в журнал регистрации')),
      numInfo,
      el('p', { textContent: 'Оформление по ГОСТ Р 7.0.97-2016: поля 20/10/20/20 мм, Times New Roman 14, номера страниц вверху со второй страницы. Пустые поля останутся пустыми строками для ручного заполнения.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Создать документ', onclick: go })),
    );
    build();
  }, false);
}

/** Журнал регистрации: все присвоенные номера; строку можно убрать (последний номер освободится). */
export function journalDialog(): void {
  dialog((box, close) => {
    const journal = loadJournal();
    const list = el('div', { class: 'jr-list' });
    const render = () => {
      list.replaceChildren();
      const rows = [...journal.entries].sort((a, b) => b.at - a.at);
      if (!rows.length) { list.append(el('p', { textContent: 'Журнал пуст. Номера появляются, когда документ создаётся мастером с отметкой «Присвоить номер».' })); return; }
      for (const e of rows) {
        const title = FORMS.find((f) => f.type === e.type)?.title ?? e.type;
        list.append(el('div', { class: 'jr-row' },
          el('span', { class: 'jr-no', textContent: `№ ${e.label}` }),
          el('span', { textContent: `${title}, ${e.date ? e.date.split('-').reverse().join('.') : ''}` }),
          el('span', { class: 'jr-subj', textContent: e.subject }),
          el('button', { textContent: '✕', title: 'Убрать из журнала', onclick: () => { journal.entries = journal.entries.filter((x) => x.id !== e.id); saveJournal(journal); render(); } })));
      }
    };
    box.append(
      el('h2', { textContent: 'Журнал регистрации' }),
      list,
      el('p', { textContent: 'Журнал хранится на этом компьютере. Номер считается отдельно по виду документа и году.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Закрыть', onclick: close })),
    );
    render();
  }, false);
}
