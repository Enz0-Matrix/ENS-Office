// SPDX-License-Identifier: AGPL-3.0-or-later
// Окно «Источники» (Ссылки → Источники): карточки, предпросмотр записи по ГОСТ, вставка ссылки «[N, с. …]».

import { NodeSelection, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { formatSource, KIND_LABEL, numbering, sourcesOf, type Source, type SourceKind } from './biblio';
import { schema } from './schema';
import { dialog, el, toast } from './ui';

/** Поля карточки по видам: [ключ, подпись, подсказка]. */
const FIELDS: Record<SourceKind, [keyof Source, string, string][]> = {
  book: [['authors', 'Авторы', 'Иванов И. И. — по одному в строке'], ['title', 'Заглавие', ''], ['info', 'Сведения', 'учебник для вузов'], ['edition', 'Издание', '2-е изд., перераб. и доп.'],
    ['city', 'Город', 'Москва'], ['publisher', 'Издательство', 'Юрайт'], ['year', 'Год', '2024'], ['pages', 'Объём', '256 с.'], ['isbn', 'ISBN', '']],
  article: [['authors', 'Авторы', 'по одному в строке'], ['title', 'Заглавие статьи', ''], ['container', 'Журнал', 'Вопросы экономики'], ['year', 'Год', ''], ['issue', 'Номер', '3'], ['pages', 'Страницы', '15–20']],
  web: [['authors', 'Авторы (если есть)', ''], ['title', 'Заглавие', ''], ['info', 'Сведения', 'сайт'], ['container', 'Сайт', ''], ['year', 'Год', ''], ['url', 'Адрес (URL)', 'https://…'], ['accessed', 'Дата обращения', 'ДД.ММ.ГГГГ']],
  law: [['title', 'Название акта', 'Об образовании в Российской Федерации'], ['info', 'Вид, дата и номер', 'федеральный закон от 29.12.2012 № 273-ФЗ'], ['container', 'Источник публикации', 'Собрание законодательства РФ'], ['year', 'Год', ''], ['issue', 'Номер', '53'], ['url', 'Адрес (URL)', ''], ['accessed', 'Дата обращения', '']],
  other: [['text', 'Готовая запись', 'Текст записи целиком']],
};

const blank = (id: number): Source => ({ id, kind: 'book', authors: [], title: '' });

export function sourcesDialog(view: EditorView): void {
  dialog((box, close) => {
    box.classList.add('sources-dialog');
    let list = sourcesOf(view.state.doc).map((s) => ({ ...s, authors: [...s.authors] }));
    let current: Source | null = list[0] ?? null;
    const listEl = el('div', { class: 'src-list' });
    const form = el('div', { class: 'src-form' });
    const preview = el('p', { class: 'src-preview' });
    const page = el('input', { type: 'text', ariaLabel: 'Страница', placeholder: 'страница', maxLength: 20 });
    const status = el('p', { class: 'src-status', role: 'status' });
    const save = () => {
      view.dispatch(view.state.tr.setDocAttribute('sources', list.map((s) => ({ ...s }))));
    };
    const drawList = () => {
      const { num, cited } = numbering(view.state.doc);
      listEl.replaceChildren(...list.map((s) => {
        const b = el('button', { class: `src-item${s === current ? ' on' : ''}`, textContent: `${num.get(s.id) ?? '•'}. ${(s.kind === 'other' ? s.text : (s.authors[0] ? s.authors[0] + ' — ' : '') + s.title) || '(пусто)'}`,
          onclick: () => { current = s; drawList(); drawForm(); } });
        if (!cited.has(s.id)) b.title = 'Нет ссылок в тексте';
        return b;
      }), el('button', { class: 'src-add', textContent: '+ Новый источник', onclick: () => {
        const s = blank(list.reduce((m, x) => Math.max(m, x.id), 0) + 1);
        list = [...list, s]; current = s; drawList(); drawForm();
      } }));
    };
    const drawForm = () => {
      form.replaceChildren();
      if (!current) { preview.textContent = ''; return; }
      const s = current;
      const kind = el('select', { ariaLabel: 'Вид источника' }, ...(Object.keys(KIND_LABEL) as SourceKind[]).map((k) => new Option(KIND_LABEL[k], k)));
      kind.value = s.kind;
      kind.addEventListener('change', () => { s.kind = kind.value as SourceKind; drawForm(); });
      form.append(el('div', { class: 'dlg-row' }, el('label', { textContent: 'Вид' }), kind));
      for (const [key, label, hint] of FIELDS[s.kind]) {
        const multi = key === 'authors' || key === 'text';
        const input = multi ? el('textarea', { rows: key === 'text' ? 4 : 3, placeholder: hint, ariaLabel: label }) : el('input', { type: 'text', placeholder: hint, ariaLabel: label });
        input.value = key === 'authors' ? s.authors.join('\n') : String(s[key] ?? '');
        input.addEventListener('input', () => {
          if (key === 'authors') s.authors = input.value.split('\n'); else (s as unknown as Record<string, string>)[key] = input.value;
          preview.textContent = formatSource(s);
        });
        form.append(el('div', { class: 'dlg-row' }, el('label', { textContent: label }), input));
      }
      preview.textContent = formatSource(s);
    };
    const apply = el('button', { textContent: 'Сохранить карточки', onclick: () => { save(); status.textContent = 'Сохранено. Список литературы обновлён.'; drawList(); } });
    const del = el('button', { class: 'danger', textContent: 'Удалить источник', onclick: () => {
      if (!current) return;
      const { cited } = numbering(view.state.doc);
      const gone = current;
      list = list.filter((x) => x !== gone);
      current = list[0] ?? null;
      save();
      status.textContent = cited.has(gone.id) ? 'Источник удалён. Ссылки на него показываются как «[?]» — их видно в тексте. Отмена — Ctrl+Z.' : 'Источник удалён.';
      drawList(); drawForm();
    } });
    const cite = el('button', { class: 'primary', textContent: 'Вставить ссылку', onclick: () => {
      if (!current) { status.textContent = 'Выберите источник.'; return; }
      if (!formatSource(current)) { status.textContent = 'Заполните карточку источника.'; return; }
      save();
      const id = current.id, pg = page.value.trim();
      close();
      const { from, to } = view.state.selection;
      // Ссылка рядом с другой ссылкой на этот же… — просто ещё одна ссылка; «[1; 3]» — выделите ссылку и вставьте второй источник.
      const sel = view.state.selection;
      const prev = sel instanceof NodeSelection && sel.node.type === schema.nodes.cite ? sel.node : null;
      const node = prev
        ? schema.nodes.cite.create({ ids: [...new Set([...(prev.attrs.ids as number[]), id])], page: prev.attrs.page || pg })
        : schema.nodes.cite.create({ ids: [id], page: pg });
      const tr = view.state.tr.replaceWith(from, to, node);
      view.dispatch(tr.setSelection(TextSelection.create(tr.doc, from + 1)));
      view.focus();
    } });
    box.append(el('h2', { textContent: 'Источники (ГОСТ Р 7.0.100-2018)' }),
      el('div', { class: 'src-body' }, listEl, el('div', {}, form, el('b', { textContent: 'Запись в списке:' }), preview)),
      status,
      el('div', { class: 'dlg-actions' }, del, apply, el('label', { class: 'src-page' }, 'С. ', page), cite, el('button', { textContent: 'Закрыть', onclick: () => { save(); close(); } })));
    drawList(); drawForm();
  }, false);
}

/** Вставить список литературы с заголовком (ГОСТ 7.32: «СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ»). Второй не вставляем. */
export function insertBiblio(view: EditorView): void {
  let have = -1;
  view.state.doc.forEach((n, pos) => { if (have < 0 && n.type === schema.nodes.biblio) have = pos; });
  if (have >= 0) {
    view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, have)).scrollIntoView());
    toast('Список литературы уже есть — он обновляется сам.');
    return;
  }
  const { $from } = view.state.selection;
  const at = $from.depth ? $from.after(1) : $from.pos;
  const title = schema.nodes.paragraph.create({ align: 'center', after: 12, pageBreakBefore: false }, schema.text('СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ', [schema.marks.strong.create()]));
  view.dispatch(view.state.tr.insert(at, [title, schema.nodes.biblio.create()]).scrollIntoView());
  view.focus();
}
