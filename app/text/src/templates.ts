// Шаблоны документов. Готовые лежат в программе (открытые, одинаковые у всех). «Мои шаблоны» — только на этом компьютере:
// хранятся в базе браузера/окна программы (IndexedDB), нигде не публикуются и не выкладываются.

import type { Node as PMNode } from 'prosemirror-model';
import { newMeta, type DocMeta } from './docx';
import { body, buildForm, FORMS, INDENT, para, twoCols, type FormType } from './forms';
import { schema } from './schema';

export interface Built { doc: PMNode; meta: DocMeta; name: string }
export interface BuiltIn { id: string; title: string; hint: string; build: () => Built }

const n = schema.nodes;

function gostMeta(left = 56.7, right = 28.35, size = 14, line = 1, where: 'header' | 'footer' = 'header'): DocMeta {
  const meta = newMeta();
  meta.page = { w: 595.3, h: 841.9, top: 56.7, bottom: 56.7, left, right };
  meta.defaults = { font: 'Times New Roman', size, line, before: 0, after: 0 };
  const field = n.doc.create(null, [n.paragraph.create({ align: 'center' }, n.field.create({ kind: 'page' }))]).toJSON();
  if (where === 'header') meta.hf.header = field; else meta.hf.footer = field;
  meta.hf.differentFirst = true;
  return meta;
}

const h = (text: string, extra: Record<string, unknown> = {}): PMNode => n.heading.create({ level: 1, align: 'center', before: 12, after: 12, ...extra }, schema.text(text.toUpperCase(), [schema.marks.strong.create()]));
const ph = (text: string): string => `[${text}]`;

/** Бланк из мастера: все поля заменены подсказками в скобках, номер и дата — пустые. */
function fromForm(type: FormType): Built {
  const spec = FORMS.find((f) => f.type === type)!;
  const values: Record<string, string> = {};
  for (const f of spec.fields) if (f.kind !== 'date') values[f.key] = f.value ?? ph(f.label);
  const b = buildForm(type, { values, numberLabel: null });
  return { ...b, name: `${spec.title} (шаблон)` };
}

const BUILT_IN: BuiltIn[] = [
  { id: 'blank-7097', title: 'Пустой документ по ГОСТ Р 7.0.97', hint: 'Поля 20/10/20/20 мм, Times New Roman 14, номера страниц вверху со второй страницы',
    build: () => ({ doc: n.doc.create(null, [n.paragraph.create({ align: 'justify', first: INDENT })]), meta: gostMeta(), name: 'Документ' }) },
  { id: 'letter', title: 'Письмо', hint: 'Реквизиты по ГОСТ Р 7.0.97: организация, дата и номер, адресат, заголовок, текст, подпись', build: () => fromForm('letter') },
  { id: 'order', title: 'Приказ', hint: 'Преамбула, «ПРИКАЗЫВАЮ», пункты, контроль, подпись', build: () => fromForm('order') },
  { id: 'act', title: 'Акт', hint: 'Основание, комиссия, что установлено, подписи', build: () => fromForm('act') },
  { id: 'protocol', title: 'Протокол', hint: 'Председатель, секретарь, присутствовали, повестка, «СЛУШАЛИ / РЕШИЛИ»', build: () => fromForm('protocol') },
  {
    id: 'memo', title: 'Служебная записка', hint: 'От кого, кому, тема, текст, подпись',
    build: () => ({
      name: 'Служебная записка (шаблон)', meta: gostMeta(),
      doc: n.doc.create(null, [
        para({ text: ph('Наименование организации'), align: 'center', bold: true }),
        para({ text: 'СЛУЖЕБНАЯ ЗАПИСКА', align: 'center', bold: true, before: 12, after: 6 }),
        twoCols('___.___.______', '№ ___'),
        para({ text: ph('Должность, Ф. И. О. адресата'), indent: 255, before: 12 }),
        para({ text: `От: ${ph('должность, Ф. И. О. автора')}`, indent: 255 }),
        para({ text: ph('О чём записка'), bold: true, before: 18, after: 12 }),
        ...body(`${ph('Суть вопроса: что произошло или что нужно.')}\n${ph('Предложение: что просите сделать, к какому сроку.')}`),
        para({ text: '', after: 6 }),
        twoCols(ph('Должность'), ph('И. И. Иванов')),
      ]),
    }),
  },
  {
    id: 'statement', title: 'Заявление', hint: 'Кому, от кого, просьба, дата и подпись',
    build: () => ({
      name: 'Заявление (шаблон)', meta: gostMeta(),
      doc: n.doc.create(null, [
        para({ text: ph('Должность руководителя, наименование организации'), indent: 255 }),
        para({ text: ph('Ф. И. О. руководителя в дательном падеже'), indent: 255 }),
        para({ text: `от ${ph('Ф. И. О., должность заявителя')}`, indent: 255, before: 6 }),
        para({ text: 'ЗАЯВЛЕНИЕ', align: 'center', bold: true, before: 24, after: 12 }),
        ...body(`Прошу ${ph('что именно просите, на каком основании, на какой срок')}.`),
        para({ text: '', after: 12 }),
        twoCols('___.___.______', ph('подпись / И. И. Иванов')),
      ]),
    }),
  },
  {
    id: 'proxy', title: 'Доверенность (простая)', hint: 'Кто доверяет, кому, что поручает, срок',
    build: () => ({
      name: 'Доверенность (шаблон)', meta: gostMeta(),
      doc: n.doc.create(null, [
        para({ text: 'ДОВЕРЕННОСТЬ', align: 'center', bold: true, after: 6 }),
        twoCols(ph('Город'), '___.___.______'),
        ...body(`Я, ${ph('Ф. И. О., дата рождения, паспорт: серия, номер, кем и когда выдан, адрес')}, доверяю ${ph('Ф. И. О. представителя, паспортные данные')} ${ph('что поручается: получить, подписать, представлять интересы…')}.`),
        ...body('Доверенность выдана без права передоверия сроком на ' + ph('срок') + '.'),
        para({ text: '', after: 12 }),
        twoCols('', ph('подпись / И. И. Иванов')),
      ]),
    }),
  },
  {
    id: 'report-732', title: 'Отчёт о НИР (ГОСТ 7.32)', hint: 'Титульный лист и разделы; поля 30/15/20/20 мм, интервал 1,5, номера внизу',
    build: () => ({
      name: 'Отчёт (шаблон)', meta: gostMeta(85, 42.5, 14, 1.5, 'footer'),
      doc: n.doc.create(null, [
        para({ text: ph('Наименование вышестоящей организации'), align: 'center' }),
        para({ text: ph('Наименование организации-исполнителя'), align: 'center', bold: true, before: 6 }),
        para({ text: 'УТВЕРЖДАЮ', align: 'right', before: 36 }),
        para({ text: ph('должность, подпись, И. О. Фамилия'), align: 'right' }),
        para({ text: 'ОТЧЁТ О НАУЧНО-ИССЛЕДОВАТЕЛЬСКОЙ РАБОТЕ', align: 'center', bold: true, before: 96 }),
        para({ text: ph('Тема работы'), align: 'center', bold: true, before: 12 }),
        para({ text: ph('вид отчёта: промежуточный / заключительный'), align: 'center', before: 12 }),
        para({ text: `Руководитель работы: ${ph('должность, подпись, И. О. Фамилия')}`, before: 96 }),
        para({ text: `${ph('Город')} ${ph('год')}`, align: 'center', before: 96 }),
        n.paragraph.create({ pageBreakBefore: true, align: 'center' }, schema.text('СОДЕРЖАНИЕ', [schema.marks.strong.create()])),
        para({ text: ph('Оглавление: введение, разделы, заключение, список источников, приложения') }),
        n.paragraph.create({ pageBreakBefore: true, align: 'center' }, schema.text('РЕФЕРАТ', [schema.marks.strong.create()])),
        ...body(ph('Объём отчёта, число рисунков, таблиц, источников; перечень ключевых слов; текст реферата')),
        h('Введение'), ...body(ph('Актуальность, цель и задачи работы')),
        h('1 Основная часть'), ...body(ph('Методы, ход и результаты работы')),
        h('Заключение'), ...body(ph('Выводы по результатам работы')),
        h('Список использованных источников'), ...body(ph('1 Автор. Название. — Город: Издательство, год. — страницы.')),
      ]),
    }),
  },
];

export const builtInTemplates = (): BuiltIn[] => BUILT_IN;

// ---------------------------------------------------------------- «Мои шаблоны»: только на этом компьютере

export interface MyTemplate { id: string; name: string; at: number; bytes: Uint8Array }

const DB = 'ens-text-private';
const STORE = 'templates';

function open(): Promise<IDBDatabase> {
  return new Promise((res, rej) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore(STORE, { keyPath: 'id' });
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error ?? new Error('Хранилище недоступно'));
  });
}

async function tx<T>(mode: IDBTransactionMode, f: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await open();
  return new Promise<T>((res, rej) => {
    const t = db.transaction(STORE, mode);
    const rq = f(t.objectStore(STORE));
    rq.onsuccess = () => res(rq.result);
    rq.onerror = () => rej(rq.error);
    t.oncomplete = () => db.close();
  });
}

export async function listMine(): Promise<MyTemplate[]> {
  try { return ((await tx('readonly', (s) => s.getAll())) as MyTemplate[]).sort((a, b) => b.at - a.at); } catch { return []; }
}
export async function saveMine(name: string, bytes: Uint8Array, id?: string): Promise<void> {
  const rec: MyTemplate = { id: id ?? `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, name, at: Date.now(), bytes };
  await tx('readwrite', (s) => s.put(rec));
}
export async function removeMine(id: string): Promise<void> { await tx('readwrite', (s) => s.delete(id)); }
export async function getMine(id: string): Promise<MyTemplate | undefined> { return tx('readonly', (s) => s.get(id)) as Promise<MyTemplate | undefined>; }
