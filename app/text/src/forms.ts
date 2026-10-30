// Мастера служебных документов: письмо, приказ, акт, протокол. Реквизиты и поля — по ГОСТ Р 7.0.97-2016.
// Чистые функции: из ответов формы строят документ ProseMirror и параметры страницы; журнал регистрации — отдельно, в localStorage.

import type { Mark, Node as PMNode } from 'prosemirror-model';
import { newMeta, type DocMeta } from './docx';
import { schema } from './schema';

export type FormType = 'letter' | 'order' | 'act' | 'protocol';

export interface Field { key: string; label: string; kind: 'text' | 'area' | 'date'; hint?: string; value?: string; rows?: number }

export interface FormSpec {
  type: FormType;
  title: string;
  /** Название счётчика в журнале. */
  journal: string;
  fields: Field[];
}

const common = (executor = true): Field[] => [
  { key: 'org', label: 'Организация', kind: 'area', hint: 'Общество с ограниченной ответственностью «Ромашка»', rows: 2 },
  { key: 'date', label: 'Дата', kind: 'date' },
  ...(executor ? [{ key: 'executor', label: 'Исполнитель', kind: 'text', hint: 'Петров П. П., тел. 8 (495) 000-00-00' } as Field] : []),
];
const signature: Field[] = [
  { key: 'position', label: 'Должность подписывающего', kind: 'text', hint: 'Директор' },
  { key: 'signer', label: 'Подписывает', kind: 'text', hint: 'И. И. Иванов' },
];

export const FORMS: FormSpec[] = [
  {
    type: 'letter', title: 'Письмо', journal: 'Исходящие письма',
    fields: [
      ...common(),
      { key: 'reply', label: 'На №, от', kind: 'text', hint: 'на № 15 от 01.10.2026 (если это ответ)' },
      { key: 'addressee', label: 'Адресат', kind: 'area', hint: 'Директору ООО «Вектор»\nСидорову С. С.\n125000, г. Москва, ул. Мира, д. 1', rows: 3 },
      { key: 'subject', label: 'Заголовок к тексту', kind: 'text', hint: 'О согласовании сроков поставки' },
      { key: 'body', label: 'Текст', kind: 'area', hint: 'Уважаемый Сергей Сергеевич!\n\nТекст письма. Абзацы — через Enter.', rows: 6 },
      { key: 'attach', label: 'Приложения', kind: 'area', hint: 'По одному на строке: Договор на 3 л. в 1 экз.', rows: 2 },
      ...signature,
    ],
  },
  {
    type: 'order', title: 'Приказ', journal: 'Приказы',
    fields: [
      ...common(false),
      { key: 'city', label: 'Место составления', kind: 'text', hint: 'г. Москва' },
      { key: 'subject', label: 'О чём приказ', kind: 'text', hint: 'О назначении ответственного за пожарную безопасность' },
      { key: 'preamble', label: 'Преамбула', kind: 'area', hint: 'В соответствии с … , в целях …', rows: 2 },
      { key: 'items', label: 'Пункты («ПРИКАЗЫВАЮ»)', kind: 'area', hint: 'По одному на строке; нумеруются сами', rows: 5 },
      { key: 'control', label: 'Контроль', kind: 'text', value: 'Контроль за исполнением настоящего приказа оставляю за собой.' },
      ...signature,
    ],
  },
  {
    type: 'act', title: 'Акт', journal: 'Акты',
    fields: [
      ...common(false),
      { key: 'city', label: 'Место составления', kind: 'text', hint: 'г. Москва' },
      { key: 'subject', label: 'Заголовок («Об …»)', kind: 'text', hint: 'О приёмке выполненных работ' },
      { key: 'basis', label: 'Основание', kind: 'text', hint: 'Основание: приказ от … № …' },
      { key: 'commission', label: 'Комиссия', kind: 'area', hint: 'Председатель — Иванов И. И., директор\nЧлены комиссии: Петров П. П., главный инженер', rows: 3 },
      { key: 'body', label: 'Что установлено', kind: 'area', hint: 'Комиссия установила: …', rows: 5 },
      ...signature,
    ],
  },
  {
    type: 'protocol', title: 'Протокол', journal: 'Протоколы',
    fields: [
      ...common(false),
      { key: 'city', label: 'Место проведения', kind: 'text', hint: 'г. Москва' },
      { key: 'meeting', label: 'Чего протокол', kind: 'text', hint: 'заседания комиссии по охране труда' },
      { key: 'chair', label: 'Председатель', kind: 'text', hint: 'И. И. Иванов' },
      { key: 'secretary', label: 'Секретарь', kind: 'text', hint: 'П. П. Петров' },
      { key: 'present', label: 'Присутствовали', kind: 'area', hint: 'По одному на строке: Сидоров С. С. — главный инженер', rows: 3 },
      { key: 'agenda', label: 'Повестка дня', kind: 'area', hint: 'По одному вопросу на строке', rows: 3 },
    ],
  },
];

// ---------------------------------------------------------------- журнал регистрации

export interface JournalEntry { id: string; type: FormType; year: number; no: number; label: string; date: string; subject: string; at: number }
export interface Journal { prefix: Record<string, string>; entries: JournalEntry[] }

const KEY = 'ens-text.journal';
export const emptyJournal = (): Journal => ({ prefix: {}, entries: [] });

export function loadJournal(): Journal {
  try {
    const j = JSON.parse(localStorage.getItem(KEY) ?? 'null') as Journal | null;
    if (j && Array.isArray(j.entries)) return { prefix: j.prefix ?? {}, entries: j.entries };
  } catch { /* пустой журнал */ }
  return emptyJournal();
}
export function saveJournal(j: Journal): void {
  try { localStorage.setItem(KEY, JSON.stringify(j)); } catch { /* без хранилища журнал живёт до закрытия вкладки */ }
}

/** Следующий номер в своём счётчике за год: наибольший зарегистрированный + 1. */
export function nextNumber(j: Journal, type: FormType, year: number): number {
  return j.entries.filter((e) => e.type === type && e.year === year).reduce((m, e) => Math.max(m, e.no), 0) + 1;
}

export const numberLabel = (prefix: string, no: number): string => `${prefix ? prefix.trim() + '/' : ''}${no}`;

// ---------------------------------------------------------------- построение документа

const m = schema.marks;
const n = schema.nodes;

export type Para = { text?: string; bold?: boolean; align?: 'left' | 'center' | 'right' | 'justify'; indent?: number; first?: number; before?: number; after?: number; size?: number; caps?: boolean; italic?: boolean };

export const INDENT = 35.4; // 1,25 см

export function para(p: Para): PMNode {
  const marks: Mark[] = [];
  if (p.bold) marks.push(m.strong.create());
  if (p.italic) marks.push(m.em.create());
  if (p.size) marks.push(m.textStyle.create({ size: p.size }));
  const text = p.caps ? (p.text ?? '').toUpperCase() : p.text ?? '';
  return n.paragraph.create({ align: p.align ?? null, indent: p.indent ?? null, first: p.first ?? null, before: p.before ?? null, after: p.after ?? null }, text ? schema.text(text, marks) : undefined);
}

export const lines = (s: string | undefined): string[] => (s ?? '').split(/\r?\n/).map((x) => x.trim()).filter(Boolean);

/** Основной текст: абзацы с отступом 1,25 см по ширине. */
export const body = (text: string | undefined): PMNode[] => lines(text).map((t) => para({ text: t, align: 'justify', first: INDENT }));

/** Строка из двух столбцов без границ: слева одно, справа другое (дата и номер; должность и подпись). */
export function twoCols(left: string, right: string, rightBold = false): PMNode {
  const cell = (t: string, align: 'left' | 'right', bold = false) => n.table_cell.create({ colwidth: [340] }, para({ text: t, align, bold }));
  return n.table.create({ borders: false }, [n.table_row.create(null, [cell(left, 'left'), cell(right, 'right', rightBold)])]);
}

export const dmy = (iso: string): string => {
  const [y, mo, d] = iso.split('-');
  return y && mo && d ? `${d}.${mo}.${y}` : '___.___.______';
};

export interface BuildOptions { values: Record<string, string>; numberLabel: string | null }
export interface Built { doc: PMNode; meta: DocMeta; name: string }

/** Шапка: наименование организации по центру, вид документа, строка «дата — номер», место составления. */
function head(spec: FormSpec, v: Record<string, string>, no: string, withKind: boolean): PMNode[] {
  const out: PMNode[] = [];
  for (const l of lines(v.org)) out.push(para({ text: l, align: 'center', bold: true, caps: false }));
  if (withKind) out.push(para({ text: spec.title, align: 'center', bold: true, caps: true, before: 12, after: 6 }));
  out.push(twoCols(dmy(v.date), `№ ${no}`));
  if (v.city) out.push(para({ text: v.city, align: 'center', after: 6 }));
  return out;
}

function sign(v: Record<string, string>): PMNode[] {
  if (!v.position && !v.signer) return [];
  return [para({ text: '', after: 6 }), twoCols(v.position ?? '', v.signer ?? '', false)];
}

export function buildForm(type: FormType, opt: BuildOptions): Built {
  const spec = FORMS.find((f) => f.type === type)!;
  const v = opt.values;
  const no = opt.numberLabel ?? '___';
  const blocks: PMNode[] = [];

  if (type === 'letter') {
    blocks.push(...head(spec, v, no, false));
    if (v.reply) blocks.push(para({ text: v.reply, size: 12, before: 6 }));
    // Адресат — в правой части поля (отступ слева около 9 см).
    for (const l of lines(v.addressee)) blocks.push(para({ text: l, indent: 255 }));
    if (v.subject) blocks.push(para({ text: v.subject, bold: true, before: 18, after: 12, align: 'left' }));
    blocks.push(...body(v.body));
    for (const [i, l] of lines(v.attach).entries()) blocks.push(para({ text: i === 0 ? `Приложение: ${l}` : l, before: i === 0 ? 12 : 0 }));
    blocks.push(...sign(v));
    if (v.executor) blocks.push(para({ text: v.executor, size: 10, before: 24 }));
  } else if (type === 'order') {
    blocks.push(...head(spec, v, no, true));
    if (v.subject) blocks.push(para({ text: v.subject, bold: true, before: 12, after: 12, align: 'left' }));
    blocks.push(...body(v.preamble));
    blocks.push(para({ text: 'ПРИКАЗЫВАЮ:', bold: true, before: 12, after: 6 }));
    lines(v.items).forEach((l, i) => blocks.push(para({ text: `${i + 1}. ${l}`, align: 'justify', first: INDENT })));
    if (v.control) blocks.push(para({ text: `${lines(v.items).length + 1}. ${v.control}`, align: 'justify', first: INDENT }));
    blocks.push(...sign(v));
  } else if (type === 'act') {
    blocks.push(...head(spec, v, no, true));
    if (v.subject) blocks.push(para({ text: v.subject, bold: true, before: 12, after: 12, align: 'left' }));
    if (v.basis) blocks.push(para({ text: v.basis, align: 'justify', first: INDENT }));
    if (lines(v.commission).length) {
      blocks.push(para({ text: 'Составлен комиссией в составе:', before: 6, first: INDENT }));
      for (const l of lines(v.commission)) blocks.push(para({ text: l, indent: INDENT }));
    }
    blocks.push(...body(v.body));
    blocks.push(...sign(v));
    blocks.push(para({ text: 'Члены комиссии:', before: 12 }));
    blocks.push(para({ text: '_______________ / _______________', before: 6 }));
  } else {
    blocks.push(...head(spec, v, no, true));
    blocks.push(para({ text: v.meeting ?? '', align: 'center', after: 12 }));
    if (v.chair) blocks.push(para({ text: `Председатель — ${v.chair}` }));
    if (v.secretary) blocks.push(para({ text: `Секретарь — ${v.secretary}` }));
    if (lines(v.present).length) {
      blocks.push(para({ text: 'Присутствовали:', before: 6 }));
      for (const l of lines(v.present)) blocks.push(para({ text: l, indent: INDENT }));
    }
    const agenda = lines(v.agenda);
    blocks.push(para({ text: 'ПОВЕСТКА ДНЯ:', bold: true, before: 12, after: 6 }));
    agenda.forEach((l, i) => blocks.push(para({ text: `${i + 1}. ${l}`, indent: INDENT })));
    agenda.forEach((l, i) => {
      blocks.push(para({ text: `${i + 1}. СЛУШАЛИ: ${l}.`, bold: false, align: 'justify', first: INDENT, before: 12 }));
      blocks.push(para({ text: 'РЕШИЛИ:', align: 'justify', first: INDENT }));
    });
    blocks.push(para({ text: '', after: 6 }));
    blocks.push(twoCols(`Председатель${v.chair ? ` ______ ${v.chair}` : ''}`, `Секретарь${v.secretary ? ` ______ ${v.secretary}` : ''}`));
  }

  const meta = newMeta();
  meta.page = { w: 595.3, h: 841.9, top: 56.7, bottom: 56.7, left: 56.7, right: 28.35 };
  meta.defaults = { font: 'Times New Roman', size: 14, line: 1, before: 0, after: 0 };
  // Номер страницы вверху по центру, со второй страницы (ГОСТ Р 7.0.97, п. 6.3).
  meta.hf.header = n.doc.create(null, [n.paragraph.create({ align: 'center' }, n.field.create({ kind: 'page' }))]).toJSON();
  meta.hf.differentFirst = true;
  const stamp = dmy(v.date ?? '');
  return { doc: n.doc.create(null, blocks), meta, name: `${spec.title} ${no === '___' ? stamp : '№ ' + no.replace(/\//g, '-')}` };
}
