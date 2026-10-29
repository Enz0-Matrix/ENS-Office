// Типограф: приводит набранный текст к правилам русской типографики — кавычки «ёлочки» и „лапки“, тире, неразрывные
// пробелы, лишние пробелы, многоточие. Работает по выделению или по всему документу, одним шагом отмены.
// Все правки — обычные шаги ProseMirror: форматирование (жирный, шрифт…) сохраняется, ссылки и числа в адресах не трогаются.

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, type EditorState, type Transaction } from 'prosemirror-state';
import { schema } from './schema';

const NBSP = ' ';

export interface Edit { from: number; to: number; text: string }

/** Правило: возвращает правки по тексту абзаца (позиции — внутри абзаца, от 0). */
interface Rule {
  id: string;
  name: string;
  find(text: string): { from: number; to: number; text: string }[];
}

const OPEN_PREV = /^$|[\s (\[{«„—–\-/]/;

/** Кавычки: открывающая, если перед ней начало, пробел, скобка или тире; внутри «ёлочек» — „лапки“. */
const quotes: Rule = {
  id: 'quotes',
  name: 'кавычки',
  find(text) {
    const out: Edit[] = [];
    let depth = 0;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (!'"«»„“”'.includes(ch)) continue;
      const prev = i === 0 ? '' : text[i - 1];
      let open: boolean;
      if (ch === '«' || ch === '„') open = true;
      else if (ch === '»' || ch === '”') open = false;
      else if (ch === '“') open = depth < 2; // “ закрывает внутреннюю „лапку“, а в одиночку — английская открывающая
      else open = prev === '' || /[\s\u00A0(\[{—–-]/.test(prev); // прямая " после буквы или знака — закрывающая
      let want: string;
      if (open) { depth++; want = depth === 1 ? '«' : '„'; } else { want = depth > 1 ? '“' : '»'; depth = Math.max(0, depth - 1); }
      if (want !== ch) out.push({ from: i, to: i + 1, text: want });
    }
    return out;
  },
};

/** Тире: « - » и « -- » между словами — длинное тире с неразрывным пробелом перед ним. */
const dashes: Rule = {
  id: 'dashes',
  name: 'тире',
  find(text) {
    const out: Edit[] = [];
    for (const m of text.matchAll(/(\S)[  ]+(?:--?|–)[  ]+(?=\S)/g)) {
      const from = m.index! + 1;
      const to = m.index! + m[0].length;
      const want = `${NBSP}— `;
      if (text.slice(from, to) !== want) out.push({ from, to, text: want });
    }
    return out;
  },
};

const ELLIPSIS: Rule = {
  id: 'ellipsis',
  name: 'многоточие',
  find: (text) => [...text.matchAll(/(?<!\.)\.{3}(?!\.)/g)].map((m) => ({ from: m.index!, to: m.index! + 3, text: '…' })),
};

/** Слова, после которых пробел неразрывный: предлоги и союзы из одной-двух букв. */
const SHORT = 'а|в|во|до|за|и|из|к|ко|на|не|ни|но|о|об|от|по|с|со|у|я|при|над|под|про|для|без';

/** Неразрывные пробелы: после коротких предлогов, в «г. Москва», «№ 5», «5 кг», инициалах, «и т. д.», датах, числах. */
const nbsp: Rule = {
  id: 'nbsp',
  name: 'неразрывные пробелы',
  find(text) {
    const out: Edit[] = [];
    const push = (from: number, to: number) => { if (text.slice(from, to) !== NBSP) out.push({ from, to, text: NBSP }); };
    // Отдельное короткое слово + пробел.
    for (const m of text.matchAll(new RegExp(`(?<![\\p{L}\\p{N}])(?:${SHORT})(?![\\p{L}\\p{N}])( )(?=[\\p{L}\\p{N}«„(])`, 'giu'))) push(m.index! + m[0].length - 1, m.index! + m[0].length);
    // Сокращения с точкой + пробел + слово/число: г., ул., д., кв., им., тел., стр., рис., табл., см., ст., п., пп., руб., коп., тыс., млн., млрд.
    for (const m of text.matchAll(/(?<![\p{L}])(?:г|гг|ул|д|кв|к|корп|им|тел|стр|рис|табл|см|ст|п|пп|пр|пер|обл|р-н|тыс|млн|млрд|руб|коп|ч|мин|сек)\.( )(?=[\p{L}\p{N}№])/giu)) push(m.index! + m[0].length - 1, m.index! + m[0].length);
    // № и § + пробел + число.
    for (const m of text.matchAll(/([№§])( )(?=\d)/g)) push(m.index! + 1, m.index! + 2);
    // Число + единица измерения / валюта / процент / «г.» / «года».
    for (const m of text.matchAll(/(\d)( )(?=(?:кг|г|т|мг|м|км|см|мм|л|мл|ч|мин|сек|с|руб|коп|тыс|млн|млрд|%|шт|чел|га|кВт|МВт|ГБ|МБ|КБ|Гц|°C|₽|\$|€)(?![\p{L}]))/gu)) push(m.index! + 1, m.index! + 2);
    // Инициалы: «А. С. Пушкин», «Пушкин А. С.».
    for (const m of text.matchAll(/(?<![\p{L}])\p{Lu}\.( )(?=\p{Lu}\.?)/gu)) push(m.index! + 2, m.index! + 3);
    for (const m of text.matchAll(/(?<![\p{L}])\p{Lu}\.( )(?=\p{Lu}\p{Ll}{2,})/gu)) push(m.index! + 2, m.index! + 3);
    // «и т. д.», «и т. п.», «и др.», «т. е.», «т. к.»
    for (const m of text.matchAll(/(?<![\p{L}])(т\.)( )(?=[дпеко]\.)/giu)) push(m.index! + 2, m.index! + 3);
    for (const m of text.matchAll(/(?<![\p{L}])(и)( )(?=(?:т\.|др\.|пр\.))/giu)) push(m.index! + 1, m.index! + 2);
    // Числа с разрядами: «1 000 000».
    for (const m of text.matchAll(/(?<![\d.,])(\d{1,3})((?: \d{3})+)(?![\d])/g)) {
      const base = m.index! + m[1].length;
      for (let i = 0; i < m[2].length; i += 4) push(base + i, base + i + 1);
    }
    // Дата: «5 мая 2026 г.» — число+месяц, год+«г.».
    for (const m of text.matchAll(/(\d{1,2})( )(?=(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)(?![\p{L}]))/giu)) push(m.index! + m[1].length, m.index! + m[1].length + 1);
    for (const m of text.matchAll(/(\d{4})( )(?=(?:г\.|года|году|гг\.)(?![\p{L}]))/gu)) push(m.index! + 4, m.index! + 5);
    return out;
  },
};

/** Лишние пробелы: двойные, перед знаками препинания, в начале и в конце абзаца; пробел после запятой, точки с запятой, двоеточия. */
const spaces: Rule = {
  id: 'spaces',
  name: 'лишние пробелы',
  find(text) {
    const out: Edit[] = [];
    const skip = (a: number, b: number) => /https?:\/\/|www\.|@|\w\.\w{2,4}\//.test(text.slice(Math.max(0, a - 12), b + 12));
    for (const m of text.matchAll(/ {2,}/g)) if (m.index! > 0 && m.index! + m[0].length < text.length) out.push({ from: m.index!, to: m.index! + m[0].length, text: ' ' });
    // Пробелы в начале абзаца: ручной отступ — убираем (отступ первой строки задаётся форматом абзаца).
    const lead = /^[ \t ]+(?=\S)/.exec(text);
    if (lead) out.push({ from: 0, to: lead[0].length, text: '' });
    const tail = /[ \t]+$/.exec(text);
    if (tail && tail.index > 0) out.push({ from: tail.index, to: text.length, text: '' });
    for (const m of text.matchAll(/[  ]+(?=[,.;:!?)»])/g)) if (!skip(m.index!, m.index!)) out.push({ from: m.index!, to: m.index! + m[0].length, text: '' });
    for (const m of text.matchAll(/(?<=[(«])[  ]+/g)) out.push({ from: m.index!, to: m.index! + m[0].length, text: '' });
    // После запятой, ; и : — пробел, если следом буква (не число после запятой: 3,5 и не адрес).
    for (const m of text.matchAll(/([\p{L}\p{N}»)]),(?=[\p{L}«(])/gu)) if (!skip(m.index!, m.index!)) out.push({ from: m.index! + 2, to: m.index! + 2, text: ' ' });
    for (const m of text.matchAll(/([\p{L}»)])([;:])(?=[\p{L}«(])/gu)) if (!skip(m.index!, m.index!) && !/https?|mailto/i.test(text.slice(Math.max(0, m.index! - 6), m.index! + 1))) out.push({ from: m.index! + 2, to: m.index! + 2, text: ' ' });
    // После точки с заглавной буквы («слово.Слово»).
    for (const m of text.matchAll(/(\p{Ll}{2,})\.(?=\p{Lu}\p{Ll})/gu)) if (!skip(m.index!, m.index!)) out.push({ from: m.index! + m[0].length, to: m.index! + m[0].length, text: ' ' });
    return out;
  },
};

export const RULES: Rule[] = [spaces, quotes, dashes, ELLIPSIS, nbsp];

export interface Report { counts: Record<string, number>; total: number }

/** Текст абзаца с позициями: каждый «лист» (картинка, поле) — один символ-заглушка, чтобы смещения совпали с позициями. */
function textOf(block: PMNode): string {
  return block.textBetween(0, block.content.size, undefined, '￼');
}

/** Применяет типограф к документу или выделению одной транзакцией. Возвращает число правок по правилам. */
export function typograph(state: EditorState, only?: string[]): { tr: Transaction; report: Report } {
  const tr = state.tr;
  const counts: Record<string, number> = {};
  const { from, to, empty } = state.selection;
  const range: [number, number] = empty ? [0, state.doc.content.size] : [from, to];
  // Проходим дважды: исправление пробелов открывает место для следующих правил («слово ,слово» → «слово, слово»).
  for (const rule of [...RULES, ...RULES]) {
    if (only && !only.includes(rule.id)) continue;
    // Правки считаем по текущему состоянию документа (после предыдущих правил) и применяем с конца.
    const edits: Edit[] = [];
    const r0 = tr.mapping.map(range[0], -1);
    const r1 = tr.mapping.map(range[1], 1);
    tr.doc.nodesBetween(Math.max(0, r0), Math.min(tr.doc.content.size, r1), (node, pos) => {
      if (!node.isTextblock) return true;
      if (node.type === schema.nodes.heading || node.type === schema.nodes.paragraph) {
        const text = textOf(node);
        for (const e of rule.find(text)) {
          const a = pos + 1 + e.from;
          const b = pos + 1 + e.to;
          // Ссылки не трогаем: адрес должен остаться как есть.
          if (b > a && tr.doc.rangeHasMark(a, b, schema.marks.link)) continue;
          edits.push({ from: a, to: b, text: e.text });
        }
      }
      return false;
    });
    if (!edits.length) continue;
    // Правки одного правила не должны пересекаться: при совпадении места удаление важнее замены.
    edits.sort((x, y) => x.from - y.from || x.to - y.to || (x.text ? 1 : 0) - (y.text ? 1 : 0));
    const clean: Edit[] = [];
    for (const e of edits) {
      const last = clean[clean.length - 1];
      if (last && e.from < last.to) continue;
      if (last && e.from === last.from && e.to === last.to) continue;
      clean.push(e);
    }
    edits.length = 0;
    edits.push(...clean);
    counts[rule.id] = (counts[rule.id] ?? 0) + edits.length;
    edits.sort((x, y) => y.from - x.from);
    for (const e of edits) {
      if (e.from === e.to) tr.insertText(e.text, e.from);
      else if (!e.text) tr.delete(e.from, e.to);
      else {
        const marks = tr.doc.nodeAt(e.from)?.marks ?? [];
        tr.replaceWith(e.from, e.to, schema.text(e.text, marks));
      }
    }
  }
  const total = Object.values(counts).reduce((a, b) => a + b, 0);
  return { tr, report: { counts, total } };
}

export function describe(r: Report): string {
  if (!r.total) return 'Типограф: исправлять нечего.';
  const names = Object.fromEntries(RULES.map((x) => [x.id, x.name]));
  return `Типограф: ${Object.entries(r.counts).map(([k, v]) => `${names[k]} — ${v}`).join(', ')}. Отменить можно Ctrl+Z.`;
}

/** Кавычки на лету: набрали « " » — получилась «ёлочка» или „лапка“ по контексту. */
export const liveQuotes = (enabled: () => boolean): Plugin => new Plugin({
  props: {
    handleTextInput(view, from, to, text) {
      if (!enabled() || text !== '"') return false;
      const $f = view.state.doc.resolve(from);
      const before = $f.parent.textBetween(0, $f.parentOffset, undefined, '￼');
      let depth = 0;
      for (const ch of before) { if (ch === '«' || ch === '„') depth++; else if (ch === '»' || ch === '“') depth = Math.max(0, depth - 1); }
      const prev = before.slice(-1);
      const opens = prev === '' || /[\s (\[{—–-]/.test(prev);
      const q = opens ? (depth === 0 ? '«' : '„') : (depth > 1 ? '“' : '»');
      view.dispatch(view.state.tr.insertText(q, from, to));
      return true;
    },
  },
});
