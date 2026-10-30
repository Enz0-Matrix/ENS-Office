// Читаемость по-русски без интернета: время чтения, длинные предложения, повторы слов, канцеляризмы, индекс Флеша (адаптация Оборневой).

import type { Node as PMNode } from 'prosemirror-model';

export interface Issue { kind: 'long' | 'repeat' | 'clerical'; text: string; note: string; pos: number; len: number }
export interface Report {
  words: number;
  sentences: number;
  /** Минуты чтения про себя, около 180 слов в минуту. */
  minutes: number;
  avgSentence: number;
  longest: number;
  /** Индекс удобочитаемости (чем выше, тем легче). */
  flesch: number;
  level: string;
  issues: Issue[];
}

const WORD = /[A-Za-zА-Яа-яЁё]+(?:-[A-Za-zА-Яа-яЁё]+)*/g;
const VOWELS = /[аеёиоуыэюяaeiouy]/gi;
/** Сокращения, после точки которых предложение не кончается. */
const ABBR = /(?:^|[\s(«"])(?:г|гг|т|д|п|пп|ул|д|стр|рис|табл|руб|коп|тыс|млн|млрд|им|см|напр|др|пр|проф|акад|доц|ст|ч|ф|кв|корп|обл|р|пос|с|к|н|в|вв)\.(?=\s|$)/gi;

/** Клише и канцеляризмы: регулярное выражение, подсказка. */
const CLERICAL: [RegExp, string][] = [
  [/осуществля(?:ть|ет|ется|ют)|осуществлени(?:е|я|ю|ем|и)\b/i, 'проще: «делать», «проводить»'],
  [/производ(?:ить|ит|ится)\b/i, 'проще: «делать», «выполнять»'],
  [/в\s+целях\b/i, 'проще: «чтобы», «для»'],
  [/в\s+связи\s+с\s+тем,?\s+что/i, 'проще: «потому что», «так как»'],
  [/на\s+сегодняшний\s+день/i, 'проще: «сейчас», «сегодня»'],
  [/имеет(?:ся)?\s+место/i, 'проще: «есть», «бывает»'],
  [/в\s+случае\s+необходимости/i, 'проще: «если нужно»'],
  [/в\s+рамках\b/i, 'проще: «при», «в»'],
  [/оказ(?:ывать|ание|ать)\s+содействи/i, 'проще: «помогать»'],
  [/принима(?:ть|ет)\s+участие/i, 'проще: «участвовать»'],
  [/вышеуказанн|вышеперечисленн|нижеследующ|вышеизложенн/i, 'проще сослаться на пункт или назвать предмет'],
  [/\bданн(?:ый|ая|ое|ые|ого|ому|ым|ых|ой)\b/i, 'часто лишнее: «этот», «такой» или можно опустить'],
  [/в\s+адрес\b/i, 'проще: «кому», «для»'],
  [/по\s+причине\b/i, 'проще: «из-за», «потому что»'],
  [/с\s+целью\b/i, 'проще: «чтобы»'],
  [/является\b/i, 'проще: «это», «есть» (если не теряется смысл)'],
  [/путём\b/i, 'проще: «с помощью», «через»'],
  [/в\s+течение\s+(?:времени|срока)/i, 'проще: «за», «в течение» с точным сроком'],
  [/представля(?:ет|ют)\s+собой/i, 'проще: «это»'],
];

/** Слова, повторы которых не считаем. */
const STOP = new Set(['который', 'которая', 'которые', 'которых', 'которого', 'также', 'этого', 'этому', 'этой', 'этих', 'более', 'после', 'перед', 'через', 'между', 'только', 'чтобы', 'когда', 'если', 'тоже', 'очень', 'можно', 'нужно', 'будет', 'были', 'было', 'есть', 'или', 'при', 'для', 'что', 'как', 'так', 'его', 'она', 'они', 'оно', 'все', 'всё', 'еще', 'ещё', 'уже', 'там', 'тут', 'где', 'без', 'под', 'над', 'про', 'вот', 'это', 'этот', 'эта', 'эти']);

const syllables = (w: string): number => Math.max(1, (w.match(VOWELS) ?? []).length);

function level(f: number): string {
  if (f >= 80) return 'очень легко читается';
  if (f >= 60) return 'легко читается';
  if (f >= 40) return 'средняя сложность — деловой текст';
  if (f >= 20) return 'трудно читается — канцелярский стиль';
  return 'очень трудно читается';
}

const stemOf = (w: string): string => { const l = w.toLowerCase().replace(/ё/g, 'е'); return l.length >= 7 ? l.slice(0, 5) : l.length >= 5 ? l.slice(0, 4) : l; };

/** Анализ текста документа (или его части: from, to — границы в позициях). */
export function analyze(doc: PMNode, from = 0, to = doc.content.size): Report {
  let words = 0, syl = 0, sentences = 0, longest = 0, sentWordsTotal = 0;
  const issues: Issue[] = [];
  const stream: { stem: string; word: string; pos: number }[] = [];

  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return true;
    let text = '';
    node.forEach((c) => { text += c.isText ? c.text : '￼'; });
    const start = pos + 1;
    // Защищаем точки сокращений и инициалов (та же длина строки, чтобы смещения не поехали).
    const safe = text.replace(ABBR, (m) => m.replace('.', '\u0001')).replace(/(?<=\b[А-ЯЁA-Z])\.(?=\s*[А-ЯЁA-Z]\.|\s+[А-ЯЁ][а-яё])/g, '\u0001');
    const re = /[^.!?…]+(?:[.!?…]+|$)/g;
    for (let m = re.exec(safe); m; m = re.exec(safe)) {
      const s = m[0];
      const ws = s.match(WORD) ?? [];
      if (!ws.length) continue;
      sentences++;
      sentWordsTotal += ws.length;
      longest = Math.max(longest, ws.length);
      if (ws.length > 30) issues.push({ kind: 'long', text: s.trim().slice(0, 90) + (s.trim().length > 90 ? '…' : ''), note: `${ws.length} слов в одном предложении — лучше разбить`, pos: start + m.index + (s.length - s.trimStart().length), len: Math.min(s.trim().length, 1) });
    }
    for (const w of text.matchAll(WORD)) {
      words++;
      syl += syllables(w[0]);
      if (w[0].length >= 4 && !STOP.has(w[0].toLowerCase())) stream.push({ stem: stemOf(w[0]), word: w[0], pos: start + (w.index ?? 0) });
    }
    for (const [re2, note] of CLERICAL) {
      // \\b в JS не знает кириллицу: границы слова задаём через проверку соседних букв.
      const g = new RegExp('(?<![А-Яа-яЁё])' + re2.source.replace(/^\\b/, '').replace(/\\b/g, '(?![А-Яа-яЁё])'), 'gi');
      for (let m = g.exec(text); m; m = g.exec(text)) issues.push({ kind: 'clerical', text: m[0], note, pos: start + m.index, len: m[0].length });
    }
    return false;
  });

  // Повтор: одна основа три раза и больше в окне из 40 слов (считаются слова ≥ 4 букв).
  const seen = new Set<string>();
  for (let i = 0; i < stream.length; i++) {
    const s = stream[i];
    if (seen.has(s.stem)) continue;
    let n = 1;
    for (let j = i + 1; j < stream.length && j < i + 40; j++) if (stream[j].stem === s.stem) n++;
    if (n >= 3) { seen.add(s.stem); issues.push({ kind: 'repeat', text: s.word, note: `повторяется ${n} ${n % 10 >= 2 && n % 10 <= 4 && (n % 100 < 10 || n % 100 >= 20) ? 'раза' : 'раз'} в пределах 40 слов`, pos: s.pos, len: s.word.length }); }
  }

  const avgSentence = sentences ? sentWordsTotal / sentences : 0;
  const asw = words ? syl / words : 0;
  // Индекс Флеша — Оборневой для русского языка: 206,835 − 1,3·ASL − 60,1·ASW.
  const flesch = words ? Math.round((206.835 - 1.3 * avgSentence - 60.1 * asw) * 10) / 10 : 0;
  issues.sort((a, b) => a.pos - b.pos);
  return { words, sentences, minutes: words / 180, avgSentence: Math.round(avgSentence * 10) / 10, longest, flesch, level: words ? level(flesch) : '', issues };
}

export const minutesText = (m: number): string => {
  if (m < 0.5) return 'меньше минуты';
  const r = Math.round(m);
  return `около ${r} ${r % 10 === 1 && r % 100 !== 11 ? 'минуты' : 'минут'}`.replace('около 1 минуты', 'около минуты');
};
