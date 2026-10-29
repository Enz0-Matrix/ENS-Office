// Русский язык без интернета: сумма прописью, ФИО и должности по падежам, дата «по-канцелярски».
// Склонение — по правилам и окончаниям, без словаря: для редких фамилий результат стоит проверить глазами.

// ---------------------------------------------------------------- сумма прописью

export type Gender = 'm' | 'f' | 'n';

const ONES: Record<Gender, string[]> = {
  m: ['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять'],
  f: ['', 'одна', 'две', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять'],
  n: ['', 'одно', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять'],
};
const TEENS = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать', 'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать'];
const TENS = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто'];
const HUNDREDS = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот', 'семьсот', 'восемьсот', 'девятьсот'];
/** Разряды: форма для 1 / 2–4 / 5+ и род. */
const SCALES: { f: [string, string, string]; g: Gender }[] = [
  { f: ['', '', ''], g: 'm' },
  { f: ['тысяча', 'тысячи', 'тысяч'], g: 'f' },
  { f: ['миллион', 'миллиона', 'миллионов'], g: 'm' },
  { f: ['миллиард', 'миллиарда', 'миллиардов'], g: 'm' },
  { f: ['триллион', 'триллиона', 'триллионов'], g: 'm' },
];

/** Форма слова по числу: 1 рубль, 2 рубля, 5 рублей, 11 рублей, 21 рубль. */
export function plural(n: bigint | number, forms: [string, string, string]): string {
  const v = Number((typeof n === 'bigint' ? n : BigInt(Math.trunc(n))) % 100n);
  const d = v % 10;
  if (v >= 11 && v <= 14) return forms[2];
  if (d === 1) return forms[0];
  if (d >= 2 && d <= 4) return forms[1];
  return forms[2];
}

function triple(n: number, g: Gender): string[] {
  const out: string[] = [];
  const h = Math.floor(n / 100), t = Math.floor((n % 100) / 10), o = n % 10;
  if (h) out.push(HUNDREDS[h]);
  if (t === 1) out.push(TEENS[o]);
  else {
    if (t) out.push(TENS[t]);
    if (o) out.push(ONES[g][o]);
  }
  return out;
}

/** Целое число словами; род — для «одна тысяча»/«одно место»/«один рубль». */
export function numberWords(value: bigint | number, gender: Gender = 'm'): string {
  let n = typeof value === 'bigint' ? value : BigInt(Math.trunc(value));
  if (n === 0n) return 'ноль';
  const neg = n < 0n;
  if (neg) n = -n;
  const parts: string[] = [];
  for (let i = 0; n > 0n; i++) {
    const chunk = Number(n % 1000n);
    n /= 1000n;
    if (i >= SCALES.length) throw new Error('Число слишком большое');
    if (!chunk) continue;
    const sc = SCALES[i];
    parts.unshift(...triple(chunk, i === 0 ? gender : sc.g), ...(i ? [plural(chunk, sc.f)] : []));
  }
  return (neg ? 'минус ' : '') + parts.join(' ');
}

export interface Currency { major: [string, string, string]; minor: [string, string, string]; majorGender: Gender; minorGender: Gender; sign: string }
export const CURRENCIES: Record<string, Currency> = {
  RUB: { major: ['рубль', 'рубля', 'рублей'], minor: ['копейка', 'копейки', 'копеек'], majorGender: 'm', minorGender: 'f', sign: '₽' },
  USD: { major: ['доллар США', 'доллара США', 'долларов США'], minor: ['цент', 'цента', 'центов'], majorGender: 'm', minorGender: 'm', sign: '$' },
  EUR: { major: ['евро', 'евро', 'евро'], minor: ['цент', 'цента', 'центов'], majorGender: 'n', minorGender: 'm', sign: '€' },
  KZT: { major: ['тенге', 'тенге', 'тенге'], minor: ['тиын', 'тиына', 'тиын'], majorGender: 'm', minorGender: 'm', sign: '₸' },
  BYN: { major: ['белорусский рубль', 'белорусских рубля', 'белорусских рублей'], minor: ['копейка', 'копейки', 'копеек'], majorGender: 'm', minorGender: 'f', sign: 'Br' },
};

/** Разбор «12 345,67»: целая часть и копейки (округление до сотых). */
export function parseAmount(input: string | number): { int: bigint; frac: number; neg: boolean } {
  let s = String(input).trim().replace(/[\s ]/g, '').replace(',', '.');
  if (!/^[-+]?\d*\.?\d*$/.test(s) || !/\d/.test(s)) throw new Error('Не число');
  const neg = s.startsWith('-');
  s = s.replace(/^[-+]/, '');
  let [i = '0', f = ''] = s.split('.');
  if (i === '') i = '0';
  let int = BigInt(i);
  let frac = parseInt((f + '00').slice(0, 2), 10);
  if (f.length > 2 && f[2] >= '5') frac++;
  if (frac === 100) { frac = 0; int += 1n; }
  return { int, frac, neg };
}

export interface SumOptions {
  currency?: string;
  /** 'words' — только слова; 'both' — «12 345 (двенадцать тысяч …) рублей 67 копеек». */
  style?: 'words' | 'both';
  /** Копейки цифрами (по умолчанию) или словами. */
  kopWords?: boolean;
  capital?: boolean;
}

const spaced = (n: bigint) => n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');

/** «12 345,67» → «двенадцать тысяч триста сорок пять рублей 67 копеек». */
export function sumWords(input: string | number, opt: SumOptions = {}): string {
  const cur = CURRENCIES[opt.currency ?? 'RUB'];
  if (!cur) throw new Error('Неизвестная валюта');
  const { int, frac, neg } = parseAmount(input);
  const major = numberWords(int, cur.majorGender);
  const kop = opt.kopWords ? numberWords(frac, cur.minorGender) : String(frac).padStart(2, '0');
  const tail = `${plural(int, cur.major)} ${kop} ${plural(frac, cur.minor)}`;
  let text = (neg ? 'минус ' : '') + (opt.style === 'both' ? `${spaced(int)} (${major}) ${tail}` : `${major} ${tail}`);
  if (opt.capital !== false) text = text[0].toUpperCase() + text.slice(1);
  return text;
}

// ---------------------------------------------------------------- дата

const MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];

/** «29» октября 2026 г. */
export function dateRu(d: Date, quoted = true): string {
  const day = String(d.getDate()).padStart(2, '0');
  return `${quoted ? `«${day}»` : String(d.getDate())} ${MONTHS_GEN[d.getMonth()]} ${d.getFullYear()} г.`;
}

// ---------------------------------------------------------------- падежи

export const CASES = ['nom', 'gen', 'dat', 'acc', 'ins', 'prep'] as const;
export type Case = typeof CASES[number];
export const CASE_LABELS: Record<Case, string> = { nom: 'Кто? (именительный)', gen: 'Кого? (родительный)', dat: 'Кому? (дательный)', acc: 'Кого? (винительный)', ins: 'Кем? (творительный)', prep: 'О ком? (предложный)' };

const idx = (c: Case) => CASES.indexOf(c);
const VOWELS = 'аеёиоуыэюя';
const isVowel = (ch: string) => VOWELS.includes(ch);
const HISS = 'жшчщ';
const VELAR = 'гкх';

/** Окончания: [nom, gen, dat, acc, ins, prep] */
type Forms = [string, string, string, string, string, string];

/** Сохранить регистр первой буквы и «ВЕРХНИЙ РЕГИСТР». */
function keepCase(src: string, out: string): string {
  if (src.length > 1 && src === src.toUpperCase() && /[А-ЯЁ]/.test(src)) return out.toUpperCase();
  return out;
}

/** Склонение существительного/имени по последним буквам основы. `masc` — мужской род (для -а/-я и -ь). */
function nounForms(word: string, masc: boolean): Forms {
  const w = word;
  const last = w.slice(-1).toLowerCase();
  const prev = w.slice(-2, -1).toLowerCase();
  const stem = w.slice(0, -1);
  if (last === 'а') {
    const gen = HISS.includes(prev) || VELAR.includes(prev) ? 'и' : 'ы';
    const ins = HISS.includes(prev) || prev === 'ц' ? 'ей' : 'ой';
    return [w, stem + gen, stem + 'е', stem + 'у', stem + ins, stem + 'е'];
  }
  if (last === 'я') {
    if (prev === 'и') return [w, stem + 'и', stem + 'и', stem + 'ю', stem + 'ей', stem + 'и'];
    if (prev === 'ь' || isVowel(prev)) return [w, stem + 'и', stem + 'е', stem + 'ю', stem + 'ей', stem + 'е'];
    return [w, stem + 'и', stem + 'е', stem + 'ю', stem + 'ей', stem + 'е'];
  }
  if (last === 'ь') {
    if (masc) return [w, stem + 'я', stem + 'ю', stem + 'я', stem + 'ем', stem + 'е'];
    return [w, stem + 'и', stem + 'и', w, stem + 'ью', stem + 'и'];
  }
  if (last === 'й') {
    if (prev === 'и') return [w, stem + 'я', stem + 'ю', stem + 'я', stem + 'ем', stem + 'и'];
    return [w, stem + 'я', stem + 'ю', stem + 'я', stem + 'ем', stem + 'е'];
  }
  if (isVowel(last)) return [w, w, w, w, w, w];
  // согласная
  const ins = HISS.includes(last) || last === 'ц' ? 'ем' : 'ом';
  return [w, w + 'а', w + 'у', w + 'а', w + ins, w + 'е'];
}

export type NameGender = 'm' | 'f';

const MALE_A_NAMES = new Set(['никита', 'илья', 'лёва', 'лева', 'данила', 'савва', 'фома', 'кузьма', 'гаврила', 'жора', 'миша', 'гоша', 'коля', 'вася', 'петя', 'толя', 'витя', 'дима', 'ваня', 'саша', 'женя', 'валя', 'слава']);
const UNISEX = new Set(['саша', 'женя', 'валя', 'слава']);
const FEMALE_SOFT = new Set(['любовь']);

export function guessGender(parts: { last?: string; first?: string; middle?: string }): NameGender {
  const mid = (parts.middle ?? '').toLowerCase();
  if (/(вна|чна|ична)$/.test(mid)) return 'f';
  if (/(вич|ьич|ич)$/.test(mid)) return 'm';
  const last = (parts.last ?? '').toLowerCase();
  if (/(ова|ева|ёва|ина|ына|ская|цкая|ая|яя)$/.test(last)) return 'f';
  if (/(ов|ев|ёв|ин|ын|ский|цкий|ой|ий|ый)$/.test(last)) return 'm';
  const first = (parts.first ?? '').toLowerCase();
  if (FEMALE_SOFT.has(first)) return 'f';
  if (MALE_A_NAMES.has(first) && !UNISEX.has(first)) return 'm';
  if (/[ая]$/.test(first)) return 'f';
  return 'm';
}

const FLEETING: Record<string, string> = { лев: 'льв', павел: 'павл', пётр: 'петр', петр: 'петр' };

function declineFirstName(name: string, gender: NameGender): Forms {
  const low = name.toLowerCase();
  if (gender === 'm' && FLEETING[low]) {
    const st = keepCase(name, name[0] + FLEETING[low].slice(1) === name ? name : (name[0] === name[0].toUpperCase() ? FLEETING[low][0].toUpperCase() : FLEETING[low][0]) + FLEETING[low].slice(1));
    return [name, st + 'а', st + 'у', st + 'а', st + 'ом', st + 'е'];
  }
  if (FEMALE_SOFT.has(low)) return nounForms(name, false);
  if (gender === 'm') {
    if (/[ая]$/.test(low)) {
      const f = nounForms(name, true);
      // Илья: -ьи/-ье/-ью/-ьёй; Никита: -ы/-е/-у/-ой — общий случай ок; Илья/Кузьма с ь:
      if (low.endsWith('ья')) return [name, name.slice(0, -1) + 'и', name.slice(0, -1) + 'е', name.slice(0, -1) + 'ю', name.slice(0, -1) + 'ёй', name.slice(0, -1) + 'е'];
      f[3] = name.slice(0, -1) + 'у';
      return f;
    }
    return nounForms(name, true);
  }
  // женские имена
  if (/ь$/.test(low)) return nounForms(name, false);
  if (!/[аяь]$/.test(low)) return [name, name, name, name, name, name];
  return nounForms(name, false);
}

function declineSurname(name: string, gender: NameGender): Forms {
  const low = name.toLowerCase();
  const same: Forms = [name, name, name, name, name, name];
  if (/(ых|их|ко|аго|яго)$/.test(low) && !/(ская|цкая)$/.test(low)) return same;
  if (gender === 'f') {
    if (/(ова|ева|ёва|ина|ына)$/.test(low)) { const s = name.slice(0, -1); return [name, s + 'ой', s + 'ой', s + 'у', s + 'ой', s + 'ой']; }
    if (/(ая|яя)$/.test(low)) { const s = name.slice(0, -2); const h = low.endsWith('яя') ? ['ей', 'ей', 'юю', 'ей', 'ей'] : ['ой', 'ой', 'ую', 'ой', 'ой']; return [name, s + h[0], s + h[1], s + h[2], s + h[3], s + h[4]]; }
    return same; // женские фамилии на согласную и т. п. не склоняются
  }
  if (/(ский|цкий|ной|ый|ий|ой)$/.test(low)) {
    if (/(ой)$/.test(low)) { const s = name.slice(0, -2); return [name, s + 'ого', s + 'ому', s + 'ого', s + 'ым', s + 'ом']; }
    const s = name.slice(0, -2);
    return [name, s + 'ого', s + 'ому', s + 'ого', /ый$/.test(low) ? s + 'ым' : s + 'им', s + 'ом'];
  }
  if (/(ов|ев|ёв|ин|ын)$/.test(low)) return [name, name + 'а', name + 'у', name + 'а', name + 'ым', name + 'е'];
  if (/ь$/.test(low) || /й$/.test(low)) return nounForms(name, true);
  if (/[ая]$/.test(low)) return nounForms(name, true);
  if (/[аеёиоуыэюя]$/.test(low)) return same;
  return nounForms(name, true);
}

function declinePatronymic(name: string, gender: NameGender): Forms {
  const low = name.toLowerCase();
  if (gender === 'f' && /на$/.test(low)) { const s = name.slice(0, -1); return [name, s + 'ы', s + 'е', s + 'у', s + 'ой', s + 'е']; }
  if (gender === 'm' && /ич$/.test(low)) return [name, name + 'а', name + 'у', name + 'а', name + 'ем', name + 'е'];
  return [name, name, name, name, name, name];
}

export interface DeclinedName { gender: NameGender; forms: Record<Case, string> }

/** Склонить «Фамилия Имя Отчество» (в любом числе частей; порядок «Имя Отчество Фамилия» не угадывается).
 *  `initials` — вернуть «Иванову И. И.» (фамилия полностью, имя и отчество инициалами). */
export function declineFullName(full: string, opt: { gender?: NameGender; initials?: boolean } = {}): DeclinedName {
  const words = full.trim().split(/\s+/).filter(Boolean);
  if (!words.length) throw new Error('Пустое ФИО');
  // «Иванов И. И.» и «И. И. Иванов»: инициалы — без склонения.
  const isInit = (w: string) => /^[А-ЯЁA-Z]\.?(-?[А-ЯЁA-Z]\.?)?$/.test(w);
  const names = words.filter((w) => !isInit(w));
  const isPat = (w?: string) => !!w && /(вич|ьич|вна|чна)$/i.test(w);
  // «Мария Ивановна Белых»: отчество вторым из трёх слов — значит, фамилия в конце.
  const surnameLast = names.length === 3 && isPat(names[1]);
  const surname = surnameLast ? names[2] : names[0];
  const first = surnameLast ? names[0] : names[1];
  const parts = { last: surname, first, middle: surnameLast ? names[1] : names[2] };
  const gender = opt.gender ?? guessGender(parts);
  const out = {} as Record<Case, string>;
  const decl = new Map<string, Forms>();
  words.forEach((w, i) => {
    if (isInit(w)) return;
    const role = w === surname ? 'last' : w === first ? 'first' : 'middle';
    // Двойные имена и фамилии через дефис склоняем по частям.
    const forms = w.includes('-') ? (w.split('-').map((p) => (role === 'last' ? declineSurname(p, gender) : declineFirstName(p, gender)))) : null;
    if (forms) decl.set(w + '#' + i, CASES.map((_, k) => forms.map((f) => f[k]).join('-')) as unknown as Forms);
    else decl.set(w + '#' + i, role === 'last' ? declineSurname(w, gender) : role === 'first' ? declineFirstName(w, gender) : declinePatronymic(w, gender));
  });
  for (const c of CASES) {
    const k = idx(c);
    out[c] = words.map((w, i) => {
      if (isInit(w)) return w.endsWith('.') ? w : w + '.';
      let r = decl.get(w + '#' + i)![k];
      if (opt.initials && w !== surname) r = r[0].toUpperCase() + '.';
      return keepCase(w, r);
    }).join(' ');
  }
  return { gender, forms: out };
}

// ---------------------------------------------------------------- должности

const ADJ = /(ый|ий|ой|ая|яя|ое|ее|ые|ие)$/;

function adjForms(w: string): Forms {
  const low = w.toLowerCase();
  const s = w.slice(0, -2);
  if (/(ая|яя)$/.test(low)) {
    const soft = low.endsWith('яя');
    if (!soft && HISS.includes(low.slice(-3, -2))) return [w, s + 'ей', s + 'ей', s + 'ую', s + 'ей', s + 'ей'];
    return soft ? [w, s + 'ей', s + 'ей', s + 'юю', s + 'ей', s + 'ей'] : [w, s + 'ой', s + 'ой', s + 'ую', s + 'ой', s + 'ой'];
  }
  if (/(ой)$/.test(low)) return [w, s + 'ого', s + 'ому', s + 'ого', s + 'ым', s + 'ом'];
  const soft = (/(ний|ий)$/.test(low) && !/[гкх]ий$/.test(low) && !/(ский|цкий)$/.test(low)) || /[жшчщ]ий$/.test(low);
  if (soft) return [w, s + 'его', s + 'ему', s + 'его', s + 'им', s + 'ем'];
  return [w, s + 'ого', s + 'ому', s + 'ого', low.endsWith('ый') ? s + 'ым' : s + 'им', s + 'ом'];
}

/** Склонить начало должности: прилагательные и главное существительное; остальное («отдела кадров») не меняется.
 *  «Генеральный директор» → «Генеральному директору», «Заместитель директора» → «Заместителю директора». */
export function declinePosition(text: string, c: Case): string {
  const words = text.trim().split(/\s+/);
  const k = idx(c);
  const out: string[] = [];
  let done = false;
  for (const w of words) {
    if (done || !/^[А-Яа-яЁё]/.test(w)) { out.push(w); continue; }
    const low = w.toLowerCase();
    if (ADJ.test(low)) {
      out.push(keepCase(w, adjForms(w)[k]));
      // «Заведующий», «Управляющая», «Исполняющий»: причастие здесь — главное слово.
      if (/(ющ|ящ|ущ|ащ)(ий|ая|ее)$/.test(low)) done = true;
      continue;
    }
    // «Директор», «Заместитель», «Начальник», «Врач»; женские на -а/-я («Заведующая» — прилагательное выше).
    const masc = !/[ая]$/.test(low) || /^(председатель|заместитель)/.test(low);
    out.push(keepCase(w, nounForms(w, masc)[k]));
    done = true;
  }
  return out.join(' ');
}
