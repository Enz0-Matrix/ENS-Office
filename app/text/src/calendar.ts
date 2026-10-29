// Рабочие дни и сроки. Производственный календарь РФ лежит в программе, без интернета.
//
// Откуда берутся нерабочие дни:
//  1) поправки пользователя (перенесённые выходные и рабочие субботы из постановления Правительства) — главный источник;
//  2) праздники РФ (ст. 112 ТК РФ) с переносом выпавших на выходные по общему правилу — это РАСЧЁТ, а не официальный календарь:
//     ежегодные постановления переносят дни по-своему. Для важных сроков сверьте с постановлением и внесите поправки.
// Сокращённые предпраздничные дни считаются рабочими.

/** Дата в виде «ГГГГ-ММ-ДД» — ключ календаря. */
export type Ymd = string;

export interface Corrections { off: Set<Ymd>; work: Set<Ymd> }

const pad = (n: number) => String(n).padStart(2, '0');
export const toYmd = (d: Date): Ymd => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
export function fromYmd(s: Ymd): Date {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!m) throw new Error(`Не дата: ${s}`);
  const d = new Date(+m[1], +m[2] - 1, +m[3]);
  if (toYmd(d) !== s) throw new Error(`Нет такой даты: ${s}`);
  return d;
}
const addDays = (d: Date, n: number): Date => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
const isWeekend = (d: Date) => d.getDay() === 0 || d.getDay() === 6;

/** Праздники РФ (месяц, день). Январские каникулы — 1–8 января. */
const HOLIDAYS: [number, number][] = [[1, 1], [1, 2], [1, 3], [1, 4], [1, 5], [1, 6], [1, 7], [1, 8], [2, 23], [3, 8], [5, 1], [5, 9], [6, 12], [11, 4]];

const cache = new Map<number, Set<Ymd>>();

/** Расчётные нерабочие дни года: праздники + перенос выпавших на выходные на ближайшие следующие будни (ст. 112 ТК РФ). */
export function holidaysOf(year: number): Set<Ymd> {
  const hit = cache.get(year);
  if (hit) return hit;
  const out = new Set<Ymd>();
  const base = HOLIDAYS.map(([m, d]) => new Date(year, m - 1, d));
  for (const d of base) out.add(toYmd(d));
  // Выходной, совпавший с праздником, переносится на следующий день, который ещё не нерабочий.
  for (const d of base) {
    if (!isWeekend(d)) continue;
    let t = addDays(d, 1);
    while (isWeekend(t) || out.has(toYmd(t))) t = addDays(t, 1);
    out.add(toYmd(t));
  }
  cache.set(year, out);
  return out;
}

export const emptyCorrections = (): Corrections => ({ off: new Set(), work: new Set() });

export function isWorkday(d: Date, c: Corrections = emptyCorrections()): boolean {
  const k = toYmd(d);
  if (c.work.has(k)) return true;
  if (c.off.has(k)) return false;
  return !isWeekend(d) && !holidaysOf(d.getFullYear()).has(k);
}

/** Дата через n рабочих дней (n < 0 — назад); сам день начала не считается. */
export function addWorkdays(start: Date, n: number, c?: Corrections): Date {
  const step = n < 0 ? -1 : 1;
  let d = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  for (let left = Math.abs(n); left > 0;) {
    d = addDays(d, step);
    if (isWorkday(d, c)) left--;
    if (Math.abs(d.getFullYear() - start.getFullYear()) > 30) throw new Error('Слишком большой срок');
  }
  return d;
}

/** Календарный срок: если конец выпал на нерабочий день, переносится на ближайший рабочий (ст. 193 ТК РФ, ст. 193 ГК РФ). */
export function addCalendarDays(start: Date, n: number, c?: Corrections, rollToWorkday = true): Date {
  let d = addDays(start, n);
  if (rollToWorkday) while (!isWorkday(d, c)) d = addDays(d, n < 0 ? -1 : 1);
  return d;
}

/** Сколько рабочих дней между датами (начало не считается, конец считается). */
export function workdaysBetween(a: Date, b: Date, c?: Corrections): number {
  if (b < a) return -workdaysBetween(b, a, c);
  let n = 0;
  for (let d = addDays(a, 1); d <= b; d = addDays(d, 1)) if (isWorkday(d, c)) n++;
  return n;
}

/** Разбор поправок: строки «+ГГГГ-ММ-ДД» — выходной, «=ГГГГ-ММ-ДД» — рабочий день; допускается ДД.ММ.ГГГГ. */
export function parseCorrections(text: string): { corr: Corrections; errors: string[] } {
  const corr = emptyCorrections();
  const errors: string[] = [];
  for (const raw of text.split(/[\n;,]+/)) {
    const line = raw.trim();
    if (!line) continue;
    const m = /^([+=])\s*(\d{4}-\d{2}-\d{2}|\d{1,2}\.\d{1,2}\.\d{4})$/.exec(line);
    if (!m) { errors.push(line); continue; }
    let key = m[2];
    if (key.includes('.')) { const [d, mo, y] = key.split('.'); key = `${y}-${pad(+mo)}-${pad(+d)}`; }
    try { fromYmd(key); } catch { errors.push(line); continue; }
    (m[1] === '+' ? corr.off : corr.work).add(key);
  }
  return { corr, errors };
}
