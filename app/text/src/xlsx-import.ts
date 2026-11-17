// SPDX-License-Identifier: AGPL-3.0-or-later
// Таблица из файла Excel (.xlsx) с живыми формулами. Через буфер обмена Excel отдаёт окну программы только
// значения (HTML без формул), поэтому точный путь — взять формулы из самого файла: лист → <c><f>…</f><v>…</v></c>.
// Каждая формула проверяется: пересчитанная по значениям таблицы, она должна дать то же, что сохранил Excel.
// Иначе (функция, которой у нас нет, ссылка на другой лист) ячейка остаётся значением — ничего не искажаем.
// Для вставки значений (Ctrl+V из Excel) — guessTotals: итоговые строка и столбец становятся =SUM(ABOVE)/=SUM(LEFT).

import { strFromU8, unzipSync } from 'fflate';
import type { PastedCell, PastedTable } from './paste-special';
import { colName, computeGrid, formatNum, parseNum } from './table-formulas';

export interface XlsxSheet { name: string; table: PastedTable; dropped: number }

const colNum = (letters: string): number => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const addr = (ref: string): { r: number; c: number } => { const m = /^([A-Z]+)(\d+)$/i.exec(ref)!; return { c: colNum(m[1]), r: parseInt(m[2], 10) - 1 }; };

/** Сдвинуть относительные адреса формулы на (dr, dc); адреса с $ не трогать (как Excel для общих формул). */
function moveFormula(f: string, dr: number, dc: number): string {
  return f.replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(!])/g, (_m, d1: string, letters: string, d2: string, digits: string) => {
    const c = d1 ? colNum(letters) : colNum(letters) + dc;
    const r = d2 ? parseInt(digits, 10) - 1 : parseInt(digits, 10) - 1 + dr;
    return `${d1}${colName(Math.max(0, c))}${d2}${Math.max(0, r) + 1}`;
  });
}

/** Адреса формулы с листа на таблицу: минус левый верхний угол; null — ссылка вне таблицы или на другой лист. */
function toTable(f: string, r0: number, c0: number, rows: number, cols: number): string | null {
  if (/!|\[/.test(f)) return null; // другой лист или книга
  let bad = false;
  const out = f.replace(/\$?([A-Za-z]{1,3})\$?(\d{1,7})(?![\w(])/g, (_m, letters: string, digits: string) => {
    const c = colNum(letters) - c0, r = parseInt(digits, 10) - 1 - r0;
    if (r < 0 || c < 0 || r >= rows || c >= cols) bad = true;
    return `${colName(Math.max(0, c))}${Math.max(0, r) + 1}`;
  });
  return bad ? null : out;
}

const attr = (tag: string, name: string): string | null => new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1] ?? null;
const unxml = (s: string): string => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/** Прочитать все листы книги .xlsx. */
export function readXlsx(bytes: Uint8Array): XlsxSheet[] {
  const files = unzipSync(bytes);
  const get = (p: string): string | null => (files[p] ? strFromU8(files[p]) : null);
  const shared = [...(get('xl/sharedStrings.xml') ?? '').matchAll(/<si>([\s\S]*?)<\/si>/g)]
    .map((m) => unxml([...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join('')));
  const wb = get('xl/workbook.xml') ?? '';
  const rels = get('xl/_rels/workbook.xml.rels') ?? '';
  const target = (id: string): string | null => {
    const m = new RegExp(`<Relationship[^>]*Id="${id}"[^>]*>`).exec(rels)?.[0];
    const t = m ? attr(m, 'Target') : null;
    return t ? (t.startsWith('/') ? t.slice(1) : `xl/${t}`) : null;
  };
  const sheets: XlsxSheet[] = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) {
    const name = unxml(attr(m[0], 'name') ?? `Лист${sheets.length + 1}`);
    const path = target(attr(m[0], 'r:id') ?? '');
    const xml = path ? get(path) : null;
    if (!xml) continue;
    sheets.push(readSheet(name, xml, shared));
  }
  return sheets;
}

function readSheet(name: string, xml: string, shared: string[]): XlsxSheet {
  const cells = new Map<string, { text: string; f: string | null; v: number }>();
  const sharedF = new Map<string, { f: string; r: number; c: number }>();
  let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;
  for (const m of xml.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const head = `<c${m[1]}>`, body = m[2] ?? '';
    const ref = attr(head, 'r');
    if (!ref) continue;
    const { r, c } = addr(ref);
    const t = attr(head, 't');
    const vRaw = /<v>([\s\S]*?)<\/v>/.exec(body)?.[1];
    const isRaw = /<is>([\s\S]*?)<\/is>/.exec(body)?.[1];
    let text = '';
    let v = NaN;
    if (t === 's' && vRaw != null) text = shared[parseInt(vRaw, 10)] ?? '';
    else if (t === 'inlineStr' && isRaw != null) text = unxml([...isRaw.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((x) => x[1]).join(''));
    else if (t === 'str' || t === 'e') text = unxml(vRaw ?? '');
    else if (t === 'b') text = vRaw === '1' ? 'ИСТИНА' : 'ЛОЖЬ';
    else if (vRaw != null && vRaw !== '') { v = parseFloat(vRaw); text = formatNum(v); }
    // Формула: обычная или общая (<f t="shared" si="0" ref="A3:E3">SUM(A1:A2)</f>, у остальных — только si).
    const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(body);
    let f: string | null = null;
    if (fm) {
      const fh = `<f${fm[1]}>`;
      const si = attr(fh, 'si');
      const ftext = fm[2] ? unxml(fm[2]) : '';
      if (attr(fh, 't') === 'shared' && si != null) {
        if (ftext) { sharedF.set(si, { f: ftext, r, c }); f = ftext; }
        else { const base = sharedF.get(si); if (base) f = moveFormula(base.f, r - base.r, c - base.c); }
      } else if (ftext) f = ftext;
    }
    if (!text && !f) continue;
    cells.set(`${r},${c}`, { text, f, v });
    minR = Math.min(minR, r); minC = Math.min(minC, c); maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
  }
  if (maxR < 0) return { name, table: { rows: [], withFormulas: 0, shifted: false }, dropped: 0 };
  const nRows = maxR - minR + 1, nCols = maxC - minC + 1;
  const rows: PastedCell[][] = Array.from({ length: nRows }, (_, r) => Array.from({ length: nCols }, (_, c) => {
    const x = cells.get(`${r + minR},${c + minC}`);
    return { text: x?.text ?? '', formula: null, colspan: 1, rowspan: 1 };
  }));
  // Формулы: перевести на адреса таблицы и проверить пересчётом.
  const text = rows.map((r) => r.map((c) => c.text));
  const formulas = rows.map((r) => r.map(() => null as string | null));
  const wanted: { r: number; c: number; v: number; text: string }[] = [];
  for (const [k, x] of cells) {
    if (!x.f) continue;
    const [r, c] = k.split(',').map(Number);
    const f = toTable(x.f, minR, minC, nRows, nCols);
    if (!f) continue;
    formulas[r - minR][c - minC] = `=${f}`;
    wanted.push({ r: r - minR, c: c - minC, v: x.v, text: x.text });
  }
  let dropped = [...cells.values()].filter((x) => x.f).length - wanted.length;
  // Неподдержанное убираем по одной, пока всё оставшееся не сходится с тем, что сохранил Excel.
  for (let guard = 0; guard < wanted.length + 1; guard++) {
    const res = computeGrid(text, formulas);
    const bad = wanted.filter((w) => formulas[w.r][w.c] && !same(res.get(`${w.r},${w.c}`) ?? '', w));
    if (!bad.length) break;
    for (const b of bad) { formulas[b.r][b.c] = null; dropped++; }
  }
  rows.forEach((r, i) => r.forEach((c, j) => { c.formula = formulas[i][j]; }));
  return { name, table: { rows, withFormulas: rows.flat().filter((c) => c.formula).length, shifted: false }, dropped };
}

function same(got: string, w: { v: number; text: string }): boolean {
  if (Number.isNaN(w.v)) return got.trim() === w.text.trim();
  const g = parseNum(got);
  return !Number.isNaN(g) && Math.abs(g - w.v) <= Math.max(0.005, Math.abs(w.v) * 1e-9);
}

/**
 * Итоги во вставленных значениях: ячейка нижней строки, равная сумме двух и более чисел над ней, становится
 * =SUM(ABOVE); ячейка правого столбца, равная сумме чисел слева, — =SUM(LEFT). Возвращает число распознанных.
 */
export function guessTotals(t: PastedTable): number {
  const rows = t.rows;
  const n = rows.length, m = Math.max(0, ...rows.map((r) => r.length));
  if (n < 2 || m < 1) return 0;
  const num = (r: number, c: number): number => { const x = rows[r]?.[c]; return x && x.colspan === 1 && x.rowspan === 1 ? parseNum(x.text) : NaN; };
  const close = (a: number, b: number) => Math.abs(a - b) <= Math.max(0.005, Math.abs(b) * 1e-9);
  let found = 0;
  // Как SUM(ABOVE)/SUM(LEFT) в table-formulas.ts: все числа столбца выше (строки левее), текст не считается.
  const sumRun = (vals: number[]): { s: number; k: number } => {
    let s = 0, k = 0;
    for (const x of vals) if (!Number.isNaN(x)) { s += x; k++; }
    return { s, k };
  };
  const last = n - 1;
  for (let c = 0; c < m; c++) {
    const v = num(last, c);
    if (Number.isNaN(v) || rows[last][c].formula) continue;
    const { s, k } = sumRun(Array.from({ length: last }, (_, i) => num(last - 1 - i, c)));
    if (k >= 2 && close(s, v)) { rows[last][c].formula = '=SUM(ABOVE)'; found++; }
  }
  const lc = m - 1;
  for (let r = 0; r < n; r++) {
    const v = num(r, lc);
    if (Number.isNaN(v) || !rows[r][lc] || rows[r][lc].formula) continue;
    const { s, k } = sumRun(Array.from({ length: lc }, (_, i) => num(r, lc - 1 - i)));
    if (k >= 2 && close(s, v)) { rows[r][lc].formula = '=SUM(LEFT)'; found++; }
  }
  t.withFormulas = rows.flat().filter((c) => c?.formula).length;
  return found;
}
