// SPDX-License-Identifier: AGPL-3.0-or-later
// Формулы в ячейках таблицы, как «Формула» в таблице Word: =SUM(B2:B4), =A2*C2, =SUM(ABOVE). Ссылки — в пределах
// своей таблицы: столбцы A, B, C…, строки 1, 2, 3… В ячейке хранится формула (атрибут formula), а виден результат;
// при любой правке документа значения пересчитываются в той же записи истории отмены. В DOCX ячейка пишется полем
// { =формула } с результатом — Word его понимает и пересчитывает (F9).

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { TableMap } from 'prosemirror-tables';

export type Grid = string[][];

/** Число из текста ячейки: «3 658 420», «1 234,5», «12,5 %», «−7» → число или NaN. */
export function parseNum(text: string): number {
  let t = text.replace(/[\s  ]/g, '').replace(/[−–]/g, '-').replace(/[₽$€]|руб\.?|р\.$/gi, '');
  const pct = t.endsWith('%');
  if (pct) t = t.slice(0, -1);
  if (/^-?\d+(,\d+)?$/.test(t)) t = t.replace(',', '.');
  if (!/^-?(\d+\.?\d*|\.\d+)(e-?\d+)?$/i.test(t)) return NaN;
  const v = parseFloat(t);
  return pct ? v / 100 : v;
}

/** Результат по-русски: разряды через неразрывный пробел, до двух знаков после запятой. */
export function formatNum(v: number): string {
  if (!Number.isFinite(v)) return '#ДЕЛ/0!';
  return new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 2, useGrouping: true }).format(v).replace(/ /g, ' ');
}

class FormulaError extends Error {}

const FUNCS: Record<string, (args: number[][]) => number> = {
  SUM: (a) => a.flat().reduce((s, x) => s + x, 0),
  AVERAGE: (a) => { const f = a.flat(); if (!f.length) throw new FormulaError('#ДЕЛ/0!'); return f.reduce((s, x) => s + x, 0) / f.length; },
  MIN: (a) => Math.min(...a.flat()),
  MAX: (a) => Math.max(...a.flat()),
  COUNT: (a) => a.flat().length,
  PRODUCT: (a) => a.flat().reduce((s, x) => s * x, 1),
  ROUND: (a) => { const [x, d = 0] = a.map((v) => v[0] ?? 0); const k = 10 ** d; return Math.round(x * k) / k; },
  ABS: (a) => Math.abs(a[0]?.[0] ?? 0),
  INT: (a) => Math.floor(a[0]?.[0] ?? 0),
};
const ALIAS: Record<string, string> = { СУММ: 'SUM', СРЗНАЧ: 'AVERAGE', МИН: 'MIN', МАКС: 'MAX', СЧЁТ: 'COUNT', СЧЕТ: 'COUNT', ПРОИЗВЕД: 'PRODUCT', ОКРУГЛ: 'ROUND', ЦЕЛОЕ: 'INT' };

const colNum = (letters: string): number => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
export const colName = (c: number): string => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

/** Вычислить формулу ячейки (r, c) таблицы. value(r, c) — значение другой ячейки (с пересчётом её формулы). */
export function evaluate(formula: string, r: number, c: number, rows: number, cols: number, value: (r: number, c: number) => number | null): number {
  // Ключи форматирования Word ( \# "0,00" ) отбрасываем: формат результата у нас свой.
  const src = formula.replace(/\\[#@*][^\\]*$/g, '').trim().replace(/^=/, '');
  let i = 0;
  const peek = () => src.slice(i).match(/^\s*/)![0].length + i;
  const skip = () => { i = peek(); };
  const eat = (re: RegExp): string | null => { skip(); const m = re.exec(src.slice(i)); if (!m || m.index !== 0) return null; i += m[0].length; return m[0]; };
  const cell = (rr: number, cc: number): number | null => {
    if (rr < 0 || cc < 0 || rr >= rows || cc >= cols) throw new FormulaError('#ССЫЛКА!');
    return value(rr, cc);
  };
  const range = (r1: number, c1: number, r2: number, c2: number): number[] => {
    const out: number[] = [];
    for (let a = Math.min(r1, r2); a <= Math.max(r1, r2); a++) for (let b = Math.min(c1, c2); b <= Math.max(c1, c2); b++) {
      const v = cell(a, b);
      if (v != null) out.push(v);
    }
    return out;
  };
  // Слова Word: числа над ячейкой (до первой пустой сверху — как в Word, все подряд), слева и т. д.
  const word = (w: string): number[] | null => {
    switch (w.toUpperCase()) {
      case 'ABOVE': case 'ВЫШЕ': return r > 0 ? range(0, c, r - 1, c) : [];
      case 'BELOW': case 'НИЖЕ': return r < rows - 1 ? range(r + 1, c, rows - 1, c) : [];
      case 'LEFT': case 'ЛЕВЕЕ': return c > 0 ? range(r, 0, r, c - 1) : [];
      case 'RIGHT': case 'ПРАВЕЕ': return c < cols - 1 ? range(r, c + 1, r, cols - 1) : [];
      default: return null;
    }
  };
  const ref = (): [number, number] | null => {
    const m = eat(/\$?[A-Za-z]{1,3}\$?\d{1,5}(?![\w(])/);
    if (!m) return null;
    const p = /\$?([A-Za-z]+)\$?(\d+)/.exec(m)!;
    return [parseInt(p[2], 10) - 1, colNum(p[1])];
  };
  // Аргумент функции: диапазон, слово Word или выражение.
  const arg = (): number[] => {
    const save = i;
    const a = ref();
    if (a && eat(/:/)) {
      const b = ref();
      if (!b) throw new FormulaError('#ОШИБКА!');
      return range(a[0], a[1], b[0], b[1]);
    }
    i = save;
    const w = eat(/(ABOVE|BELOW|LEFT|RIGHT|ВЫШЕ|НИЖЕ|ЛЕВЕЕ|ПРАВЕЕ)\b/i);
    if (w) return word(w)!;
    return [expr()];
  };
  const primary = (): number => {
    if (eat(/\(/)) { const v = expr(); if (!eat(/\)/)) throw new FormulaError('#ОШИБКА!'); return v; }
    const num = eat(/\d+(\.\d+)?(e-?\d+)?/i);
    if (num) return parseFloat(num);
    const save = i;
    const fn = eat(/[A-Za-zА-Яа-яЁё]+(?=\s*\()/);
    if (fn) {
      const name = ALIAS[fn.toUpperCase()] ?? fn.toUpperCase();
      const f = FUNCS[name];
      if (!f) throw new FormulaError('#ИМЯ?');
      eat(/\(/);
      const args: number[][] = [];
      if (!eat(/\)/)) {
        do args.push(arg()); while (eat(/[;,]/));
        if (!eat(/\)/)) throw new FormulaError('#ОШИБКА!');
      }
      return f(args);
    }
    i = save;
    const a = ref();
    if (a) return cell(a[0], a[1]) ?? 0;
    throw new FormulaError('#ОШИБКА!');
  };
  const unary = (): number => { if (eat(/-/)) return -unary(); if (eat(/\+/)) return unary(); let v = primary(); if (eat(/%/)) v /= 100; return v; };
  const power = (): number => { const b = unary(); return eat(/\^/) ? b ** power() : b; };
  const term = (): number => {
    let v = power();
    for (let op = eat(/[*/]/); op; op = eat(/[*/]/)) {
      const x = power();
      if (op === '/' && x === 0) throw new FormulaError('#ДЕЛ/0!');
      v = op === '*' ? v * x : v / x;
    }
    return v;
  };
  function expr(): number {
    let v = term();
    for (let op = eat(/[+-]/); op; op = eat(/[+-]/)) v = op === '+' ? v + term() : v - term();
    return v;
  }
  const v = expr();
  skip();
  if (i < src.length) throw new FormulaError('#ОШИБКА!');
  return v;
}

/** Значения всех формул сетки (текст ячеек и формулы по координатам): результат или текст ошибки. */
export function computeGrid(text: Grid, formulas: (string | null)[][]): Map<string, string> {
  const rows = text.length, cols = Math.max(0, ...text.map((r) => r.length));
  const memo = new Map<string, number | string>();
  const busy = new Set<string>();
  const val = (r: number, c: number): number | null => {
    const f = formulas[r]?.[c];
    if (!f) { const n = parseNum(text[r]?.[c] ?? ''); return Number.isNaN(n) ? null : n; }
    const k = `${r},${c}`;
    if (!memo.has(k)) {
      if (busy.has(k)) throw new FormulaError('#ЦИКЛ!');
      busy.add(k);
      try { memo.set(k, evaluate(f, r, c, rows, cols, val)); }
      catch (e) { if (e instanceof FormulaError) memo.set(k, e.message); else throw e; }
      finally { busy.delete(k); }
    }
    const m = memo.get(k)!;
    if (typeof m === 'string') throw new FormulaError(m);
    return m;
  };
  const out = new Map<string, string>();
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) {
    if (!formulas[r]?.[c]) continue;
    try { out.set(`${r},${c}`, formatNum(val(r, c)!)); } catch (e) { out.set(`${r},${c}`, e instanceof FormulaError ? e.message : '#ОШИБКА!'); }
  }
  return out;
}

/** Сетка таблицы: текст ячеек, формулы и позиции ячеек (объединённые ячейки — в левом верхнем углу). */
export function tableGrid(table: PMNode, pos: number): { text: Grid; formulas: (string | null)[][]; at: (number | null)[][] } {
  const map = TableMap.get(table);
  const text: Grid = [], formulas: (string | null)[][] = [], at: (number | null)[][] = [];
  for (let r = 0; r < map.height; r++) {
    text.push([]); formulas.push([]); at.push([]);
    for (let c = 0; c < map.width; c++) {
      const cp = map.map[r * map.width + c];
      const rect = map.findCell(cp);
      const own = rect.top === r && rect.left === c;
      const cell = table.nodeAt(cp)!;
      text[r].push(own ? cell.textContent : '');
      formulas[r].push(own ? (cell.attrs.formula as string | null) : null);
      at[r].push(own ? pos + 1 + cp : null);
    }
  }
  return { text, formulas, at };
}

/** Переписать результаты формул в документе: сделка для appendTransaction (или null, если всё уже верно). */
export function recalcTables(state: EditorState): Transaction | null {
  let tr: Transaction | null = null;
  const jobs: { pos: number; value: string }[] = [];
  state.doc.descendants((node, pos) => {
    if (node.type.name !== 'table') return node.isBlock && !node.isTextblock;
    let has = false;
    node.descendants((n) => { if (n.attrs.formula) has = true; return !has && n.type.name !== 'paragraph'; });
    if (!has) return false;
    const g = tableGrid(node, pos);
    const res = computeGrid(g.text, g.formulas);
    for (const [k, v] of res) {
      const [r, c] = k.split(',').map(Number);
      const cp = g.at[r][c];
      if (cp == null || g.text[r][c] === v) continue;
      jobs.push({ pos: cp, value: v });
    }
    return false;
  });
  // С конца документа: позиции раньше по тексту не сдвигаются.
  for (const j of jobs.sort((a, b) => b.pos - a.pos)) {
    tr ??= state.tr;
    const cell = tr.doc.nodeAt(j.pos);
    if (!cell) continue;
    const first = cell.firstChild;
    // Пишем в первый абзац ячейки, сохраняя его оформление и оформление первого текста.
    const marks = first?.firstChild?.marks ?? [];
    const from = j.pos + 2, to = j.pos + 1 + (first?.nodeSize ?? 2) - 1;
    tr.replaceWith(from, to, state.schema.text(j.value, marks));
  }
  return tr ? tr.setMeta('ens-formulas', true) : null;
}

const key = new PluginKey('ens-table-formulas');

/** Пересчёт формул после любой правки документа (в той же записи истории). */
export const formulasPlugin = new Plugin({
  key,
  appendTransaction: (trs, _old, state) => (trs.some((t) => t.docChanged && !t.getMeta('ens-formulas')) ? recalcTables(state) : null),
});
