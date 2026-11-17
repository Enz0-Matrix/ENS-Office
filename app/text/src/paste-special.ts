// SPDX-License-Identifier: AGPL-3.0-or-later
// Вставка таблиц из электронных таблиц и «Специальная вставка» (Ctrl+Alt+V), как в Word.
// Excel кладёт в буфер HTML-таблицу, у ячеек с формулами — атрибут x:fmla (=SUM(B2:B4)) с адресами ЛИСТА, а не
// вставленного куска. Откуда взят кусок, в буфере не сказано, поэтому смещение подбирается проверкой: при верном
// смещении каждая формула, пересчитанная по вставленным значениям, даёт то же число, что показывал Excel.
// Так таблица приходит с живыми формулами (table-formulas.ts): поменяли число — итог пересчитался.

import type { Node as PMNode, Schema } from 'prosemirror-model';
import { colName, computeGrid, parseNum } from './table-formulas';

export interface PastedCell { text: string; formula: string | null; colspan: number; rowspan: number }
export interface PastedTable { rows: PastedCell[][]; withFormulas: number; shifted: boolean }

/** Похоже ли HTML на таблицу из Excel/Calc (а не на таблицу с веб-страницы). */
export const isSpreadsheetHtml = (html: string): boolean =>
  /<table/i.test(html) && /urn:schemas-microsoft-com:office:excel|x:num|x:fmla|ProgId\s*content="Excel|sdval|sdnum|LibreOffice|ENS-Table/i.test(html);

/** Сдвинуть адреса в формуле на (-dr, -dc): с адресов листа на адреса вставленной таблицы. null — адрес вне таблицы. */
export function shiftFormula(f: string, dr: number, dc: number, rows: number, cols: number): string | null {
  let bad = false;
  const out = f.replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(])/g, (_m, _d1, letters: string, _d2, digits: string) => {
    const c = [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 - dc;
    const r = parseInt(digits, 10) - 1 - dr;
    if (r < 0 || c < 0 || r >= rows || c >= cols) bad = true;
    return `${colName(Math.max(0, c))}${Math.max(0, r) + 1}`;
  });
  return bad ? null : out;
}

/** Адреса в формуле: наименьшая строка и столбец (с нуля) — откуда начинать подбор смещения. */
function refsOf(f: string): { r: number; c: number }[] {
  return [...f.matchAll(/\$?([A-Za-z]{1,3})\$?(\d{1,7})(?![\w(])/g)].map((m) => ({ c: [...m[1].toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1, r: parseInt(m[2], 10) - 1 }));
}

/** Разобрать HTML с таблицей: ячейки, объединения, формулы x:fmla с подбором смещения. */
export function parseSpreadsheetHtml(html: string): PastedTable | null {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  const table = doc.querySelector('table');
  if (!table) return null;
  const rows: PastedCell[][] = [];
  const raw: { r: number; c: number; fmla: string; shown: number }[] = [];
  const taken: boolean[][] = [];
  [...table.querySelectorAll('tr')].forEach((tr, r) => {
    rows[r] ??= [];
    taken[r] ??= [];
    let c = 0;
    for (const td of tr.querySelectorAll('td, th')) {
      while (taken[r][c]) c++;
      const colspan = Math.max(1, parseInt(td.getAttribute('colspan') ?? '1', 10) || 1);
      const rowspan = Math.max(1, parseInt(td.getAttribute('rowspan') ?? '1', 10) || 1);
      const text = (td.textContent ?? '').replace(/\s+/g, ' ').trim();
      const fmla = td.getAttribute('x:fmla');
      rows[r][c] = { text, formula: null, colspan, rowspan };
      if (fmla) {
        const xn = td.getAttribute('x:num');
        const shown = xn != null && xn !== '' ? parseFloat(xn) : parseNum(text);
        raw.push({ r, c, fmla, shown });
      }
      for (let a = 0; a < rowspan; a++) for (let b = 0; b < colspan; b++) { (taken[r + a] ??= [])[c + b] = true; }
      c += colspan;
    }
  });
  const nRows = rows.length, nCols = Math.max(0, ...rows.map((r) => r.length));
  if (!nRows || !nCols) return null;
  let shifted = false;
  if (raw.length) {
    const text = Array.from({ length: nRows }, (_, r) => Array.from({ length: nCols }, (_, c) => rows[r]?.[c]?.text ?? ''));
    // Кандидаты смещения: адреса формул минус адрес внутри таблицы, от которого они могли прийти.
    const refs = raw.flatMap((x) => refsOf(x.fmla));
    const minR = Math.min(...refs.map((x) => x.r)), minC = Math.min(...refs.map((x) => x.c));
    const fits = (dr: number, dc: number): (string | null)[][] | null => {
      const formulas = Array.from({ length: nRows }, () => Array<string | null>(nCols).fill(null));
      for (const x of raw) {
        const f = shiftFormula(x.fmla, dr, dc, nRows, nCols);
        if (!f) return null;
        formulas[x.r][x.c] = f;
      }
      const res = computeGrid(text, formulas);
      for (const x of raw) {
        const v = parseNum(res.get(`${x.r},${x.c}`) ?? '');
        if (Number.isNaN(x.shown)) continue; // не с чем сверить — не мешаем
        if (Number.isNaN(v) || Math.abs(v - x.shown) > Math.max(0.01, Math.abs(x.shown) * 1e-6)) return null;
      }
      return formulas;
    };
    let found: (string | null)[][] | null = null;
    for (let dr = Math.max(0, minR - nRows + 1); dr <= minR && !found; dr++) {
      for (let dc = Math.max(0, minC - nCols + 1); dc <= minC && !found; dc++) {
        found = fits(dr, dc);
        if (found && (dr || dc)) shifted = true;
      }
    }
    if (found) for (const x of raw) rows[x.r][x.c].formula = found[x.r][x.c];
  }
  return { rows, withFormulas: rows.flat().filter((c) => c?.formula).length, shifted };
}

/** Таблица из текста с табуляцией (так копирует ENS-Table и любая электронная таблица): строки — по переводу строки. */
export function parseTsv(text: string): PastedTable | null {
  const lines = text.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n');
  if (!lines.length || !lines.some((l) => l.includes('\t'))) return null;
  const rows = lines.map((l) => l.split('\t').map((t) => ({ text: t.trim(), formula: /^=/.test(t.trim()) ? t.trim() : null, colspan: 1, rowspan: 1 })));
  return { rows, withFormulas: rows.flat().filter((c) => c.formula).length, shifted: false };
}

/** Узел таблицы ProseMirror из разобранной таблицы (границы — как у обычной вставленной таблицы). */
export function tableNode(schema: Schema, t: PastedTable): PMNode {
  const n = schema.nodes;
  const rows = t.rows.map((r) => n.table_row.create(null, r.filter(Boolean).map((c) => {
    const right = !Number.isNaN(parseNum(c.text)) || !!c.formula;
    const para = n.paragraph.create(right ? { align: 'right' } : null, c.text ? schema.text(c.text) : undefined);
    return n.table_cell.create({ colspan: c.colspan, rowspan: c.rowspan, formula: c.formula }, para);
  })));
  return n.table.create({ borders: true }, rows);
}
