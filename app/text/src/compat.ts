// Индикатор расхождений с Word: что в документе может выглядеть иначе после сохранения в docx и открытия в Word/LibreOffice.
// Проверка по правилам, без запуска Word: честно говорит о рисках, а не обещает «вёрстка 1:1».

import type { Node as PMNode } from 'prosemirror-model';
import type { DocMeta } from './docx';

export interface Finding {
  id: string;
  level: 'warn' | 'info';
  title: string;
  detail: string;
  /** Сколько мест затронуто. */
  count: number;
  /** Первое место в документе — для перехода. */
  pos?: number;
}

/** Шрифты, которые есть в Word на Windows и Mac из коробки, и свободные метрические аналоги из Linux/Р7/LibreOffice. */
const SAFE_FONTS = new Set([
  'times new roman', 'arial', 'calibri', 'cambria', 'courier new', 'verdana', 'tahoma', 'segoe ui', 'georgia', 'consolas', 'trebuchet ms',
  'comic sans ms', 'impact', 'palatino linotype', 'century gothic', 'garamond', 'book antiqua', 'lucida console', 'arial narrow', 'calibri light',
  'candara', 'constantia', 'corbel', 'symbol', 'wingdings', 'cambria math', 'segoe ui symbol', 'franklin gothic medium',
  'liberation sans', 'liberation serif', 'liberation mono', 'carlito', 'caladea',
]);
/** Шрифты, которые ENS-Text подставляет сам, а в Word их нет. */
const BUNDLED = new Set(['inter', 'roboto condensed', 'roboto condensed light']);

const isTable = (n: PMNode) => n.type.name === 'table';

export function analyze(doc: PMNode, meta: DocMeta, pages: number): Finding[] {
  const out: Finding[] = [];
  const textW = (meta.page.w - meta.page.left - meta.page.right) / 0.75; // px
  const fonts = new Map<string, { count: number; pos: number }>();
  const noteFont = (name: string | null | undefined, pos: number) => {
    if (!name) return;
    const e = fonts.get(name);
    if (e) e.count++; else fonts.set(name, { count: 1, pos });
  };
  noteFont(meta.defaults.font, 0);
  let fracSizes = 0, fracPos: number | undefined;
  const fracNote = (size: number | null | undefined, pos: number) => {
    if (size && Math.abs(size * 2 - Math.round(size * 2)) > 1e-6) { fracSizes++; fracPos ??= pos; }
  };
  fracNote(meta.defaults.size, 0);
  const badImg = new Map<string, { count: number; pos: number }>();
  let wideImg = 0, wideImgPos: number | undefined;
  let wideTable = 0, wideTablePos: number | undefined;
  let nested = 0, nestedPos: number | undefined;

  doc.descendants((node, pos, parent) => {
    if (node.isText) {
      for (const m of node.marks) if (m.type.name === 'textStyle') { noteFont(m.attrs.font, pos); fracNote(m.attrs.size, pos); }
      return false;
    }
    if (node.type.name === 'image') {
      const src = String(node.attrs.src ?? '');
      const mime = /^data:image\/([a-z0-9.+-]+)/i.exec(src)?.[1]?.toLowerCase();
      if (mime && !['png', 'jpeg', 'jpg', 'gif', 'bmp'].includes(mime)) {
        const e = badImg.get(mime);
        if (e) e.count++; else badImg.set(mime, { count: 1, pos });
      }
      if (node.attrs.width && node.attrs.width > textW + 1) { wideImg++; wideImgPos ??= pos; }
      return false;
    }
    if (isTable(node)) {
      if (parent && parent.type.name === 'table_cell') { nested++; nestedPos ??= pos; }
      let widest = 0;
      node.forEach((row) => {
        let w = 0, known = true;
        row.forEach((cell) => {
          const cw = cell.attrs.colwidth as number[] | null;
          if (cw && cw.length) w += cw.reduce((a, b) => a + b, 0); else known = false;
        });
        if (known) widest = Math.max(widest, w);
      });
      if (widest > textW * 1.02) { wideTable++; wideTablePos ??= pos; }
    }
    return true;
  });

  for (const [name, e] of fonts) {
    const low = name.toLowerCase();
    if (SAFE_FONTS.has(low)) continue;
    out.push({
      id: `font:${name}`, level: 'warn', count: e.count, pos: e.pos,
      title: `Шрифт «${name}» может не оказаться у получателя`,
      detail: BUNDLED.has(low)
        ? 'Этот шрифт встроен в ENS-Text, но в Word его нет: Word подставит другой, ширина букв изменится, строки и страницы сдвинутся. Для отправки лучше Times New Roman, Arial или Calibri.'
        : 'Если шрифта нет на компьютере получателя, Word подставит другой: строки и страницы сдвинутся. Для отправки лучше Times New Roman, Arial или Calibri — или отправьте PDF.',
    });
  }
  for (const [mime, e] of badImg) {
    out.push({ id: `img:${mime}`, level: 'warn', count: e.count, pos: e.pos, title: `Рисунок формата ${mime.toUpperCase()}`, detail: 'Старые версии Word и многие просмотрщики не показывают такие рисунки. Надёжнее PNG или JPEG.' });
  }
  if (wideImg) out.push({ id: 'img:wide', level: 'warn', count: wideImg, pos: wideImgPos, title: 'Рисунок шире полосы текста', detail: 'В Word он может уйти за поле или сжаться, и страница сверстается иначе.' });
  if (wideTable) out.push({ id: 'table:wide', level: 'warn', count: wideTable, pos: wideTablePos, title: 'Таблица шире полосы текста', detail: 'Word подгонит ширину столбцов по-своему — переносы строк в ячейках и разрывы страниц изменятся.' });
  if (nested) out.push({ id: 'table:nested', level: 'warn', count: nested, pos: nestedPos, title: 'Вложенная таблица', detail: 'Word и LibreOffice по-разному разрывают такие таблицы между страницами и подбирают высоту строк.' });
  if (fracSizes) out.push({ id: 'size:frac', level: 'warn', count: fracSizes, pos: fracPos, title: 'Дробный кегль', detail: 'В docx кегль хранится с шагом 0,5 пт — значения вроде 11,3 пт будут округлены, текст чуть сдвинется.' });
  for (const l of meta.lossy) out.push({ id: `lossy:${l}`, level: 'warn', count: 1, title: 'Не перенесено при открытии', detail: `${l}. Этого в документе уже нет.` });

  out.push({
    id: 'pages', level: 'info', count: pages,
    title: `ENS-Text насчитал страниц: ${pages}`,
    detail: 'Разбиение на страницы в Word может отличаться на строку-две (по-разному переносятся слова и считается высота строки), особенно в длинных документах и при подстановке шрифтов. Если важна точная вёрстка — отправляйте PDF (Печать → Сохранить как PDF).',
  });
  return out;
}

export const verdict = (f: Finding[]): string => {
  const warns = f.filter((x) => x.level === 'warn').length;
  return warns ? `Возможны расхождения с Word: ${warns}` : 'Серьёзных расхождений с Word не ожидается';
};
