// Схема документа ENS-Text (ProseMirror). Единицы: кегль и отступы — пункты (pt), как в docx;
// размеры изображений и ширины столбцов таблиц — пиксели CSS (96 на дюйм).

import { Schema, type DOMOutputSpec, type MarkSpec, type Node as PMNode, type NodeSpec } from 'prosemirror-model';
import { tableNodes } from 'prosemirror-tables';
import { fontStack } from '../../table/src/fonts';
import { FLOAT_CLASS, floatStyle, type Float } from './floats';

/** Свойства абзаца: хранятся прямо в узле (как прямое форматирование в docx). */
export interface ParaAttrs {
  align: 'left' | 'center' | 'right' | 'justify' | null;
  /** Отступ слева, pt. */
  indent: number | null;
  /** Отступ первой строки, pt (отрицательный — выступ). */
  first: number | null;
  before: number | null;
  after: number | null;
  /** Межстрочный интервал: множитель (1, 1.15, 1.5, 2). */
  line: number | null;
  pageBreakBefore: boolean;
}

const paraAttrs = {
  align: { default: null },
  indent: { default: null },
  first: { default: null },
  before: { default: null },
  after: { default: null },
  line: { default: null },
  pageBreakBefore: { default: false },
  /** Метка на полях: 'red' — проверить, 'green' — готово. Живёт в документе и версиях, в docx не сохраняется. */
  flag: { default: null },
};

function paraStyle(a: ParaAttrs): string {
  const s: string[] = [];
  if (a.align) s.push(`text-align:${a.align}`);
  if (a.indent) s.push(`margin-left:${a.indent}pt`);
  if (a.first) s.push(`text-indent:${a.first}pt`);
  if (a.before != null) s.push(`margin-top:${a.before}pt`);
  if (a.after != null) s.push(`margin-bottom:${a.after}pt`);
  // Одинарный интервал Word — естественная высота строки шрифта; остальные кратны ей (~1,15 кегля).
  if (a.line) s.push(`line-height:${a.line === 1 ? 'normal' : (a.line * 1.15).toFixed(3)}`);
  return s.join(';');
}

function paraDom(tag: string, node: PMNode): DOMOutputSpec {
  const a = node.attrs as ParaAttrs;
  const attrs: Record<string, string> = {};
  const style = paraStyle(a);
  if (style) attrs.style = style;
  if (a.pageBreakBefore) attrs['data-page-break'] = '1';
  if ((a as ParaAttrs & { flag?: string | null }).flag) attrs['data-flag'] = String((a as ParaAttrs & { flag?: string }).flag);
  return [tag, attrs, 0];
}

/** Разбор свойств абзаца из вставляемого HTML (style="…"). */
function parseParaAttrs(el: HTMLElement): Partial<ParaAttrs> {
  const st = el.style;
  const pt = (v: string) => {
    const m = /^(-?[\d.]+)(pt|px|cm|mm)?$/.exec(v.trim());
    if (!m) return null;
    const n = parseFloat(m[1]);
    return m[2] === 'px' ? n * 0.75 : m[2] === 'cm' ? n * 28.3465 : m[2] === 'mm' ? n * 2.83465 : n;
  };
  const align = st.textAlign as ParaAttrs['align'];
  return {
    align: ['left', 'center', 'right', 'justify'].includes(align as string) ? align : null,
    indent: st.marginLeft ? pt(st.marginLeft) : null,
    first: st.textIndent ? pt(st.textIndent) : null,
  };
}

const nodes: Record<string, NodeSpec> = {
  doc: { content: 'block+' },
  paragraph: {
    content: 'inline*',
    group: 'block',
    attrs: paraAttrs,
    parseDOM: [{ tag: 'p', getAttrs: (d) => parseParaAttrs(d as HTMLElement) }],
    toDOM: (n) => paraDom('p', n),
  },
  heading: {
    content: 'inline*',
    group: 'block',
    defining: true,
    attrs: { ...paraAttrs, level: { default: 1 } },
    parseDOM: [1, 2, 3, 4, 5, 6].map((level) => ({ tag: `h${level}`, getAttrs: (d: HTMLElement | string) => ({ level, ...parseParaAttrs(d as HTMLElement) }) })),
    toDOM: (n) => paraDom(`h${n.attrs.level}`, n),
  },
  bullet_list: {
    content: 'list_item+',
    group: 'block',
    attrs: { scheme: { default: null } },
    parseDOM: [{ tag: 'ul' }],
    toDOM: () => ['ul', 0],
  },
  ordered_list: {
    content: 'list_item+',
    group: 'block',
    /** fmt: decimal | lowerLetter | upperLetter | lowerRoman | upperRoman. */
    attrs: { order: { default: 1 }, fmt: { default: 'decimal' }, /** Схема многоуровневой нумерации (numbering.ts) — у внешнего списка. */ scheme: { default: null } },
    parseDOM: [{ tag: 'ol', getAttrs: (d) => ({ order: +((d as HTMLElement).getAttribute('start') ?? 1) }) }],
    toDOM: (n) => ['ol', { start: n.attrs.order === 1 ? undefined : n.attrs.order, 'data-fmt': n.attrs.fmt }, 0],
  },
  list_item: {
    content: 'paragraph block*',
    defining: true,
    parseDOM: [{ tag: 'li' }],
    toDOM: () => ['li', 0],
  },
  image: {
    inline: true,
    group: 'inline',
    draggable: true,
    /** `crop` — обрезка рисунка, доли от исходного размера: [слева, сверху, справа, снизу] (как a:srcRect в docx). */
    attrs: { src: {}, alt: { default: '' }, width: { default: null }, height: { default: null }, crop: { default: null }, /** Плавающий рисунок (floats.ts); null — в строке. */ float: { default: null } },
    parseDOM: [{
      tag: 'img[src]',
      getAttrs: (d) => {
        const el = d as HTMLElement;
        const crop = (el.getAttribute('data-crop') ?? '').split(',').map(Number);
        let float = null;
        try { float = JSON.parse(el.getAttribute('data-float') ?? 'null'); } catch { float = null; }
        return {
          float,
          src: el.getAttribute('src'), alt: el.getAttribute('alt') ?? '', width: parseFloat(el.getAttribute('width') ?? '') || null, height: parseFloat(el.getAttribute('height') ?? '') || null,
          crop: crop.length === 4 && crop.every((x) => Number.isFinite(x)) ? crop : null,
        };
      },
    }],
    toDOM: (n) => {
      const a: Record<string, string> = { src: n.attrs.src, alt: n.attrs.alt };
      if (n.attrs.width) a.width = String(Math.round(n.attrs.width));
      if (n.attrs.height) a.height = String(Math.round(n.attrs.height));
      const c = n.attrs.crop as number[] | null;
      // Размер из документа: ширина и пропорции (aspect-ratio) — рисунок можно растянуть, а при сжатии до ширины страницы пропорции остаются.
      const st: string[] = [];
      if (n.attrs.width && n.attrs.height) st.push(`width:${Math.round(n.attrs.width)}px`, `aspect-ratio:${Math.round(n.attrs.width)}/${Math.round(n.attrs.height)}`, 'height:auto');
      if (c && c.some((x) => x > 0)) {
        // Обрезка: окно просмотра по исходной картинке (object-view-box), размер рамки — как в документе.
        const pc = c.map((x) => `${(x * 100).toFixed(3)}%`);
        a['data-crop'] = c.map((x) => x.toFixed(5)).join(',');
        st.push(`object-view-box:inset(${pc[1]} ${pc[2]} ${pc[3]} ${pc[0]})`);
      }
      if (st.length) a.style = st.join(';');
      const fl = n.attrs.float as Float | null;
      if (fl) {
        a['data-float'] = JSON.stringify(fl);
        a.class = FLOAT_CLASS[fl.mode];
        a.style = `${floatStyle(fl, n.attrs.width, n.attrs.height)};${a.style ?? ''}`;
      }
      return ['img', a];
    },
  },
  /**
   * Разрыв раздела («со следующей страницы»). `page` — параметры страницы раздела, который начинается после разрыва
   * (размер, ориентация, поля); null — как у предыдущего. Первый раздел описан в DocMeta.page.
   */
  section_break: {
    group: 'block',
    atom: true,
    selectable: true,
    attrs: { page: { default: null } },
    parseDOM: [{ tag: 'div.sect-break', getAttrs: (d) => { try { return { page: JSON.parse((d as HTMLElement).getAttribute('data-page') ?? 'null') }; } catch { return { page: null }; } } }],
    toDOM: (n) => {
      const p = n.attrs.page as { w: number; h: number } | null;
      const label = p ? `Разрыв раздела · ${p.w > p.h ? 'альбомная' : 'книжная'}` : 'Разрыв раздела';
      return ['div', { class: 'sect-break', 'data-label': label, 'data-page': JSON.stringify(p), contenteditable: 'false' }];
    },
  },
  /** Поле номера страницы (PAGE) или числа страниц (NUMPAGES): в колонтитулах заменяется числом на каждой странице. */
  field: {
    inline: true,
    group: 'inline',
    atom: true,
    attrs: { kind: { default: 'page' } },
    parseDOM: [{ tag: 'span.fld', getAttrs: (d) => ({ kind: (d as HTMLElement).getAttribute('data-kind') === 'pages' ? 'pages' : 'page' }) }],
    toDOM: (n) => ['span', { class: 'fld', 'data-kind': n.attrs.kind, title: n.attrs.kind === 'pages' ? 'Число страниц' : 'Номер страницы' }, n.attrs.kind === 'pages' ? '‹N›' : '‹№›'],
  },
  hard_break: {
    inline: true,
    group: 'inline',
    selectable: false,
    parseDOM: [{ tag: 'br' }],
    toDOM: () => ['br'],
  },
  text: { group: 'inline' },
  ...tableNodes({
    tableGroup: 'block',
    cellContent: 'block+',
    cellAttributes: {
      background: {
        default: null,
        getFromDOM: (d) => (d as HTMLElement).style.backgroundColor || null,
        setDOMAttr: (v, attrs) => { if (v) attrs.style = `${attrs.style ?? ''};background-color:${v}`; },
      },
    },
  }),
};

// Таблица: признак «с границами» (стиль «Сетка таблицы» в docx).
nodes.table = {
  ...nodes.table,
  attrs: { borders: { default: true } },
  toDOM: (n) => ['table', { class: n.attrs.borders ? 'bordered' : 'plain' }, ['tbody', 0]],
};

/** Выделение «Основной текст» не нужно: текст без шрифта наследует шрифт документа. */
const marks: Record<string, MarkSpec> = {
  link: {
    attrs: { href: {} },
    inclusive: false,
    parseDOM: [{ tag: 'a[href]', getAttrs: (d) => ({ href: (d as HTMLElement).getAttribute('href') }) }],
    toDOM: (m) => ['a', { href: m.attrs.href, title: m.attrs.href }, 0],
  },
  strong: {
    parseDOM: [{ tag: 'strong' }, { tag: 'b' }, { style: 'font-weight', getAttrs: (v) => /^(bold(er)?|[5-9]\d\d)$/.test(v as string) && null }],
    toDOM: () => ['strong', 0],
  },
  em: { parseDOM: [{ tag: 'i' }, { tag: 'em' }, { style: 'font-style=italic' }], toDOM: () => ['em', 0] },
  underline: { parseDOM: [{ tag: 'u' }, { style: 'text-decoration=underline' }], toDOM: () => ['u', 0] },
  strike: { parseDOM: [{ tag: 's' }, { tag: 'del' }, { tag: 'strike' }, { style: 'text-decoration=line-through' }], toDOM: () => ['s', 0] },
  sup: { excludes: 'sub', parseDOM: [{ tag: 'sup' }], toDOM: () => ['sup', 0] },
  sub: { excludes: 'sup', parseDOM: [{ tag: 'sub' }], toDOM: () => ['sub', 0] },
  /** Шрифт, кегль (pt), цвет текста и заливка (#RRGGBB) — одним слоем. */
  textStyle: {
    attrs: { font: { default: null }, size: { default: null }, color: { default: null }, bg: { default: null } },
    parseDOM: [{
      tag: 'span[style]',
      getAttrs: (d) => {
        const st = (d as HTMLElement).style;
        const font = st.fontFamily ? st.fontFamily.split(',')[0].replace(/["']/g, '').trim() : null;
        const px = /([\d.]+)px/.exec(st.fontSize);
        const pt = /([\d.]+)pt/.exec(st.fontSize);
        const size = pt ? parseFloat(pt[1]) : px ? Math.round(parseFloat(px[1]) * 0.75 * 2) / 2 : null;
        const color = st.color || null;
        const bg = st.backgroundColor || null;
        return font || size || color || bg ? { font, size, color, bg } : false;
      },
    }],
    toDOM: (m) => {
      const s: string[] = [];
      if (m.attrs.font) s.push(`font-family:${fontStack(m.attrs.font)}`);
      if (m.attrs.size) s.push(`font-size:${m.attrs.size}pt`);
      if (m.attrs.color) s.push(`color:${m.attrs.color}`);
      if (m.attrs.bg) s.push(`background-color:${m.attrs.bg}`);
      return ['span', { style: s.join(';') }, 0];
    },
  },
};

export const schema = new Schema({ nodes, marks });

/** Нумерация для CSS: list-style-type по формату. */
export const LIST_TYPES: Record<string, string> = {
  decimal: 'decimal', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman',
};
