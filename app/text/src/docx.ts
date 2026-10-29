// Чтение и запись .docx (Office Open XML): абзацы, заголовки, начертание, шрифты и цвета,
// списки, таблицы, изображения, гиперссылки, разрывы страниц, размер страницы и поля.
// Чего версия 0.1 не умеет (колонтитулы, сноски, примечания, фигуры, формулы, разделы…),
// не пропадает молча: список попадает в DocMeta.lossy, и программа не даёт сохранить поверх
// исходного файла (ADR-005: исходный файл не портим).

import { Mark, type Node as PMNode } from 'prosemirror-model';
import { TableMap } from 'prosemirror-tables';
import { strFromU8, strToU8, unzipSync, zipSync, type Zippable } from 'fflate';
import { schema } from './schema';

export interface PageSetup { w: number; h: number; top: number; bottom: number; left: number; right: number }
export interface DocDefaults { font: string; size: number; line: number; before: number; after: number }
export interface DocMeta {
  page: PageSetup;
  defaults: DocDefaults;
  /** Что при открытии не удалось перенести (человекочитаемо, с количеством). */
  lossy: string[];
}

/** A4, поля 2 / 2 / 3 / 1,5 см — привычные для российских документов. */
export const A4: PageSetup = { w: 595.3, h: 841.9, top: 56.7, bottom: 56.7, left: 85, right: 42.5 };
export const DEFAULTS: DocDefaults = { font: 'Times New Roman', size: 12, line: 1, before: 0, after: 0 };

export function newMeta(): DocMeta {
  return { page: { ...A4 }, defaults: { ...DEFAULTS }, lossy: [] };
}

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};
const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

const HIGHLIGHT: Record<string, string> = {
  yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF',
};

// ---------------------------------------------------------------- чтение XML

function parseXml(text: string): Document {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) throw new Error('Повреждённый XML внутри docx');
  return doc;
}

const kids = (el: Element | null | undefined, name?: string): Element[] =>
  el ? [...el.children].filter((c) => !name || c.localName === name) : [];
const kid = (el: Element | null | undefined, name: string): Element | undefined => kids(el, name)[0];
const attr = (el: Element | null | undefined, name: string): string | null =>
  el ? (el.getAttributeNS(NS.w, name) ?? el.getAttribute(`w:${name}`)) : null;
const relAttr = (el: Element, name: string): string | null => el.getAttributeNS(NS.r, name) ?? el.getAttribute(`r:${name}`);
const num = (v: string | null): number | null => (v == null || v === '' || isNaN(+v) ? null : +v);
/** Переключатель Word: <w:b/> — да, <w:b w:val="0"/> — нет. */
const flag = (el: Element | undefined): boolean | undefined => {
  if (!el) return undefined;
  const v = attr(el, 'val');
  return !(v === '0' || v === 'false' || v === 'off');
};
const descendants = (el: Element, name: string): Element[] => [...el.getElementsByTagName('*')].filter((e) => e.localName === name);

// ---------------------------------------------------------------- свойства

/** Шрифты темы документа (word/theme/theme1.xml): в docx шрифт часто задан не именем, а ссылкой на тему. */
let themeFonts: { minor: string; major: string } = { minor: '', major: '' };

interface RFmt { b?: boolean; i?: boolean; u?: boolean; s?: boolean; sup?: boolean; sub?: boolean; font?: string; size?: number; color?: string; bg?: string }
interface PFmt {
  align?: 'left' | 'center' | 'right' | 'justify'; left?: number; first?: number; before?: number; after?: number; line?: number;
  pageBreakBefore?: boolean; numId?: number; ilvl?: number; styleId?: string;
}
interface StyleInfo { id: string; name: string; type: string; basedOn?: string; isDefault: boolean; p: PFmt; r: RFmt }

function parseRPr(rPr: Element | undefined): RFmt {
  const f: RFmt = {};
  if (!rPr) return f;
  const set = <K extends keyof RFmt>(k: K, v: RFmt[K] | undefined) => { if (v !== undefined) f[k] = v; };
  set('b', flag(kid(rPr, 'b')));
  set('i', flag(kid(rPr, 'i')));
  set('s', flag(kid(rPr, 'strike')));
  const u = kid(rPr, 'u');
  if (u) set('u', attr(u, 'val') !== 'none');
  const va = attr(kid(rPr, 'vertAlign'), 'val');
  if (va === 'superscript') f.sup = true;
  else if (va === 'subscript') f.sub = true;
  const fonts = kid(rPr, 'rFonts');
  const themed = (k: string | null) => (k ? (/^major/i.test(k) ? themeFonts.major : themeFonts.minor) : '');
  const font = attr(fonts, 'ascii') ?? attr(fonts, 'hAnsi') ?? (themed(attr(fonts, 'asciiTheme') ?? attr(fonts, 'hAnsiTheme')) || null) ?? attr(fonts, 'cs') ?? attr(fonts, 'eastAsia');
  if (font) f.font = font;
  const sz = num(attr(kid(rPr, 'sz'), 'val'));
  if (sz) f.size = sz / 2;
  const color = attr(kid(rPr, 'color'), 'val');
  if (color && color !== 'auto' && /^[0-9a-f]{6}$/i.test(color)) f.color = `#${color.toUpperCase()}`;
  const hl = attr(kid(rPr, 'highlight'), 'val');
  if (hl && HIGHLIGHT[hl]) f.bg = HIGHLIGHT[hl];
  const fill = attr(kid(rPr, 'shd'), 'fill');
  if (fill && fill !== 'auto' && /^[0-9a-f]{6}$/i.test(fill)) f.bg = `#${fill.toUpperCase()}`;
  return f;
}

function parsePPr(pPr: Element | undefined): PFmt {
  const f: PFmt = {};
  if (!pPr) return f;
  const jc = attr(kid(pPr, 'jc'), 'val');
  if (jc) f.align = jc === 'center' ? 'center' : jc === 'right' || jc === 'end' ? 'right' : jc === 'both' || jc === 'distribute' ? 'justify' : 'left';
  const ind = kid(pPr, 'ind');
  if (ind) {
    const left = num(attr(ind, 'left') ?? attr(ind, 'start'));
    if (left != null) f.left = left / 20;
    const first = num(attr(ind, 'firstLine'));
    const hang = num(attr(ind, 'hanging'));
    if (hang) f.first = -hang / 20;
    else if (first != null) f.first = first / 20;
  }
  const sp = kid(pPr, 'spacing');
  if (sp) {
    const before = num(attr(sp, 'before'));
    const after = num(attr(sp, 'after'));
    if (before != null) f.before = before / 20;
    if (after != null) f.after = after / 20;
    const line = num(attr(sp, 'line'));
    const rule = attr(sp, 'lineRule');
    if (line && (!rule || rule === 'auto')) f.line = Math.round((line / 240) * 100) / 100;
  }
  if (flag(kid(pPr, 'pageBreakBefore'))) f.pageBreakBefore = true;
  const np = kid(pPr, 'numPr');
  if (np) {
    f.numId = num(attr(kid(np, 'numId'), 'val')) ?? undefined;
    f.ilvl = num(attr(kid(np, 'ilvl'), 'val')) ?? undefined;
  }
  const st = attr(kid(pPr, 'pStyle'), 'val');
  if (st) f.styleId = st;
  return f;
}

// ---------------------------------------------------------------- импорт

interface ListInfo { kind: 'bullet' | 'ordered'; fmt: string }

interface Ctx {
  files: Record<string, Uint8Array>;
  rels: Map<string, { target: string; type: string; external: boolean }>;
  styles: Map<string, StyleInfo>;
  defaultPara?: StyleInfo;
  defaults: DocDefaults;
  /** numId → уровни. */
  nums: Map<number, ListInfo[]>;
  lossy: Map<string, number>;
  /** Форматирование знаков из стиля текущего абзаца: под прямым форматированием прогона. */
  baseR: RFmt;
  /** Стек сложных полей (TOC, PAGE, HYPERLINK…): в какой фазе, и ссылка, если поле — HYPERLINK. */
  fields: { phase: 'instr' | 'result'; instr: string; href?: string }[];
}

const lose = (ctx: Ctx, what: string, n = 1) => ctx.lossy.set(what, (ctx.lossy.get(what) ?? 0) + n);

function readRels(files: Record<string, Uint8Array>, path: string): Ctx['rels'] {
  const out: Ctx['rels'] = new Map();
  const dir = path.slice(0, path.lastIndexOf('/') + 1);
  const relsPath = `${dir}_rels/${path.slice(dir.length)}.rels`;
  if (!files[relsPath]) return out;
  for (const r of descendants(parseXml(strFromU8(files[relsPath])).documentElement, 'Relationship')) {
    const external = r.getAttribute('TargetMode') === 'External';
    let target = r.getAttribute('Target') ?? '';
    if (!external) target = target.startsWith('/') ? target.slice(1) : normalizePath(dir + target);
    out.set(r.getAttribute('Id') ?? '', { target, type: r.getAttribute('Type') ?? '', external });
  }
  return out;
}

function normalizePath(p: string): string {
  const out: string[] = [];
  for (const part of p.split('/')) {
    if (part === '..') out.pop();
    else if (part && part !== '.') out.push(part);
  }
  return out.join('/');
}

function readStyles(xml: string | undefined, ctx: Ctx): void {
  if (!xml) return;
  const root = parseXml(xml).documentElement;
  const dd = kid(root, 'docDefaults');
  const rd = parseRPr(kid(kid(dd, 'rPrDefault'), 'rPr'));
  const pd = parsePPr(kid(kid(dd, 'pPrDefault'), 'pPr'));
  if (rd.font) ctx.defaults.font = rd.font;
  if (rd.size) ctx.defaults.size = rd.size;
  if (pd.line) ctx.defaults.line = pd.line;
  if (pd.before != null) ctx.defaults.before = pd.before;
  if (pd.after != null) ctx.defaults.after = pd.after;
  for (const s of kids(root, 'style')) {
    const id = attr(s, 'styleId');
    if (!id) continue;
    ctx.styles.set(id, {
      id, name: attr(kid(s, 'name'), 'val') ?? id, type: attr(s, 'type') ?? 'paragraph',
      basedOn: attr(kid(s, 'basedOn'), 'val') ?? undefined, isDefault: attr(s, 'default') === '1',
      p: parsePPr(kid(s, 'pPr')), r: parseRPr(kid(s, 'rPr')),
    });
  }
  ctx.defaultPara = [...ctx.styles.values()].find((s) => s.type === 'paragraph' && s.isDefault);
  // Обычный стиль («Normal») может переопределять шрифт и размер.
  if (ctx.defaultPara) {
    const n = resolveStyle(ctx, ctx.defaultPara.id);
    if (n.r.font) ctx.defaults.font = n.r.font;
    if (n.r.size) ctx.defaults.size = n.r.size;
    if (n.p.line) ctx.defaults.line = n.p.line;
    if (n.p.before != null) ctx.defaults.before = n.p.before;
    if (n.p.after != null) ctx.defaults.after = n.p.after;
  }
}

function resolveStyle(ctx: Ctx, id: string | undefined, depth = 0): { p: PFmt; r: RFmt } {
  const s = id ? ctx.styles.get(id) : undefined;
  if (!s || depth > 20) return { p: {}, r: {} };
  const base = s.basedOn ? resolveStyle(ctx, s.basedOn, depth + 1) : { p: {}, r: {} };
  return { p: { ...base.p, ...s.p }, r: { ...base.r, ...s.r } };
}

function readNumbering(xml: string | undefined, ctx: Ctx): void {
  if (!xml) return;
  const root = parseXml(xml).documentElement;
  const abstracts = new Map<number, ListInfo[]>();
  for (const a of kids(root, 'abstractNum')) {
    const levels: ListInfo[] = [];
    for (const l of kids(a, 'lvl')) {
      const fmt = attr(kid(l, 'numFmt'), 'val') ?? 'decimal';
      levels[num(attr(l, 'ilvl')) ?? 0] = fmt === 'bullet' || fmt === 'none' ? { kind: 'bullet', fmt } : { kind: 'ordered', fmt };
    }
    abstracts.set(num(attr(a, 'abstractNumId')) ?? 0, levels);
  }
  for (const n of kids(root, 'num')) {
    const levels = abstracts.get(num(attr(kid(n, 'abstractNumId'), 'val')) ?? -1);
    if (levels) ctx.nums.set(num(attr(n, 'numId')) ?? 0, levels);
  }
}

function listInfo(ctx: Ctx, numId: number | undefined, ilvl: number): ListInfo | null {
  if (!numId) return null;
  const levels = ctx.nums.get(numId);
  if (!levels) return { kind: 'bullet', fmt: 'bullet' };
  return levels[ilvl] ?? levels[levels.length - 1] ?? { kind: 'bullet', fmt: 'bullet' };
}

const ORDERED_FMTS = ['decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'];
const orderedFmt = (fmt: string) => (ORDERED_FMTS.includes(fmt) ? fmt : 'decimal');

function bytesToDataUrl(bytes: Uint8Array, mime: string): string {
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return `data:${mime};base64,${btoa(bin)}`;
}

const MIME: Record<string, string> = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' };

/** Состояние разбора одного абзаца: накопленное содержимое и ссылка. */
interface InlineAcc { nodes: PMNode[]; chunks: { nodes: PMNode[]; breakBefore: boolean }[]; breakNext: boolean }

function marksFor(ctx: Ctx, fmt: RFmt, href: string | undefined): readonly Mark[] {
  const m: Mark[] = [];
  const d = ctx.defaults;
  if (fmt.b) m.push(schema.marks.strong.create());
  if (fmt.i) m.push(schema.marks.em.create());
  if (fmt.u) m.push(schema.marks.underline.create());
  if (fmt.s) m.push(schema.marks.strike.create());
  if (fmt.sup) m.push(schema.marks.sup.create());
  else if (fmt.sub) m.push(schema.marks.sub.create());
  const font = fmt.font && fmt.font !== d.font ? fmt.font : null;
  const size = fmt.size && fmt.size !== d.size ? fmt.size : null;
  if (font || size || fmt.color || fmt.bg) m.push(schema.marks.textStyle.create({ font, size, color: fmt.color ?? null, bg: fmt.bg ?? null }));
  if (href) m.push(schema.marks.link.create({ href }));
  return Mark.setFrom(m);
}

function addText(acc: InlineAcc, text: string, marks: readonly Mark[]): void {
  if (text) acc.nodes.push(schema.text(text, marks));
}

function parseDrawing(ctx: Ctx, el: Element, acc: InlineAcc, marks: readonly Mark[]): void {
  const blip = descendants(el, 'blip')[0];
  const embed = blip ? relAttr(blip, 'embed') : null;
  const rel = embed ? ctx.rels.get(embed) : undefined;
  const bytes = rel ? ctx.files[rel.target] : undefined;
  const ext = rel?.target.split('.').pop()?.toLowerCase() ?? '';
  if (!bytes || !MIME[ext]) {
    lose(ctx, blip ? 'изображения в неподдерживаемом формате' : 'фигуры и надписи');
    return;
  }
  const extent = descendants(el, 'extent')[0];
  const cx = num(extent?.getAttribute('cx') ?? null);
  const cy = num(extent?.getAttribute('cy') ?? null);
  const docPr = descendants(el, 'docPr')[0];
  const isAnchor = !!descendants(el, 'anchor').length;
  if (isAnchor) lose(ctx, 'обтекание изображений (изображения стали строчными)');
  acc.nodes.push(schema.nodes.image.create({
    src: bytesToDataUrl(bytes, MIME[ext]),
    alt: docPr?.getAttribute('descr') ?? '',
    width: cx ? Math.round(cx / 9525) : null,
    height: cy ? Math.round(cy / 9525) : null,
  }, null, marks));
}

function parseRun(ctx: Ctx, r: Element, acc: InlineAcc, href: string | undefined, charFmt: boolean): void {
  const rPr = kid(r, 'rPr');
  let fmt = parseRPr(rPr);
  const rs = attr(kid(rPr, 'rStyle'), 'val');
  fmt = { ...ctx.baseR, ...(rs && charFmt ? resolveStyle(ctx, rs).r : {}), ...fmt };
  const marks = marksFor(ctx, fmt, href);
  const field = ctx.fields[ctx.fields.length - 1];
  for (const c of r.children) {
    switch (c.localName) {
      case 'fldChar': {
        const t = attr(c, 'fldCharType');
        if (t === 'begin') ctx.fields.push({ phase: 'instr', instr: '' });
        else if (t === 'separate' && field) {
          field.phase = 'result';
          const m = /^\s*HYPERLINK\s+"([^"]+)"/i.exec(field.instr);
          if (m) field.href = m[1];
          else if (!/^\s*HYPERLINK/i.test(field.instr)) lose(ctx, 'поля (оглавление, номера страниц и т. п. стали обычным текстом)');
        } else if (t === 'end') ctx.fields.pop();
        break;
      }
      case 'instrText':
        if (field) field.instr += c.textContent ?? '';
        break;
      case 't':
        if (!field || field.phase === 'result') addText(acc, c.textContent ?? '', field?.href && !href ? marksFor(ctx, fmt, field.href) : marks);
        break;
      case 'delText':
        break;
      case 'tab':
        if (!field || field.phase === 'result') addText(acc, '\t', marks);
        break;
      case 'noBreakHyphen':
        addText(acc, '‑', marks);
        break;
      case 'br':
      case 'cr':
        if (c.localName === 'br' && attr(c, 'type') === 'page') {
          acc.chunks.push({ nodes: acc.nodes, breakBefore: acc.breakNext });
          acc.nodes = [];
          acc.breakNext = true;
        } else if (!field || field.phase === 'result') {
          acc.nodes.push(schema.nodes.hard_break.create(null, null, marks));
        }
        break;
      case 'drawing':
        if (!field || field.phase === 'result') parseDrawing(ctx, c, acc, marks);
        break;
      case 'pict':
      case 'object': {
        const img = descendants(c, 'imagedata')[0];
        const id = img ? relAttr(img, 'id') : null;
        const rel = id ? ctx.rels.get(id) : undefined;
        const ext = rel?.target.split('.').pop()?.toLowerCase() ?? '';
        const bytes = rel ? ctx.files[rel.target] : undefined;
        if (bytes && MIME[ext]) {
          const style = descendants(c, 'shape')[0]?.getAttribute('style') ?? '';
          const w = /width:([\d.]+)pt/.exec(style);
          const h = /height:([\d.]+)pt/.exec(style);
          acc.nodes.push(schema.nodes.image.create({
            src: bytesToDataUrl(bytes, MIME[ext]), alt: '',
            width: w ? Math.round(parseFloat(w[1]) * 4 / 3) : null, height: h ? Math.round(parseFloat(h[1]) * 4 / 3) : null,
          }, null, marks));
        } else lose(ctx, 'фигуры, надписи и вложенные объекты');
        break;
      }
      case 'AlternateContent': {
        const choice = kids(c, 'Choice')[0];
        if (choice) for (const d of kids(choice, 'drawing')) parseDrawing(ctx, d, acc, marks);
        break;
      }
      case 'footnoteReference':
      case 'endnoteReference':
        lose(ctx, 'сноски');
        break;
      case 'commentReference':
        lose(ctx, 'примечания');
        break;
      default:
        break;
    }
  }
}

function walkInline(ctx: Ctx, el: Element, acc: InlineAcc, href: string | undefined): void {
  for (const c of el.children) {
    switch (c.localName) {
      case 'r':
        parseRun(ctx, c, acc, href, !href);
        break;
      case 'hyperlink': {
        const id = relAttr(c, 'id');
        const anchor = attr(c, 'anchor');
        const target = id ? ctx.rels.get(id)?.target : anchor ? `#${anchor}` : undefined;
        walkInline(ctx, c, acc, target ?? href);
        break;
      }
      case 'fldSimple': {
        const instr = attr(c, 'instr') ?? '';
        const m = /^\s*HYPERLINK\s+"([^"]+)"/i.exec(instr);
        if (!m) lose(ctx, 'поля (оглавление, номера страниц и т. п. стали обычным текстом)');
        walkInline(ctx, c, acc, m ? m[1] : href);
        break;
      }
      case 'ins':
        lose(ctx, 'исправления (приняты как есть)');
        walkInline(ctx, c, acc, href);
        break;
      case 'del':
        lose(ctx, 'исправления (приняты как есть)');
        break;
      case 'sdt':
        walkInline(ctx, kid(c, 'sdtContent') ?? c, acc, href);
        break;
      case 'smartTag':
      case 'sdtContent':
      case 'customXml':
        walkInline(ctx, c, acc, href);
        break;
      case 'oMath':
      case 'oMathPara':
        lose(ctx, 'формулы');
        break;
      default:
        break;
    }
  }
}

interface Entry { node: PMNode; list?: { key: number; level: number; info: ListInfo } }

const HEADING_RE = /^(?:heading|заголовок)\s*(\d)$/i;

function headingLevel(ctx: Ctx, styleId: string | undefined): number | null {
  if (!styleId) return null;
  const s = ctx.styles.get(styleId);
  const m = HEADING_RE.exec(s?.name ?? styleId.replace(/(\d)$/, ' $1'));
  return m ? Math.min(6, Math.max(1, +m[1])) : null;
}

function paraEntries(ctx: Ctx, p: Element): Entry[] {
  const pPr = kid(p, 'pPr');
  if (kid(pPr, 'sectPr')) lose(ctx, 'разделы (разрывы разделов убраны)');
  const direct = parsePPr(pPr);
  const styleId = direct.styleId ?? ctx.defaultPara?.id;
  const level = headingLevel(ctx, direct.styleId);
  // Стиль заголовка рисует сама программа; остальные стили переносятся прямым форматированием.
  const fromStyle = level ? { p: {}, r: {} } : resolveStyle(ctx, styleId);
  const para: PFmt = { ...fromStyle.p, ...direct };
  const info = listInfo(ctx, para.numId, para.ilvl ?? 0);
  const d = ctx.defaults;

  const acc: InlineAcc = { nodes: [], chunks: [], breakNext: false };
  ctx.baseR = fromStyle.r;
  walkInline(ctx, p, acc, undefined);
  ctx.baseR = {};
  acc.chunks.push({ nodes: acc.nodes, breakBefore: acc.breakNext });

  const entries: Entry[] = [];
  acc.chunks.forEach((chunk, k) => {
    // Абзац, состоящий только из разрыва страницы, не оставляет пустого абзаца до разрыва.
    if (!chunk.nodes.length && !chunk.breakBefore && k < acc.chunks.length - 1) return;
    const attrs = {
      align: para.align && para.align !== 'left' ? para.align : null,
      indent: !info && para.left ? para.left : null,
      first: !info && para.first ? para.first : null,
      before: para.before != null && para.before !== d.before ? para.before : null,
      after: para.after != null && para.after !== d.after ? para.after : null,
      line: para.line && para.line !== d.line ? para.line : null,
      pageBreakBefore: !!(chunk.breakBefore || (k === 0 && para.pageBreakBefore)),
      ...(level ? { level } : {}),
    };
    entries.push({
      node: schema.nodes[level ? 'heading' : 'paragraph'].create(attrs, chunk.nodes),
      list: info ? { key: para.numId ?? 0, level: para.ilvl ?? 0, info } : undefined,
    });
  });
  return entries;
}

/** Склеивает подряд идущие абзацы-элементы списка во вложенные списки. */
function assemble(entries: Entry[]): PMNode[] {
  const out: PMNode[] = [];
  interface Open { key: number; level: number; kind: 'bullet' | 'ordered'; fmt: string; items: PMNode[][] }
  const stack: Open[] = [];
  const build = (o: Open): PMNode => {
    const items = o.items.map((c) => schema.nodes.list_item.create(null, c));
    return o.kind === 'bullet'
      ? schema.nodes.bullet_list.create(null, items)
      : schema.nodes.ordered_list.create({ order: 1, fmt: orderedFmt(o.fmt) }, items);
  };
  const close = () => {
    const o = stack.pop()!;
    const node = build(o);
    const parent = stack[stack.length - 1];
    if (parent) {
      if (!parent.items.length) parent.items.push([schema.nodes.paragraph.create()]);
      parent.items[parent.items.length - 1].push(node);
    } else out.push(node);
  };
  for (const e of entries) {
    if (!e.list) {
      while (stack.length) close();
      out.push(e.node);
      continue;
    }
    const { key, level, info } = e.list;
    while (stack.length && stack[stack.length - 1].level > level) close();
    let top = stack[stack.length - 1];
    if (top && top.level === level && (top.key !== key || top.kind !== info.kind)) {
      close();
      top = stack[stack.length - 1];
    }
    if (!top || top.level < level) {
      stack.push({ key, level, kind: info.kind, fmt: info.fmt, items: [] });
      top = stack[stack.length - 1];
    }
    // Заголовок внутри списка в элемент списка не кладём: список закрывается.
    if (e.node.type === schema.nodes.heading) {
      while (stack.length) close();
      out.push(e.node);
      continue;
    }
    top.items.push([e.node]);
  }
  while (stack.length) close();
  return out;
}

function tableHasBorders(ctx: Ctx, tbl: Element): boolean {
  const pr = kid(tbl, 'tblPr');
  const b = kid(pr, 'tblBorders');
  if (b) return kids(b).some((s) => !['nil', 'none'].includes(attr(s, 'val') ?? 'none'));
  const style = attr(kid(pr, 'tblStyle'), 'val');
  const name = style ? ctx.styles.get(style)?.name ?? style : '';
  if (/grid|сетка/i.test(name) || /grid|сетка/i.test(style ?? '')) return true;
  return descendants(tbl, 'tcBorders').length > 0;
}

function parseTable(ctx: Ctx, tbl: Element): PMNode {
  const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((c) => (num(attr(c, 'w')) ?? 0) / 15);
  interface CellB { span: number; rowspan: number; widths: number[] | null; bg: string | null; content: PMNode[] }
  const rows: CellB[][] = [];
  const active: (CellB | null)[] = [];
  for (const tr of kids(tbl, 'tr')) {
    const cells: CellB[] = [];
    let g = 0;
    for (const tc of kids(tr, 'tc')) {
      const pr = kid(tc, 'tcPr');
      const span = Math.max(1, num(attr(kid(pr, 'gridSpan'), 'val')) ?? 1);
      const vm = kid(pr, 'vMerge');
      const restart = attr(vm, 'val') === 'restart';
      if (vm && !restart) {
        const origin = active[g];
        if (origin) origin.rowspan++;
        g += span;
        continue;
      }
      let content = walkBlocks(ctx, tc);
      if (!content.length) content = [schema.nodes.paragraph.create()];
      const widths = grid.length >= g + span ? grid.slice(g, g + span).map((w) => Math.round(w)) : null;
      const fill = attr(kid(pr, 'shd'), 'fill');
      const cb: CellB = { span, rowspan: 1, widths: widths && widths.every((w) => w > 0) ? widths : null, bg: fill && fill !== 'auto' && /^[0-9a-f]{6}$/i.test(fill) ? `#${fill}` : null, content };
      cells.push(cb);
      for (let k = 0; k < span; k++) active[g + k] = restart ? cb : null;
      g += span;
    }
    rows.push(cells);
  }
  const mk = () => schema.nodes.table.create({ borders: tableHasBorders(ctx, tbl) },
    rows.map((cells) => schema.nodes.table_row.create(null, cells.map((c) =>
      schema.nodes.table_cell.create({ colspan: c.span, rowspan: c.rowspan, colwidth: c.widths, background: c.bg }, c.content)))));
  const t = mk();
  t.check();
  return t;
}

function walkBlocks(ctx: Ctx, parent: Element): PMNode[] {
  const entries: Entry[] = [];
  const visit = (el: Element) => {
    for (const c of el.children) {
      switch (c.localName) {
        case 'p':
          entries.push(...paraEntries(ctx, c));
          break;
        case 'tbl':
          try {
            entries.push({ node: parseTable(ctx, c) });
          } catch {
            lose(ctx, 'таблицы сложной структуры (преобразованы в текст)');
            for (const tr of descendants(c, 'tr')) {
              const text = kids(tr, 'tc').map((tc) => tc.textContent ?? '').join('\t');
              entries.push({ node: schema.nodes.paragraph.create(null, text ? schema.text(text) : undefined) });
            }
          }
          break;
        case 'sdt':
          visit(kid(c, 'sdtContent') ?? c);
          break;
        case 'customXml':
        case 'sdtContent':
          visit(c);
          break;
        default:
          break;
      }
    }
  };
  visit(parent);
  return assemble(entries);
}

export interface Imported { doc: PMNode; meta: DocMeta }

export function importDocx(bytes: Uint8Array): Imported {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch {
    throw new Error('Файл не похож на документ docx (это не ZIP-архив)');
  }
  const ct = files['[Content_Types].xml'];
  if (!ct) throw new Error('Файл не похож на документ docx (нет [Content_Types].xml)');
  let main = 'word/document.xml';
  const rootRels = files['_rels/.rels'] ? strFromU8(files['_rels/.rels']) : '';
  const m = /Type="[^"]*\/officeDocument"[^>]*Target="\/?([^"]+)"|Target="\/?([^"]+)"[^>]*Type="[^"]*\/officeDocument"/.exec(rootRels);
  if (m) main = m[1] ?? m[2];
  if (!files[main]) throw new Error('В файле нет основного документа Word (word/document.xml)');
  if (/spreadsheetml|presentationml/.test(strFromU8(ct)) && !/wordprocessingml/.test(strFromU8(ct))) {
    throw new Error('Это не документ Word: файл содержит таблицу или презентацию');
  }

  const dir = main.slice(0, main.lastIndexOf('/') + 1);
  const ctx: Ctx = {
    files, rels: readRels(files, main), styles: new Map(), defaults: { ...DEFAULTS }, nums: new Map(),
    lossy: new Map(), baseR: {}, fields: [],
  };
  for (const [, rel] of ctx.rels) {
    if (/\/(header|footer)\d*$/.test(rel.type) && files[rel.target]) {
      const text = descendants(parseXml(strFromU8(files[rel.target])).documentElement, 't').map((t) => t.textContent).join('').trim();
      if (text) lose(ctx, 'колонтитулы');
    }
    if (/\/comments$/.test(rel.type)) lose(ctx, 'примечания');
    if (/\/(footnotes|endnotes)$/.test(rel.type)) { /* считаем по ссылкам в тексте */ }
  }
  themeFonts = { minor: '', major: '' };
  const themePath = [...ctx.rels.values()].find((r) => /\/theme$/.test(r.type))?.target;
  if (themePath && files[themePath]) {
    const th = parseXml(strFromU8(files[themePath])).documentElement;
    const face = (n: string) => descendants(th, n)[0] ? descendants(descendants(th, n)[0], 'latin')[0]?.getAttribute('typeface') ?? '' : '';
    themeFonts = { minor: face('minorFont'), major: face('majorFont') };
  }
  const stylesPath = [...ctx.rels.values()].find((r) => /\/styles$/.test(r.type))?.target ?? `${dir}styles.xml`;
  const numPath = [...ctx.rels.values()].find((r) => /\/numbering$/.test(r.type))?.target ?? `${dir}numbering.xml`;
  // Двойной сброс: у Normal может не быть своих значений — берём значения по умолчанию программы.
  readStyles(files[stylesPath] ? strFromU8(files[stylesPath]) : undefined, ctx);
  readNumbering(files[numPath] ? strFromU8(files[numPath]) : undefined, ctx);

  const root = parseXml(strFromU8(files[main])).documentElement;
  const body = kid(root, 'body');
  if (!body) throw new Error('Документ Word пуст или повреждён (нет w:body)');
  let blocks = walkBlocks(ctx, body);
  if (!blocks.length) blocks = [schema.nodes.paragraph.create()];
  const last = blocks[blocks.length - 1];
  if (last.type === schema.nodes.table) blocks.push(schema.nodes.paragraph.create());

  const page = { ...A4 };
  const sect = kid(body, 'sectPr');
  if (sect) {
    const sz = kid(sect, 'pgSz');
    const mar = kid(sect, 'pgMar');
    let w = (num(attr(sz, 'w')) ?? 0) / 20;
    let h = (num(attr(sz, 'h')) ?? 0) / 20;
    if (w > 100 && h > 100) {
      if (attr(sz, 'orient') === 'landscape' && w < h) [w, h] = [h, w];
      page.w = w;
      page.h = h;
    }
    const mv = (n: string, fallback: number) => { const v = num(attr(mar, n)); return v != null && v >= 0 ? v / 20 : fallback; };
    page.top = mv('top', page.top);
    page.bottom = mv('bottom', page.bottom);
    page.left = mv('left', page.left);
    page.right = mv('right', page.right);
    if (kids(sect, 'headerReference').length || kids(sect, 'footerReference').length) {
      // Учтено выше только если колонтитул непустой; пустые не считаем потерей.
    }
    if (kid(sect, 'cols') && (num(attr(kid(sect, 'cols'), 'num')) ?? 1) > 1) lose(ctx, 'колонки (текст идёт в одну колонку)');
  }
  const lossy = [...ctx.lossy].map(([what, n]) => (n > 1 ? `${what}: ${n}` : what));
  return { doc: schema.nodes.doc.create(null, blocks), meta: { page, defaults: ctx.defaults, lossy } };
}

// ---------------------------------------------------------------- экспорт

const esc = (s: string) => s.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const tw = (pt: number) => Math.round(pt * 20);

interface ExportCtx {
  meta: DocMeta;
  rels: { id: string; type: string; target: string; external: boolean }[];
  media: { name: string; bytes: Uint8Array; ext: string }[];
  nums: { abstract: string; num: string }[];
  nextRel: number;
  nextDoc: number;
  linkIds: Map<string, string>;
}

function relId(ctx: ExportCtx, type: string, target: string, external = false): string {
  const id = `rId${ctx.nextRel++}`;
  ctx.rels.push({ id, type, target, external });
  return id;
}

function rPrXml(marks: readonly Mark[], link: boolean): string {
  let font: string | null = null, size: number | null = null, color: string | null = null, bg: string | null = null;
  let b = false, i = false, u = false, s = false, sup = false, sub = false;
  for (const m of marks) {
    switch (m.type.name) {
      case 'strong': b = true; break;
      case 'em': i = true; break;
      case 'underline': u = true; break;
      case 'strike': s = true; break;
      case 'sup': sup = true; break;
      case 'sub': sub = true; break;
      case 'textStyle': ({ font, size, color, bg } = m.attrs as { font: string | null; size: number | null; color: string | null; bg: string | null }); break;
      default: break;
    }
  }
  let x = link ? '<w:rStyle w:val="Hyperlink"/>' : '';
  if (font) x += `<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:cs="${esc(font)}"/>`;
  if (b) x += '<w:b/><w:bCs/>';
  if (i) x += '<w:i/><w:iCs/>';
  if (s) x += '<w:strike/>';
  if (color) x += `<w:color w:val="${color.replace('#', '').toUpperCase()}"/>`;
  if (size) x += `<w:sz w:val="${Math.round(size * 2)}"/><w:szCs w:val="${Math.round(size * 2)}"/>`;
  if (u) x += '<w:u w:val="single"/>';
  if (bg) x += `<w:shd w:val="clear" w:color="auto" w:fill="${bg.replace('#', '').toUpperCase()}"/>`;
  if (sup) x += '<w:vertAlign w:val="superscript"/>';
  if (sub) x += '<w:vertAlign w:val="subscript"/>';
  return x ? `<w:rPr>${x}</w:rPr>` : '';
}

function imageXml(ctx: ExportCtx, node: PMNode): string {
  const m = /^data:([^;,]+);base64,(.*)$/s.exec(node.attrs.src as string);
  if (!m) return '';
  const ext = Object.entries(MIME).find(([, v]) => v === m[1])?.[0] ?? 'png';
  const bin = atob(m[2]);
  const bytes = new Uint8Array(bin.length);
  for (let k = 0; k < bin.length; k++) bytes[k] = bin.charCodeAt(k);
  const name = `image${ctx.media.length + 1}.${ext === 'jpeg' ? 'jpg' : ext}`;
  ctx.media.push({ name, bytes, ext: ext === 'jpeg' ? 'jpg' : ext });
  const id = relId(ctx, REL_IMAGE, `media/${name}`);
  const cx = Math.round((node.attrs.width || 200) * 9525);
  const cy = Math.round((node.attrs.height || 150) * 9525);
  const n = ctx.nextDoc++;
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${n}" name="Рисунок ${n}" descr="${esc(node.attrs.alt || '')}"/>` +
    '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
    `<pic:nvPicPr><pic:cNvPr id="${n}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
}

function runsXml(ctx: ExportCtx, para: PMNode): string {
  let out = '';
  let openHref: string | null = null;
  const closeLink = () => { if (openHref !== null) { out += '</w:hyperlink>'; openHref = null; } };
  para.forEach((child) => {
    const link = child.marks.find((m) => m.type.name === 'link');
    const href = link ? (link.attrs.href as string) : null;
    if (href !== openHref) {
      closeLink();
      if (href !== null) {
        if (href.startsWith('#')) out += `<w:hyperlink w:anchor="${esc(href.slice(1))}" w:history="1">`;
        else {
          let id = ctx.linkIds.get(href);
          if (!id) { id = relId(ctx, REL_HYPERLINK, href, true); ctx.linkIds.set(href, id); }
          out += `<w:hyperlink r:id="${id}" w:history="1">`;
        }
        openHref = href;
      }
    }
    const rPr = rPrXml(child.marks, !!link);
    if (child.type.name === 'text') {
      // Табуляции и переводы строк внутри текста — отдельные элементы.
      let buf = '';
      const flush = () => { if (buf) { out += `<w:r>${rPr}<w:t xml:space="preserve">${esc(buf)}</w:t></w:r>`; buf = ''; } };
      for (const ch of child.text ?? '') {
        if (ch === '\t') { flush(); out += `<w:r>${rPr}<w:tab/></w:r>`; }
        else if (ch === '\n') { flush(); out += `<w:r>${rPr}<w:br/></w:r>`; }
        else buf += ch;
      }
      flush();
    } else if (child.type.name === 'hard_break') {
      out += `<w:r>${rPr}<w:br/></w:r>`;
    } else if (child.type.name === 'image') {
      out += imageXml(ctx, child);
    }
  });
  closeLink();
  return out;
}

function pPrXml(node: PMNode, extra: { num?: { id: number; level: number }; indentLeft?: number } = {}): string {
  const a = node.attrs;
  let x = '';
  if (node.type.name === 'heading') x += `<w:pStyle w:val="Heading${a.level}"/>`;
  if (a.pageBreakBefore) x += '<w:pageBreakBefore/>';
  if (extra.num) x += `<w:numPr><w:ilvl w:val="${extra.num.level}"/><w:numId w:val="${extra.num.id}"/></w:numPr>`;
  if (a.before != null || a.after != null || a.line) {
    x += '<w:spacing' + (a.before != null ? ` w:before="${tw(a.before)}"` : '') + (a.after != null ? ` w:after="${tw(a.after)}"` : '') +
      (a.line ? ` w:line="${Math.round(a.line * 240)}" w:lineRule="auto"` : '') + '/>';
  }
  const left = (a.indent ?? 0) + (extra.indentLeft ?? 0);
  if (left || a.first) x += `<w:ind w:left="${tw(left)}"` + (a.first ? (a.first < 0 ? ` w:hanging="${tw(-a.first)}"` : ` w:firstLine="${tw(a.first)}"`) : '') + '/>';
  if (a.align) x += `<w:jc w:val="${a.align === 'justify' ? 'both' : a.align}"/>`;
  return x ? `<w:pPr>${x}</w:pPr>` : '';
}

function paragraphXml(ctx: ExportCtx, node: PMNode, extra?: Parameters<typeof pPrXml>[1]): string {
  return `<w:p>${pPrXml(node, extra)}${runsXml(ctx, node)}</w:p>`;
}

/** Определения нумерации одного верхнего списка: уровни со своим форматом. */
function numberingFor(ctx: ExportCtx, list: PMNode): number {
  const levels: { kind: string; fmt: string }[] = [];
  const scan = (l: PMNode, level: number) => {
    levels[level] ??= l.type.name === 'bullet_list' ? { kind: 'bullet', fmt: 'bullet' } : { kind: 'ordered', fmt: orderedFmt(l.attrs.fmt as string) };
    l.forEach((item) => item.forEach((c) => { if (c.type.name === 'bullet_list' || c.type.name === 'ordered_list') scan(c, level + 1); }));
  };
  scan(list, 0);
  const id = ctx.nums.length + 1;
  const bullets = ['•', '◦', '▪'];
  let lvls = '';
  for (let k = 0; k < 9; k++) {
    const l = levels[k] ?? levels[levels.length - 1] ?? { kind: 'bullet', fmt: 'bullet' };
    const left = 720 * (k + 1);
    if (l.kind === 'bullet') {
      lvls += `<w:lvl w:ilvl="${k}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${bullets[k % 3]}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:hint="default"/></w:rPr></w:lvl>`;
    } else {
      lvls += `<w:lvl w:ilvl="${k}"><w:start w:val="1"/><w:numFmt w:val="${l.fmt}"/><w:lvlText w:val="%${k + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`;
    }
  }
  ctx.nums.push({
    abstract: `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${lvls}</w:abstractNum>`,
    num: `<w:num w:numId="${id}"><w:abstractNumId w:val="${id}"/></w:num>`,
  });
  return id;
}

function listXml(ctx: ExportCtx, list: PMNode, numId: number, level: number): string {
  let out = '';
  list.forEach((item) => {
    let firstPara = true;
    item.forEach((c) => {
      if (c.type.name === 'paragraph' || c.type.name === 'heading') {
        out += paragraphXml(ctx, c, firstPara ? { num: { id: numId, level } } : { indentLeft: 36 * (level + 1) });
        firstPara = false;
      } else if (c.type.name === 'bullet_list' || c.type.name === 'ordered_list') {
        out += listXml(ctx, c, numId, level + 1);
      } else if (c.type.name === 'table') {
        out += tableXml(ctx, c);
      }
    });
  });
  return out;
}

function tableXml(ctx: ExportCtx, table: PMNode): string {
  const map = TableMap.get(table);
  const usable = (ctx.meta.page.w - ctx.meta.page.left - ctx.meta.page.right) * 20;
  const widths: number[] = new Array(map.width).fill(0);
  // Ширины столбцов: из colwidth любой ячейки, иначе поровну.
  for (let r = 0; r < map.height; r++) {
    for (let c = 0; c < map.width; c++) {
      const pos = map.map[r * map.width + c];
      const cell = table.nodeAt(pos);
      const cw = cell?.attrs.colwidth as number[] | null;
      if (cw && !widths[c]) {
        const start = map.colCount(pos);
        widths[c] = (cw[c - start] ?? 0) * 15;
      }
    }
  }
  const known = widths.filter((w) => w > 0).reduce((a, b) => a + b, 0);
  const unknown = widths.filter((w) => !w).length;
  const fill = unknown ? Math.max(720, Math.round((usable - known) / unknown)) : 0;
  for (let c = 0; c < widths.length; c++) if (!widths[c]) widths[c] = fill || Math.round(usable / map.width);
  const total = widths.reduce((a, b) => a + b, 0);

  const borders = table.attrs.borders
    ? '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') + '</w:tblBorders>'
    : '';
  let out = `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/>${borders}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>` +
    widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('') + '</w:tblGrid>';
  for (let r = 0; r < map.height; r++) {
    out += '<w:tr>';
    for (let c = 0; c < map.width; c++) {
      const pos = map.map[r * map.width + c];
      const rect = map.findCell(pos);
      if (rect.left !== c) continue; // середина объединённой ячейки
      const cell = table.nodeAt(pos)!;
      const span = rect.right - rect.left;
      const cont = rect.top !== r;
      const w = widths.slice(rect.left, rect.right).reduce((a, b) => a + b, 0);
      let pr = `<w:tcW w:w="${w}" w:type="dxa"/>`;
      if (span > 1) pr += `<w:gridSpan w:val="${span}"/>`;
      if (rect.bottom - rect.top > 1) pr += cont ? '<w:vMerge/>' : '<w:vMerge w:val="restart"/>';
      if (cell.attrs.background && !cont) pr += `<w:shd w:val="clear" w:color="auto" w:fill="${(cell.attrs.background as string).replace('#', '')}"/>`;
      let inner = '';
      if (!cont) inner = blocksXml(ctx, cell);
      if (!inner || !inner.endsWith('</w:p>')) inner += '<w:p/>';
      out += `<w:tc><w:tcPr>${pr}</w:tcPr>${inner}</w:tc>`;
    }
    out += '</w:tr>';
  }
  return out + '</w:tbl>';
}

function blocksXml(ctx: ExportCtx, parent: PMNode): string {
  let out = '';
  parent.forEach((n) => {
    switch (n.type.name) {
      case 'paragraph':
      case 'heading':
        out += paragraphXml(ctx, n);
        break;
      case 'bullet_list':
      case 'ordered_list':
        out += listXml(ctx, n, numberingFor(ctx, n), 0);
        break;
      case 'table':
        out += tableXml(ctx, n);
        break;
      default:
        break;
    }
  });
  return out;
}

function stylesXml(meta: DocMeta): string {
  const d = meta.defaults;
  const f = esc(d.font);
  const heading = (n: number, size: number, extra = '') =>
    `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>` +
    `<w:pPr><w:keepNext/><w:spacing w:before="${n === 1 ? 240 : 200}" w:after="120"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:bCs/>${extra}<w:sz w:val="${size * 2}"/><w:szCs w:val="${size * 2}"/></w:rPr></w:style>`;
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:styles xmlns:w="${NS.w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:eastAsia="${f}" w:cs="${f}"/>` +
    `<w:sz w:val="${Math.round(d.size * 2)}"/><w:szCs w:val="${Math.round(d.size * 2)}"/><w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:spacing w:after="${tw(d.after)}" w:line="${Math.round(d.line * 240)}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
    heading(1, 18) + heading(2, 16) + heading(3, 14) + heading(4, 13) + heading(5, 12) + heading(6, 12) +
    '</w:styles>';
}

export function exportDocx(doc: PMNode, meta: DocMeta, title = ''): Uint8Array {
  const ctx: ExportCtx = { meta, rels: [], media: [], nums: [], nextRel: 10, nextDoc: 1, linkIds: new Map() };
  let body = blocksXml(ctx, doc);
  if (body.endsWith('</w:tbl>')) body += '<w:p/>';
  const p = meta.page;
  const sect = `<w:sectPr><w:pgSz w:w="${tw(p.w)}" w:h="${tw(p.h)}"${p.w > p.h ? ' w:orient="landscape"' : ''}/>` +
    `<w:pgMar w:top="${tw(p.top)}" w:right="${tw(p.right)}" w:bottom="${tw(p.bottom)}" w:left="${tw(p.left)}" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;
  const docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    `<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<w:body>${body}${sect}</w:body></w:document>`;

  const base = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const fixed = [
    { id: 'rId1', type: `${base}/styles`, target: 'styles.xml', external: false },
    { id: 'rId2', type: `${base}/settings`, target: 'settings.xml', external: false },
    ...(ctx.nums.length ? [{ id: 'rId3', type: `${base}/numbering`, target: 'numbering.xml', external: false }] : []),
  ];
  const relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    [...fixed, ...ctx.rels].map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('') + '</Relationships>';

  const exts = new Set(ctx.media.map((m) => m.ext));
  const ctypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    [...exts].map((e) => `<Default Extension="${e}" ContentType="${MIME[e] ?? 'application/octet-stream'}"/>`).join('') +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>' +
    (ctx.nums.length ? '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' : '') +
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>';

  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const files: Zippable = {
    '[Content_Types].xml': strToU8(ctypes),
    '_rels/.rels': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${base}/officeDocument" Target="word/document.xml"/>` +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      `<Relationship Id="rId3" Type="${base}/extended-properties" Target="docProps/app.xml"/></Relationships>`),
    'word/document.xml': strToU8(docXml),
    'word/_rels/document.xml.rels': strToU8(relsXml),
    'word/styles.xml': strToU8(stylesXml(meta)),
    'word/settings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="${NS.w}"><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`),
    'docProps/core.xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      `<dc:title>${esc(title)}</dc:title><dc:creator>ENS-Tekst</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`),
    'docProps/app.xml': strToU8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>ENS-Tekst</Application></Properties>'),
  };
  if (ctx.nums.length) {
    files['word/numbering.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${NS.w}">${ctx.nums.map((n) => n.abstract).join('')}${ctx.nums.map((n) => n.num).join('')}</w:numbering>`);
  }
  for (const m of ctx.media) files[`word/media/${m.name}`] = [m.bytes, { level: 0 }];
  return zipSync(files, { level: 6 });
}
