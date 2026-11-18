// SPDX-License-Identifier: AGPL-3.0-or-later
// Список литературы по ГОСТ Р 7.0.100-2018 (библиографическая запись) из карточек источников, без интернета.
//
// Карточки лежат в атрибуте документа `sources` (отмена и версии файла работают как с текстом). Ссылка в тексте —
// узел `cite` («[1, с. 25]», «[2; 5]»): номер источника — по порядку первого упоминания, пересчитывается сам.
// Узел `biblio` — сам список: строки строятся по карточкам при каждой правке. Лучше, чем в Word: оформление
// по ГОСТ, а не по APA; источники, на которые нет ссылок в тексте, отмечены.

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet, type EditorView, type NodeView } from 'prosemirror-view';
import { schema } from './schema';

export type SourceKind = 'book' | 'article' | 'web' | 'law' | 'other';
export interface Source {
  id: number;
  kind: SourceKind;
  /** Авторы: «Иванов И. И.» по одному в строке (порядок как на титуле). */
  authors: string[];
  title: string;
  /** Сведения, относящиеся к заглавию: «учебник», «монография», «федеральный закон от … № …». */
  info?: string;
  /** Журнал или сборник (для статьи), сайт (для электронного ресурса). */
  container?: string;
  edition?: string;
  city?: string;
  publisher?: string;
  year?: string;
  /** Номер журнала, том. */
  issue?: string;
  /** Объём («256 с.») или страницы статьи («15–20»). */
  pages?: string;
  url?: string;
  /** Дата обращения (ДД.ММ.ГГГГ). */
  accessed?: string;
  isbn?: string;
  /** Готовая запись для вида «другое». */
  text?: string;
}

export const KIND_LABEL: Record<SourceKind, string> = { book: 'Книга', article: 'Статья', web: 'Электронный ресурс', law: 'Нормативный акт', other: 'Другое (готовая запись)' };

/** «Иванов И. И.» → заголовок записи «Иванов, И. И.» и сведения об ответственности «И. И. Иванов». */
function nameParts(a: string): { head: string; resp: string } {
  const m = /^\s*([^\s,]+),?\s+(.+?)\s*$/.exec(a);
  if (!m) return { head: a.trim(), resp: a.trim() };
  return { head: `${m[1]}, ${m[2]}`, resp: `${m[2]} ${m[1]}` };
}
const end = (s: string) => (/[.!?]$/.test(s) ? s : s + '.');
/** Следующая область записи: «. – » (точка не удваивается, если запись уже кончается точкой). */
const area = (r: string, part: string) => `${end(r)} – ${part}`;

/** Библиографическая запись по ГОСТ Р 7.0.100-2018 (упрощённо: основные области для частых видов). */
export function formatSource(s: Source): string {
  if (s.kind === 'other') return (s.text ?? '').trim();
  const authors = s.authors.map((a) => a.trim()).filter(Boolean);
  // До трёх авторов — запись под заголовком (первый автор); четыре и более — под заглавием, в ответственности «[и др.]».
  const head = authors.length && authors.length <= 3 ? nameParts(authors[0]).head + ' ' : '';
  const resp = authors.length ? ' / ' + (authors.length <= 3 ? authors.map((a) => nameParts(a).resp).join(', ') : authors.slice(0, 3).map((a) => nameParts(a).resp).join(', ') + ' [и др.]') : '';
  let r = head + s.title.trim() + (s.info ? ` : ${s.info.trim()}` : '') + resp;
  if (s.kind === 'article') {
    r += ` // ${(s.container ?? '').trim()}`;
    if (s.year) r = area(r, `${s.year}`);
    if (s.issue) r = area(r, `№ ${s.issue.replace(/^№\s*/, '')}`);
    if (s.pages) r = area(r, `С. ${s.pages.replace(/^С\.\s*/i, '')}`);
    return end(r);
  }
  if (s.kind === 'web') {
    if (s.container) r += ` // ${s.container.trim()}`;
    if (s.year) r = area(r, `${s.year}`);
    if (s.url) r = area(r, `URL: ${s.url.trim()}${s.accessed ? ` (дата обращения: ${s.accessed.trim()})` : ''}`);
    return end(r);
  }
  if (s.kind === 'law') {
    if (s.container) r += ` // ${s.container.trim()}`;
    if (s.year) r = area(r, `${s.year}`);
    if (s.issue) r = area(r, `№ ${s.issue.replace(/^№\s*/, '')}`);
    if (s.url) r = area(r, `URL: ${s.url.trim()}${s.accessed ? ` (дата обращения: ${s.accessed.trim()})` : ''}`);
    return end(r);
  }
  if (s.edition) r = area(r, `${end(s.edition.trim())}`);
  const pub = [s.city?.trim(), s.publisher?.trim()].filter(Boolean).join(' : ');
  if (pub || s.year) r = area(r, `${pub}${pub && s.year ? ', ' : ''}${s.year ?? ''}`);
  if (s.pages) r = area(r, `${/(^|\s)с\.?$/.test(s.pages.trim()) ? s.pages.trim() : s.pages.trim() + ' с'}`);
  if (s.isbn) r = area(r, `ISBN ${s.isbn.replace(/^ISBN\s*/i, '')}`);
  return end(r);
}

export const sourcesOf = (doc: PMNode): Source[] => (doc.attrs.sources as Source[] | undefined) ?? [];

/** Номера источников по порядку первого упоминания; источники без ссылок — после них. */
export function numbering(doc: PMNode): { order: Source[]; num: Map<number, number>; cited: Set<number> } {
  const all = sourcesOf(doc);
  const byId = new Map(all.map((s) => [s.id, s]));
  const order: Source[] = [];
  const cited = new Set<number>();
  doc.descendants((node) => {
    if (node.type !== schema.nodes.cite) return true;
    for (const id of node.attrs.ids as number[]) if (byId.has(id) && !cited.has(id)) { cited.add(id); order.push(byId.get(id)!); }
    return false;
  });
  for (const s of all) if (!cited.has(s.id)) order.push(s);
  return { order, num: new Map(order.map((s, i) => [s.id, i + 1])), cited };
}

/** Текст ссылки: «[1, с. 25]», «[2; 5]», «[?]» — источник удалён. */
export function citeLabel(node: PMNode, num: Map<number, number>): string {
  const ids = node.attrs.ids as number[];
  const nums = ids.map((id) => num.get(id));
  if (nums.some((n) => n == null)) return '[?]';
  const page = (node.attrs.page as string).trim();
  return `[${(nums as number[]).join('; ')}${page ? `, с. ${page.replace(/^с\.\s*/i, '')}` : ''}]`;
}

function build(doc: PMNode): DecorationSet {
  const decos: Decoration[] = [];
  let num: Map<number, number> | null = null;
  doc.descendants((node, pos) => {
    if (node.type !== schema.nodes.cite) return true;
    num ??= numbering(doc).num;
    decos.push(Decoration.node(pos, pos + 1, { 'data-label': citeLabel(node, num) }));
    return false;
  });
  return decos.length ? DecorationSet.create(doc, decos) : DecorationSet.empty;
}

export const biblioPlugin: Plugin<DecorationSet> = new Plugin<DecorationSet>({
  state: {
    init: (_c, state) => build(state.doc),
    apply: (tr, prev, _o, state) => (tr.docChanged ? build(state.doc) : prev),
  },
  props: { decorations(state) { return biblioPlugin.getState(state); } },
});

/** Список литературы на листе: строки по карточкам, перерисовка при каждой правке документа. */
class BiblioView implements NodeView {
  dom: HTMLElement;
  private sig = '';
  constructor(private view: EditorView) {
    this.dom = document.createElement('div');
    this.dom.className = 'biblio';
    this.dom.contentEditable = 'false';
    live.add(this);
    this.render();
  }
  render(): void {
    const { order, cited } = numbering(this.view.state.doc);
    const rows = order.map((s, i) => [i + 1, formatSource(s), cited.has(s.id)] as const);
    const sig = JSON.stringify(rows);
    if (sig === this.sig) return;
    const grew = this.sig !== '' && JSON.parse(this.sig).length !== rows.length;
    this.sig = sig;
    this.dom.replaceChildren(...(rows.length ? rows.map(([n, text, used]) => {
      const p = document.createElement('p');
      p.className = `biblio-row${used ? '' : ' uncited'}`;
      p.textContent = `${n}. ${text}`;
      if (!used) p.title = 'На этот источник нет ссылок в тексте';
      return p;
    }) : [Object.assign(document.createElement('p'), { className: 'biblio-empty', textContent: 'Список пуст: добавьте источники (Ссылки → Источники) и ссылки на них в тексте.' })]));
    if (grew) onResize();
  }
  update(node: PMNode): boolean { return node.type === schema.nodes.biblio; }
  stopEvent(): boolean { return false; }
  ignoreMutation(): boolean { return true; }
  destroy(): void { live.delete(this); }
}
const live = new Set<BiblioView>();
let onResize: () => void = () => {};
export function setBiblioResize(f: () => void): void { onResize = f; }
export const biblioViews = { biblio: (_n: PMNode, view: EditorView) => new BiblioView(view) };
export function refreshBiblio(): void { for (const v of live) v.render(); }
