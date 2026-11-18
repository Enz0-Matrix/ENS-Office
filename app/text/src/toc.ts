// SPDX-License-Identifier: AGPL-3.0-or-later
// Оглавление (Ссылки → Оглавление). Лучше, чем в Word: строки и номера страниц обновляются сами при каждой
// правке и перестройке страниц, без «Обновить поле»; щелчок по строке ведёт к заголовку.
//
// Узел `toc` хранит только число уровней. Строки рисует вид узла по заголовкам документа; номера страниц
// даёт вёрстка (main.ts передаёт функцию через setPageResolver). В docx оглавление пишется полем TOC
// с готовым результатом (строки, точки, номера), поэтому Word показывает его сразу.

import type { Node as PMNode } from 'prosemirror-model';
import { TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { EditorView, NodeView } from 'prosemirror-view';
import { listCaptions } from './captions';
import { schema } from './schema';

export interface TocEntry { level: number; text: string; pos: number; page?: number }

/** Заголовки до уровня levels (пустые пропускаются). */
export function tocEntries(doc: PMNode, levels: number, caption: string | null = null): TocEntry[] {
  // Список иллюстраций: строки — подписи с этим именем («Рисунок 2 – Схема»), ссылка на абзац подписи.
  if (caption) return listCaptions(doc).filter((c) => c.name === caption).map((c) => ({ level: 1, text: `${c.name.replace(/_/g, ' ')} ${c.n}${c.text ? ' – ' + c.text : ''}`, pos: c.block }));
  const out: TocEntry[] = [];
  doc.descendants((node, pos) => {
    if (node.type === schema.nodes.heading) {
      const text = node.textContent.replace(/\s+/g, ' ').trim();
      if (text && (node.attrs.level as number) <= levels) out.push({ level: node.attrs.level as number, text, pos });
      return false;
    }
    // Заголовки внутри таблиц и рамок в оглавление не идут (как в Word по умолчанию — только основной текст).
    return node.type === schema.nodes.doc;
  });
  return out;
}

type PageResolver = (pos: number) => number | undefined;
let resolvePage: PageResolver = () => undefined;
let onResize: () => void = () => {};
let shownDoc: () => PMNode | null = () => null;
/** Вёрстка: номер страницы (с 1) для позиции в показанном документе и что делать, если оглавление изменило высоту. */
export function setPageResolver(f: PageResolver, resized: () => void, doc: () => PMNode): void { resolvePage = f; onResize = resized; shownDoc = doc; }
/** Номер страницы для записи в файл: известен, только если документ — тот, что на экране. */
export const tocPage = (doc: PMNode, pos: number): number | undefined => (doc === shownDoc() ? resolvePage(pos) : undefined);

const live = new Set<TocView>();

class TocView implements NodeView {
  dom: HTMLElement;
  private sig = '';
  private rows = -1;
  constructor(private node: PMNode, private view: EditorView, private getPos: () => number | undefined) {
    this.dom = document.createElement('div');
    this.dom.className = 'toc';
    this.dom.contentEditable = 'false';
    this.dom.dataset.levels = String(node.attrs.levels);
    this.dom.title = 'Оглавление обновляется само. Щелчок по строке — переход к заголовку.';
    this.dom.addEventListener('mousedown', (e) => {
      const row = (e.target as HTMLElement).closest<HTMLElement>('.toc-row');
      if (!row || e.button !== 0) return;
      e.preventDefault();
      const pos = Number(row.dataset.pos);
      const doc = this.view.state.doc;
      if (pos >= 0 && pos < doc.content.size) {
        this.view.dispatch(this.view.state.tr.setSelection(TextSelection.create(doc, Math.min(pos + 1, doc.content.size))).scrollIntoView());
        this.view.focus();
      }
    });
    live.add(this);
    this.render();
  }
  render(): void {
    const entries = tocEntries(this.view.state.doc, this.node.attrs.levels as number, this.node.attrs.caption as string | null).map((e) => ({ ...e, page: resolvePage(e.pos) }));
    const sig = JSON.stringify(entries.map((e) => [e.level, e.text, e.page, e.pos]));
    if (sig === this.sig) return;
    this.sig = sig;
    this.dom.replaceChildren(...(entries.length ? entries.map((e) => {
      const row = document.createElement('div');
      row.className = 'toc-row';
      row.dataset.level = String(e.level);
      row.dataset.pos = String(e.pos);
      const text = document.createElement('span'); text.className = 'toc-text'; text.textContent = e.text;
      const dots = document.createElement('span'); dots.className = 'toc-dots';
      const page = document.createElement('span'); page.className = 'toc-page'; page.textContent = e.page ? String(e.page) : '';
      row.append(text, dots, page);
      return row;
    }) : [Object.assign(document.createElement('div'), { className: 'toc-empty', textContent: this.node.attrs.caption ? `Список пуст: вставьте подписи «${this.node.attrs.caption}» (Ссылки → Название).` : 'Оглавление пусто: оформите названия разделов стилями «Заголовок 1–3» (вкладка «Главная» → «Стили»).' })]));
    // Число строк изменилось — изменилась высота: страницы нужно перестроить.
    if (entries.length !== this.rows) { const first = this.rows < 0; this.rows = entries.length; if (!first) onResize(); }
  }
  update(node: PMNode): boolean {
    if (node.type !== this.node.type) return false;
    if (node.attrs.levels !== this.node.attrs.levels) this.sig = '';
    this.node = node;
    this.dom.dataset.levels = String(node.attrs.levels);
    this.render();
    return true;
  }
  stopEvent(): boolean { return true; }
  ignoreMutation(): boolean { return true; }
  destroy(): void { live.delete(this); }
}

export const tocViews = { toc: (node: PMNode, view: EditorView, getPos: () => number | undefined) => new TocView(node, view, getPos) };

/** Перерисовать все оглавления (после правки и после перестройки страниц). Дёшево: без изменений DOM не трогается. */
export function refreshTocs(): void { for (const v of live) v.render(); }

/** Вставить оглавление: заголовок «СОДЕРЖАНИЕ» (ГОСТ 7.32: по центру, полужирный) и само оглавление. */
export function insertToc(state: EditorState, levels = 3): Transaction {
  const { $from } = state.selection;
  const at = $from.depth ? $from.before(1) : $from.pos;
  const title = schema.nodes.paragraph.create({ align: 'center', after: 12 }, schema.text('СОДЕРЖАНИЕ', [schema.marks.strong.create()]));
  const toc = schema.nodes.toc.create({ levels });
  const tr = state.tr.insert(at, [title, toc]);
  return tr.scrollIntoView();
}

/** Оглавление в документе есть? Тогда кнопка «Оглавление» меняет число уровней, а не вставляет второе. */
export function findToc(doc: PMNode): { pos: number; node: PMNode } | null {
  let found: { pos: number; node: PMNode } | null = null;
  doc.forEach((node, pos) => { if (!found && node.type === schema.nodes.toc && !node.attrs.caption) found = { pos, node }; });
  return found;
}

/** Список иллюстраций с этим именем подписи (если уже есть — второй не вставляем). */
export function findFigureList(doc: PMNode, caption: string): { pos: number; node: PMNode } | null {
  let found: { pos: number; node: PMNode } | null = null;
  doc.forEach((node, pos) => { if (!found && node.type === schema.nodes.toc && node.attrs.caption === caption) found = { pos, node }; });
  return found;
}

/** Вставить список иллюстраций: заголовок «ПЕРЕЧЕНЬ РИСУНКОВ» (или таблиц, формул) и сам список. */
export function insertFigureList(state: EditorState, caption: string): Transaction {
  const { $from } = state.selection;
  const at = $from.depth ? $from.before(1) : $from.pos;
  const word = ({ 'Рисунок': 'РИСУНКОВ', 'Таблица': 'ТАБЛИЦ', 'Формула': 'ФОРМУЛ' } as Record<string, string>)[caption] ?? caption.toUpperCase();
  const title = schema.nodes.paragraph.create({ align: 'center', after: 12 }, schema.text(`ПЕРЕЧЕНЬ ${word}`, [schema.marks.strong.create()]));
  return state.tr.insert(at, [title, schema.nodes.toc.create({ levels: 1, caption })]).scrollIntoView();
}
