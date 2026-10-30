// Колонтитулы: показ на листах, измерение высоты и правка на месте (двойной щелчок по верхнему или нижнему полю).
// Содержимое — документ ProseMirror (JSON в DocMeta.hf); поля «номер страницы» и «число страниц» заменяются числами.

import { DOMSerializer } from 'prosemirror-model';
import { EditorView } from 'prosemirror-view';
import { createState } from './editor';
import type { HeaderFooter, HfDoc, PageSetup } from './docx';
import { PX } from './pagination';
import { schema } from './schema';

const serializer = DOMSerializer.fromSchema(schema);

export type HfKind = 'header' | 'footer';

/** Какое содержимое показывает страница i (с нуля): особое для первой или обычное. */
export function hfFor(hf: HeaderFooter, kind: HfKind, pageIndex: number): HfDoc {
  const first = hf.differentFirst && pageIndex === 0;
  if (kind === 'header') return first ? hf.firstHeader : hf.header;
  return first ? hf.firstFooter : hf.footer;
}

export function setHf(hf: HeaderFooter, kind: HfKind, first: boolean, json: HfDoc): void {
  if (kind === 'header') { if (first) hf.firstHeader = json; else hf.header = json; }
  else if (first) hf.firstFooter = json; else hf.footer = json;
}

/** Колонтитул для показа на листе: поля заменены числами. null — пусто. */
export function renderHf(json: HfDoc, kind: HfKind, page: number, pages: number): HTMLElement | null {
  if (!json) return null;
  const doc = schema.nodeFromJSON(json);
  const box = document.createElement('div');
  box.className = `ProseMirror hf hf-${kind}`;
  box.append(serializer.serializeFragment(doc.content));
  for (const f of box.querySelectorAll<HTMLElement>('.fld')) {
    f.className = 'fld-val';
    f.removeAttribute('title');
    f.textContent = String(f.dataset.kind === 'pages' ? pages : page);
  }
  return box;
}

/** Высота колонтитула в px (самого высокого из вариантов); 0 — колонтитулов нет. */
export function measureHf(hf: HeaderFooter, page: PageSetup, kind: HfKind): number {
  const variants = kind === 'header' ? [hf.header, hf.differentFirst ? hf.firstHeader : null] : [hf.footer, hf.differentFirst ? hf.firstFooter : null];
  let max = 0;
  const probe = document.createElement('div');
  probe.style.cssText = `position:absolute;left:-10000px;top:0;visibility:hidden;width:${page.w - page.left - page.right}pt`;
  document.body.append(probe);
  for (const json of variants) {
    const el = renderHf(json, kind, 99, 99);
    if (!el) continue;
    probe.replaceChildren(el);
    max = Math.max(max, el.offsetHeight);
  }
  probe.remove();
  return max;
}

export interface HfSession {
  kind: HfKind;
  first: boolean;
  pageIndex: number;
  view: EditorView;
  close(): void;
}

/** Правка колонтитула на месте: редактор поверх нужной страницы. */
export function openHfEditor(o: {
  sheet: HTMLElement;
  kind: HfKind;
  first: boolean;
  pageIndex: number;
  /** Где лист начинается по вертикали, px, и смещение листа по горизонтали (если он уже самого широкого). */
  pageTop: number;
  offsetX?: number;
  page: PageSetup;
  hf: HeaderFooter;
  onChange: () => void;
}): HfSession {
  const json = o.kind === 'header' ? (o.first ? o.hf.firstHeader : o.hf.header) : o.first ? o.hf.firstFooter : o.hf.footer;
  const doc = json ? schema.nodeFromJSON(json) : schema.nodes.doc.createAndFill()!;
  const host = document.createElement('div');
  host.className = `hf-edit hf-edit-${o.kind}`;
  const dist = (o.kind === 'header' ? o.hf.headerDist : o.hf.footerDist) * PX;
  host.style.left = `calc(${o.offsetX ?? 0}px + ${o.page.left}pt)`;
  host.style.width = `${o.page.w - o.page.left - o.page.right}pt`;
  const base = o.pageTop;
  if (o.kind === 'header') host.style.top = `${base + dist}px`;
  else { host.style.top = `${base + o.page.h * PX - dist}px`; host.style.transform = 'translateY(-100%)'; }
  const label = document.createElement('div');
  label.className = 'hf-label';
  label.textContent = `${o.kind === 'header' ? 'Верхний' : 'Нижний'} колонтитул${o.hf.differentFirst ? (o.first ? ' первой страницы' : '') : ''}`;
  host.append(label);
  o.sheet.append(host);

  const view: EditorView = new EditorView(host, {
    state: createState(doc),
    attributes: { spellcheck: 'true', lang: 'ru', 'aria-label': 'Колонтитул' },
    dispatchTransaction(tr) {
      view.updateState(view.state.apply(tr));
      if (!tr.docChanged) return;
      const d = view.state.doc;
      const blank = d.childCount === 1 && d.firstChild!.content.size === 0;
      setHf(o.hf, o.kind, o.first, blank ? null : (d.toJSON() as HfDoc));
      o.onChange();
    },
  });
  const session: HfSession = {
    kind: o.kind, first: o.first, pageIndex: o.pageIndex, view,
    close() { view.destroy(); host.remove(); },
  };
  view.focus();
  return session;
}
