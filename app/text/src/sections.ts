// Разделы документа: у каждого свои размер и ориентация страницы, поля. Первый раздел описан в DocMeta.page,
// каждый следующий — в узле section_break перед ним.

import type { Node as PMNode } from 'prosemirror-model';
import type { PageSetup } from './docx';

export interface Section {
  index: number;
  /** Позиция первого блока раздела (для первого — 0). */
  from: number;
  /** Позиция узла-разрыва перед разделом (у первого нет). */
  breakPos?: number;
  page: PageSetup;
}

export function sectionsOf(doc: PMNode, base: PageSetup): Section[] {
  const out: Section[] = [{ index: 0, from: 0, page: base }];
  let cur = base;
  doc.forEach((node, off) => {
    if (node.type.name !== 'section_break') return;
    cur = (node.attrs.page as PageSetup | null) ?? cur;
    out.push({ index: out.length, from: off + node.nodeSize, breakPos: off, page: cur });
  });
  return out;
}

/** Раздел, в котором стоит позиция (разрыв относится к предыдущему разделу). */
export function sectionAt(secs: Section[], pos: number): Section {
  let r = secs[0];
  for (const s of secs) if (s.from <= pos) r = s;
  return r;
}

/** Строка-подпись для сравнения наборов разделов (меняется, когда меняется число или параметры). */
export const sectionsSig = (secs: Section[]): string => secs.map((s) => `${s.page.w}x${s.page.h}:${s.page.top},${s.page.bottom},${s.page.left},${s.page.right}`).join('|');
