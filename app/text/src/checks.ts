// Проверка оформления: замечания по документу (ручные пробелы и отступы, пустые абзацы, номера, набранные вручную,
// «выбивающийся» шрифт) и проверка по профилю оформления (ГОСТ Р 7.0.97-2016, ГОСТ 7.32-2017 или свой): поля, шрифт и кегль,
// интервал, абзацный отступ, выравнивание, номера страниц, типографика. Каждое замечание — со ссылкой на места и кнопкой «Исправить».
//
// Значения профилей — рабочие, собраны из открытых источников; в организации может действовать своя редакция.
// Профиль можно править (поле «Свой»), и сверять его с принятым в организации регламентом.

import type { Node as PMNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import type { DocDefaults, DocMeta, PageSetup } from './docx';
import { parseManual, convertManual } from './listtools';
import { schema } from './schema';
import { typograph } from './typograph';

const MM = 2.83465;

export interface Profile {
  id: string;
  name: string;
  note: string;
  /** Наименьшие поля, мм. */
  margins: { left: number; right: number; top: number; bottom: number };
  /** Допустимые шрифты и размеры (pt): шрифт → [от, до]. */
  fonts: Record<string, [number, number]>;
  /** Допустимые множители межстрочного интервала; предпочтительный — первый. */
  lines: number[];
  /** Абзацный отступ, см (0 — не проверять). */
  indentCm: number;
  justify: boolean;
  pageNumbers: { where: 'header' | 'footer'; align: 'center' | 'right' | 'left'; skipFirst: boolean } | null;
}

export const PROFILES: Profile[] = [
  {
    id: 'gost7097', name: 'ГОСТ Р 7.0.97-2016 (организационно-распорядительные документы)',
    note: 'Поля: левое 20, правое 10, верхнее и нижнее 20 мм; Times New Roman 13–14 пт или Arial 12–13 пт; интервал 1 или 1,5; отступ 1,25 см; номера страниц вверху по центру, со второй страницы.',
    margins: { left: 20, right: 10, top: 20, bottom: 20 },
    fonts: { 'Times New Roman': [13, 14], Arial: [12, 13], Verdana: [12, 13] },
    lines: [1, 1.5], indentCm: 1.25, justify: true, pageNumbers: { where: 'header', align: 'center', skipFirst: true },
  },
  {
    id: 'gost732', name: 'ГОСТ 7.32-2017 (отчёт о научно-исследовательской работе)',
    note: 'Поля: левое 30, правое 15, верхнее и нижнее 20 мм; Times New Roman 12–14 пт; интервал 1,5; отступ 1,25 см; номера страниц внизу по центру, на титульном листе не ставятся.',
    margins: { left: 30, right: 15, top: 20, bottom: 20 },
    fonts: { 'Times New Roman': [12, 14] },
    lines: [1.5], indentCm: 1.25, justify: true, pageNumbers: { where: 'footer', align: 'center', skipFirst: true },
  },
];

export interface Issue {
  id: string;
  ok: boolean;
  title: string;
  detail?: string;
  /** Позиции блоков (для «Показать»). */
  at?: number[];
  fix?: () => void;
}

export interface CheckHost {
  state(): EditorState;
  dispatch(tr: Transaction): void;
  meta(): DocMeta;
  changePage(p: Partial<PageSetup>): void;
  setDefaults(p: Partial<DocDefaults>): void;
  /** Номер страницы в колонтитул по профилю (и «особый первой страницы», если нужно). */
  pageNumbers(where: 'header' | 'footer', align: string, skipFirst: boolean): void;
}

const bodyBlocks = (doc: PMNode): { node: PMNode; pos: number }[] => {
  const out: { node: PMNode; pos: number }[] = [];
  doc.forEach((node, pos) => { if (node.type === schema.nodes.paragraph) out.push({ node, pos }); });
  return out;
};

const isBodyText = (n: PMNode) => n.textContent.trim().length > 60 && n.attrs.align !== 'center' && n.attrs.align !== 'right';

const effFont = (n: PMNode, from: number, to: number, d: DocDefaults): { font: string; size: number }[] => {
  const out: { font: string; size: number }[] = [];
  n.nodesBetween(from, to, (c) => {
    if (!c.isText) return true;
    const ts = c.marks.find((m) => m.type.name === 'textStyle');
    out.push({ font: (ts?.attrs.font as string) || d.font, size: (ts?.attrs.size as number) || d.size });
    return false;
  });
  return out;
};

export function runChecks(host: CheckHost, profile: Profile | null): Issue[] {
  const issues: Issue[] = [];
  const st = host.state();
  const meta = host.meta();
  const doc = st.doc;
  const blocks = bodyBlocks(doc);
  const dispatch = (tr: Transaction) => host.dispatch(tr);

  // --- Замечания по документу (без профиля)
  const lead = blocks.filter(({ node }) => /^[ \t ]{2,}\S/.test(node.textContent) || /^\t/.test(node.textContent));
  issues.push({
    id: 'lead', ok: !lead.length, title: lead.length ? `Отступ абзаца набран пробелами или табуляцией: ${lead.length}` : 'Отступы заданы форматом абзаца, а не пробелами',
    detail: 'Отступ первой строки должен быть свойством абзаца: пробелы и табуляции «плывут» при смене шрифта.', at: lead.map((b) => b.pos),
    fix: () => {
      const tr = host.state().tr;
      for (const { node, pos } of [...bodyBlocks(host.state().doc)].reverse()) {
        const m = /^[ \t ]+(?=\S)/.exec(node.textContent);
        if (m && (m[0].length >= 2 || m[0].includes('\t'))) {
          tr.delete(pos + 1, pos + 1 + m[0].length);
          tr.setNodeMarkup(pos, undefined, { ...node.attrs, first: node.attrs.first ?? 35.4 });
        }
      }
      dispatch(tr);
    },
  });

  const dbl = blocks.filter(({ node }) => /\S {2,}\S/.test(node.textContent));
  issues.push({
    id: 'spaces', ok: !dbl.length, title: dbl.length ? `Несколько пробелов подряд: ${dbl.length} абзацев` : 'Двойных пробелов нет',
    at: dbl.map((b) => b.pos), fix: () => { const r = typograph(host.state(), ['spaces']); host.dispatch(r.tr); },
  });

  // Пустые абзацы подряд (≥2) — «интервал» пустыми строками.
  const empties: number[] = [];
  let run = 0;
  doc.forEach((node, pos) => {
    if (node.type === schema.nodes.paragraph && node.content.size === 0) { run++; if (run >= 2) empties.push(pos); } else run = 0;
  });
  issues.push({
    id: 'empty', ok: !empties.length, title: empties.length ? `Лишние пустые абзацы подряд: ${empties.length}` : 'Пустых абзацев подряд нет',
    detail: 'Расстояние между абзацами задаётся интервалом, а не пустыми строками.', at: empties,
    fix: () => {
      const tr = host.state().tr;
      let r = 0;
      const del: [number, number][] = [];
      host.state().doc.forEach((node, pos) => {
        if (node.type === schema.nodes.paragraph && node.content.size === 0) { r++; if (r >= 2) del.push([pos, pos + node.nodeSize]); } else r = 0;
      });
      for (const [a, b] of del.reverse()) tr.delete(a, b);
      dispatch(tr);
    },
  });

  // Номера, набранные вручную.
  const manual = blocks.filter(({ node }) => parseManual(node.textContent)).length;
  issues.push({
    id: 'manual', ok: manual < 2, title: manual >= 2 ? `Нумерация набрана вручную: ${manual} абзацев` : 'Ручной нумерации нет',
    detail: 'Ручные «1.», «1.1», «а)» не пересчитываются при вставке пункта. Превратите их в настоящий список.',
    fix: () => { convertManual(host.state(), (tr) => host.dispatch(tr)); },
  });

  // «Выбивающийся» шрифт: основной текст другого шрифта или размера, чем у большинства.
  const usage = new Map<string, number>();
  let total = 0;
  for (const { node } of blocks) {
    for (const f of effFont(node, 0, node.content.size, meta.defaults)) { const k = `${f.font}|${f.size}`; usage.set(k, (usage.get(k) ?? 0) + 1); total++; }
  }
  const dominant = [...usage.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
  const odd = dominant ? blocks.filter(({ node }) => {
    const fs = effFont(node, 0, node.content.size, meta.defaults);
    return fs.length > 0 && fs.every((f) => `${f.font}|${f.size}` !== dominant) && isBodyText(node);
  }) : [];
  issues.push({
    id: 'odd', ok: !odd.length, title: odd.length ? `Абзацы со шрифтом, выбивающимся из основного (${dominant?.replace('|', ' ')} пт): ${odd.length}` : 'Шрифт основного текста единый',
    at: odd.map((b) => b.pos),
    fix: () => {
      if (!dominant) return;
      const [font, size] = dominant.split('|');
      const tr = host.state().tr;
      for (const { node, pos } of bodyBlocks(host.state().doc)) {
        const fs = effFont(node, 0, node.content.size, meta.defaults);
        if (!fs.length || fs.some((f) => `${f.font}|${f.size}` === dominant) || !isBodyText(node)) continue;
        const ts = schema.marks.textStyle;
        tr.removeMark(pos + 1, pos + 1 + node.content.size, ts);
        if (font !== meta.defaults.font || +size !== meta.defaults.size) tr.addMark(pos + 1, pos + 1 + node.content.size, ts.create({ font, size: +size, color: null, bg: null }));
      }
      dispatch(tr);
    },
  });

  if (!profile) return issues;

  // --- Проверка по профилю
  const p = meta.page;
  const A4 = Math.abs(p.w - 595.3) < 3 && Math.abs(p.h - 841.9) < 3;
  issues.push({
    id: 'a4', ok: A4, title: A4 ? 'Формат листа A4, книжная ориентация' : `Формат листа ${Math.round(p.w / MM)}×${Math.round(p.h / MM)} мм — нужен A4 (210×297 мм), книжная`,
    fix: () => host.changePage({ w: 595.3, h: 841.9 }),
  });
  const m = profile.margins;
  const bad: string[] = [];
  if (p.left < m.left * MM - 1) bad.push(`левое ${Math.round(p.left / MM)} (нужно ≥ ${m.left})`);
  if (p.right < m.right * MM - 1) bad.push(`правое ${Math.round(p.right / MM)} (нужно ≥ ${m.right})`);
  if (p.top < m.top * MM - 1) bad.push(`верхнее ${Math.round(p.top / MM)} (нужно ≥ ${m.top})`);
  if (p.bottom < m.bottom * MM - 1) bad.push(`нижнее ${Math.round(p.bottom / MM)} (нужно ≥ ${m.bottom})`);
  issues.push({
    id: 'margins', ok: !bad.length, title: bad.length ? `Поля, мм: ${bad.join('; ')}` : `Поля не меньше ${m.left}/${m.right}/${m.top}/${m.bottom} мм`,
    fix: () => host.changePage({ left: Math.max(p.left, m.left * MM), right: Math.max(p.right, m.right * MM), top: Math.max(p.top, m.top * MM), bottom: Math.max(p.bottom, m.bottom * MM) }),
  });

  // Шрифт и размер.
  const allowed = (font: string, size: number) => { const r = profile.fonts[font]; return !!r && size >= r[0] && size <= r[1]; };
  const wrongFont = blocks.filter(({ node }) => effFont(node, 0, node.content.size, meta.defaults).some((f) => !allowed(f.font, f.size)));
  const preferred = Object.entries(profile.fonts)[0];
  issues.push({
    id: 'font', ok: !wrongFont.length,
    title: wrongFont.length ? `Шрифт или размер вне профиля: ${wrongFont.length} абзацев` : 'Шрифт и размер соответствуют профилю',
    detail: `Допустимо: ${Object.entries(profile.fonts).map(([f, [a, b]]) => `${f} ${a}${a === b ? '' : `–${b}`} пт`).join('; ')}.`,
    at: wrongFont.map((b) => b.pos),
    fix: () => {
      const [font, [, size]] = preferred;
      host.setDefaults({ font, size });
      const tr = host.state().tr;
      const ts = schema.marks.textStyle;
      const defs = { ...meta.defaults, font, size };
      for (const { node, pos } of bodyBlocks(host.state().doc)) {
        node.descendants((c, off) => {
          if (!c.isText) return true;
          const m1 = c.marks.find((x) => x.type === ts);
          const f = (m1?.attrs.font as string) || defs.font;
          const s = (m1?.attrs.size as number) || defs.size;
          if (m1 && !allowed(f, s)) {
            const a = pos + 1 + off;
            tr.removeMark(a, a + c.nodeSize, ts);
            if (m1.attrs.color || m1.attrs.bg) tr.addMark(a, a + c.nodeSize, ts.create({ font: null, size: null, color: m1.attrs.color, bg: m1.attrs.bg }));
          }
          return false;
        });
      }
      dispatch(tr);
    },
  });

  // Межстрочный интервал.
  const badLine = blocks.filter(({ node }) => node.textContent.length > 0 && !profile.lines.includes((node.attrs.line as number | null) ?? meta.defaults.line));
  issues.push({
    id: 'line', ok: !badLine.length, title: badLine.length ? `Межстрочный интервал вне профиля: ${badLine.length} абзацев` : `Межстрочный интервал: ${profile.lines.join(' или ')}`,
    at: badLine.map((b) => b.pos),
    fix: () => {
      host.setDefaults({ line: profile.lines[0] });
      const tr = host.state().tr;
      for (const { node, pos } of bodyBlocks(host.state().doc)) if (node.attrs.line && !profile.lines.includes(node.attrs.line as number)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, line: null });
      dispatch(tr);
    },
  });

  // Абзацный отступ и выравнивание основного текста.
  if (profile.indentCm) {
    const want = profile.indentCm * 28.3465;
    const noInd = blocks.filter(({ node }) => isBodyText(node) && Math.abs(((node.attrs.first as number | null) ?? 0) - want) > 2);
    issues.push({
      id: 'indent', ok: !noInd.length, title: noInd.length ? `Абзацный отступ не ${profile.indentCm} см: ${noInd.length} абзацев` : `Абзацный отступ ${profile.indentCm} см`,
      at: noInd.map((b) => b.pos),
      fix: () => {
        const tr = host.state().tr;
        for (const { node, pos } of bodyBlocks(host.state().doc)) if (isBodyText(node)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, first: want, indent: null });
        dispatch(tr);
      },
    });
  }
  if (profile.justify) {
    const noJ = blocks.filter(({ node }) => isBodyText(node) && node.attrs.align !== 'justify');
    issues.push({
      id: 'justify', ok: !noJ.length, title: noJ.length ? `Текст не выровнен по ширине: ${noJ.length} абзацев` : 'Основной текст выровнен по ширине',
      at: noJ.map((b) => b.pos),
      fix: () => {
        const tr = host.state().tr;
        for (const { node, pos } of bodyBlocks(host.state().doc)) if (isBodyText(node)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, align: 'justify' });
        dispatch(tr);
      },
    });
  }

  // Номера страниц.
  if (profile.pageNumbers) {
    const pn = profile.pageNumbers;
    const hasField = (j: Record<string, unknown> | null) => !!j && JSON.stringify(j).includes('"type":"field"');
    const hf = meta.hf;
    const where = pn.where === 'header' ? hf.header : hf.footer;
    const first = pn.where === 'header' ? hf.firstHeader : hf.firstFooter;
    const okNum = hasField(where) && (!pn.skipFirst || (hf.differentFirst && !hasField(first)));
    issues.push({
      id: 'pagenum', ok: okNum,
      title: okNum ? 'Номера страниц стоят по профилю' : `Номера страниц: нужны ${pn.where === 'header' ? 'вверху' : 'внизу'} по центру${pn.skipFirst ? ', на первой странице не ставятся' : ''}`,
      fix: () => host.pageNumbers(pn.where, pn.align, pn.skipFirst),
    });
  }

  // Типографика.
  const typo = typograph(st);
  issues.push({
    id: 'typo', ok: typo.report.total === 0, title: typo.report.total ? `Типографика: ${typo.report.total} замен (кавычки, тире, пробелы)` : 'Типографика в порядке',
    fix: () => { host.dispatch(typograph(host.state()).tr); },
  });
  return issues;
}
