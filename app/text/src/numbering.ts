// Многоуровневая нумерация списков: «1.», «1.1.», «1.2.14.Б», «Статья 1.», «а)» и любые свои схемы.
//
// Номера считаются здесь, по структуре документа (вложенные списки), а не браузером и не «память» Word:
// вложенный список всегда начинается заново под каждым пунктом выше, номер не «слетает» и не продолжается
// сам по себе. Схема (шаблон номера на каждом уровне) хранится у внешнего списка и пишется в docx
// как настоящая многоуровневая нумерация (numbering.xml), поэтому Word и Р7 показывают то же самое.

import type { Node as PMNode } from 'prosemirror-model';
import { Plugin } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';

export interface Level {
  /** Формат числа: decimal, decimalZero, lowerLetter, upperLetter, lowerRoman, upperRoman, russianLower, russianUpper. */
  fmt: string;
  /** Шаблон: %1 — номер первого уровня, %2 — второго и т. д. («%1.%2.%3.», «Статья %1.», «%4)»). */
  text: string;
  start: number;
  /** Ширина номера (отступ текста от номера), pt. */
  hang: number;
}

export type Scheme = Level[];

export const FORMATS: { value: string; label: string; sample: string }[] = [
  { value: 'decimal', label: '1, 2, 3', sample: '1' },
  { value: 'decimalZero', label: '01, 02, 03', sample: '01' },
  { value: 'lowerLetter', label: 'a, b, c', sample: 'a' },
  { value: 'upperLetter', label: 'A, B, C', sample: 'A' },
  { value: 'lowerRoman', label: 'i, ii, iii', sample: 'i' },
  { value: 'upperRoman', label: 'I, II, III', sample: 'I' },
  { value: 'russianLower', label: 'а, б, в', sample: 'а' },
  { value: 'russianUpper', label: 'А, Б, В', sample: 'А' },
];

const RU = 'абвгдежзиклмнопрстуфхцчшщэюя';
const LAT = 'abcdefghijklmnopqrstuvwxyz';

function roman(n: number): string {
  const t: [number, string][] = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
  let s = '';
  for (const [v, r] of t) while (n >= v) { s += r; n -= v; }
  return s;
}

/** Номер в заданном формате. Буквы после конца алфавита удваиваются (я, аа, бб…), как в Word. */
export function formatNumber(n: number, fmt: string): string {
  if (n < 1) n = 1;
  const letters = (abc: string) => abc[(n - 1) % abc.length].repeat(Math.ceil(n / abc.length));
  switch (fmt) {
    case 'decimalZero': return n < 10 ? `0${n}` : String(n);
    case 'lowerLetter': return letters(LAT);
    case 'upperLetter': return letters(LAT).toUpperCase();
    case 'lowerRoman': return roman(n);
    case 'upperRoman': return roman(n).toUpperCase();
    case 'russianLower': return letters(RU);
    case 'russianUpper': return letters(RU).toUpperCase();
    default: return String(n);
  }
}

/** Надпись номера: подставляет в шаблон номера уровней (каждый — в формате своего уровня). */
export function labelFor(scheme: Scheme, level: number, counters: number[]): string {
  const lv = scheme[Math.min(level, scheme.length - 1)];
  return lv.text.replace(/%(\d)/g, (_m, d: string) => {
    const i = +d - 1;
    const l = scheme[Math.min(i, scheme.length - 1)];
    return formatNumber(counters[i] ?? l.start, l.fmt);
  });
}

const lvl = (fmt: string, text: string, hang = 36, start = 1): Level => ({ fmt, text, start, hang });

/** Схема «как раньше»: на каждом уровне свой простой номер «1.», формат берётся у вложенного списка. */
export const SIMPLE: Scheme = Array.from({ length: 9 }, (_, i) => lvl(i % 3 === 0 ? 'decimal' : i % 3 === 1 ? 'lowerLetter' : 'lowerRoman', `%${i + 1}.`));

export interface Preset { id: string; name: string; scheme: Scheme | null }

const legal = (dot: boolean): Scheme => Array.from({ length: 9 }, (_, i) => lvl('decimal', `${Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.')}${dot ? '.' : ''}`, 36 + i * 9));

export const PRESETS: Preset[] = [
  { id: 'simple', name: '1. 2. 3. — обычный (без схемы)', scheme: null },
  { id: 'legal', name: '1. → 1.1. → 1.1.1.', scheme: legal(true) },
  { id: 'gost', name: 'Деловой: 1 → 1.1 → 1.1.1 → а) → 1)', scheme: [
    lvl('decimal', '%1', 28), lvl('decimal', '%1.%2', 36), lvl('decimal', '%1.%2.%3', 45), lvl('russianLower', '%4)', 28), lvl('decimal', '%5)', 28),
    lvl('russianLower', '%6)', 28), lvl('decimal', '%7)', 28), lvl('russianLower', '%8)', 28), lvl('decimal', '%9)', 28)] },
  { id: 'letters', name: '1. → 1.1. → 1.1.1. → 1.1.1.А', scheme: [
    lvl('decimal', '%1.', 28), lvl('decimal', '%1.%2.', 36), lvl('decimal', '%1.%2.%3.', 45), lvl('russianUpper', '%1.%2.%3.%4', 54), lvl('decimal', '%5)', 28),
    lvl('russianLower', '%6)', 28), lvl('decimal', '%7)', 28), lvl('russianLower', '%8)', 28), lvl('decimal', '%9)', 28)] },
  { id: 'items', name: '1. → а) → 1)', scheme: [lvl('decimal', '%1.', 28), lvl('russianLower', '%2)', 28), lvl('decimal', '%3)', 28), lvl('russianLower', '%4)', 28), lvl('decimal', '%5)', 28), lvl('russianLower', '%6)', 28), lvl('decimal', '%7)', 28), lvl('russianLower', '%8)', 28), lvl('decimal', '%9)', 28)] },
  { id: 'article', name: 'Статья 1. → 1.1.', scheme: [lvl('decimal', 'Статья %1.', 64), lvl('decimal', '%1.%2.', 36), lvl('decimal', '%1.%2.%3.', 45), lvl('russianLower', '%4)', 28), lvl('decimal', '%5)', 28), lvl('russianLower', '%6)', 28), lvl('decimal', '%7)', 28), lvl('russianLower', '%8)', 28), lvl('decimal', '%9)', 28)] },
  { id: 'roman', name: 'I. → A. → 1. → a.', scheme: [lvl('upperRoman', '%1.', 36), lvl('upperLetter', '%2.', 28), lvl('decimal', '%3.', 28), lvl('lowerLetter', '%4.', 28), lvl('lowerRoman', '%5.', 28), lvl('decimal', '%6.', 28), lvl('lowerLetter', '%7.', 28), lvl('lowerRoman', '%8.', 28), lvl('decimal', '%9.', 28)] },
];

export const schemeEquals = (a: Scheme | null, b: Scheme | null) => JSON.stringify(a) === JSON.stringify(b);

/** Распознанная схема: если совпадает с готовой — её название. */
export function presetOf(s: Scheme | null): string {
  return PRESETS.find((p) => schemeEquals(p.scheme, s))?.id ?? (s ? 'custom' : 'simple');
}

// --- Подсчёт номеров по документу

export interface Marker { pos: number; label: string; hang: number; level: number }

/** Номера пунктов всех нумерованных списков документа. pos — позиция узла list_item. */
export function computeMarkers(doc: PMNode): Marker[] {
  const out: Marker[] = [];
  const walkList = (list: PMNode, listPos: number, level: number, counters: number[], inherited: Scheme | null): void => {
    const scheme = (list.attrs.scheme as Scheme | null) ?? inherited;
    const ordered = list.type.name === 'ordered_list';
    // order — номер, с которого начат этот список; 1 означает «как задано схемой» (обычно тоже 1).
    let n = ordered ? ((list.attrs.order as number) !== 1 ? (list.attrs.order as number) : (scheme?.[Math.min(level, scheme.length - 1)]?.start ?? 1)) : 0;
    list.forEach((item, off) => {
      const itemPos = listPos + 1 + off;
      if (ordered) {
        const cs = [...counters];
        cs[level] = n;
        const sch = scheme ?? [{ fmt: list.attrs.fmt as string, text: '%1.', start: 1, hang: 36 }];
        // Без схемы номер уровня — один («1.») в формате самого списка.
        const label = scheme ? labelFor(scheme, level, cs) : `${formatNumber(n, sch[0].fmt)}.`;
        out.push({ pos: itemPos, label, hang: scheme ? scheme[Math.min(level, scheme.length - 1)].hang : 36, level });
        n++;
        item.forEach((child, coff) => {
          if (child.type.name === 'ordered_list' || child.type.name === 'bullet_list') walkList(child, itemPos + 1 + coff, level + 1, cs, scheme);
        });
      } else {
        // Маркированный список: своих номеров нет, но под ним могут быть нумерованные — считаем с нуля.
        item.forEach((child, coff) => {
          if (child.type.name === 'ordered_list' || child.type.name === 'bullet_list') walkList(child, itemPos + 1 + coff, level + 1, counters, scheme);
        });
      }
    });
  };
  const top = (node: PMNode, pos: number): void => {
    node.forEach((child, off) => {
      const p = pos + 1 + off;
      if (child.type.name === 'ordered_list' || child.type.name === 'bullet_list') walkList(child, p, 0, [], null);
      else if (child.type.name === 'table') child.forEach((row, ro) => row.forEach((cell, co) => top(cell, p + 1 + ro + 1 + co)));
    });
  };
  top(doc, -1);
  return out;
}

function markerDom(label: string): HTMLElement {
  const s = document.createElement('span');
  s.className = 'num';
  s.contentEditable = 'false';
  s.textContent = label;
  return s;
}

/** Показ номеров: подпись перед пунктом (виджет) и ширина номера пункта (стиль li). */
export const numberingPlugin: Plugin<DecorationSet> = new Plugin<DecorationSet>({
  state: {
    init: (_c, state) => build(state.doc),
    apply: (tr, prev, _o, state) => (tr.docChanged ? build(state.doc) : prev),
  },
  props: { decorations(state): DecorationSet | undefined { return numberingPlugin.getState(state); } },
});

function build(doc: PMNode): DecorationSet {
  const decos: Decoration[] = [];
  for (const m of computeMarkers(doc)) {
    decos.push(Decoration.widget(m.pos + 1, () => markerDom(m.label), { side: -1, key: `n${m.label}`, ignoreSelection: true }));
    decos.push(Decoration.node(m.pos, m.pos + doc.nodeAt(m.pos)!.nodeSize, { style: `--hang:${m.hang}pt`, class: 'numbered' }));
  }
  return DecorationSet.create(doc, decos);
}

/** Пример надписей для окна настройки: «1.», «1.2.», «1.2.14.Б» — как будет выглядеть. */
export function sampleLabels(scheme: Scheme, levels = 5): string[] {
  const sample = [1, 2, 14, 2, 3, 1, 1, 1, 1];
  return Array.from({ length: Math.min(levels, scheme.length) }, (_, i) => labelFor(scheme, i, sample.slice(0, i + 1)));
}
