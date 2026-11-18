// SPDX-License-Identifier: AGPL-3.0-or-later
// Вставка → Символ: наборы знаков, поиск по названию по-русски («рубль», «стрелка вправо», «градус»), недавние.
// Вставка → Формула (Alt+=, как в Word): запись в стиле LaTeX с живым просмотром и кнопками частых шаблонов.

import { Plugin, TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { parseTex, toMathML } from './math';
import { schema } from './schema';
import { dialog, el } from './ui';

const SETS: [string, string][] = [
  ['Частые', '« » „ “ — – № § © ® ™ ° ± × ÷ ≈ ≠ ≤ ≥ ∞ ½ ¼ ¾ ₽ € $ … • → ← ↑ ↓ ✓ ✗'],
  ['Математика', '± × ÷ − ≠ ≈ ≤ ≥ ∞ √ ∛ ∑ ∏ ∫ ∂ ∆ ∇ ∈ ∉ ⊂ ⊃ ∪ ∩ ∧ ∨ ¬ ∀ ∃ ∅ ‰ ° ′ ″ ² ³ ¼ ½ ¾ ⅓ ⅔'],
  ['Греческие', 'α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ τ υ φ χ ψ ω Α Β Γ Δ Ε Ζ Η Θ Λ Ξ Π Σ Φ Ψ Ω'],
  ['Валюты', '₽ $ € £ ¥ ¢ ₴ ₸ ₹ ₿ ₩ ₺ ₪ ¤'],
  ['Стрелки', '← ↑ → ↓ ↔ ↕ ↖ ↗ ↘ ↙ ⇐ ⇑ ⇒ ⇓ ⇔ ↩ ↪ ↺ ↻ ➔ ▲ ▼ ◄ ►'],
  ['Пунктуация', '« » „ “ ” ‘ ’ ‹ › — – ‐ … · • ‣ § ¶ † ‡ № ¿ ¡ ‽'],
  ['Знаки', '© ® ™ ℃ ℉ ☎ ✉ ✓ ✔ ✗ ✘ ★ ☆ ♥ ♦ ♠ ♣ ☀ ☁ ☂ ☺ ■ □ ● ○ ◆ ◇'],
  ['Индексы', '⁰ ¹ ² ³ ⁴ ⁵ ⁶ ⁷ ⁸ ⁹ ⁺ ⁻ ⁼ ⁽ ⁾ ₀ ₁ ₂ ₃ ₄ ₅ ₆ ₇ ₈ ₉'],
  ['Римские', 'Ⅰ Ⅱ Ⅲ Ⅳ Ⅴ Ⅵ Ⅶ Ⅷ Ⅸ Ⅹ Ⅺ Ⅻ Ⅼ Ⅽ Ⅾ Ⅿ'],
];

/** Названия по-русски для поиска (через запятую — синонимы). */
const NAMES: Record<string, string> = {
  '«': 'кавычка ёлочка открывающая', '»': 'кавычка ёлочка закрывающая', '„': 'кавычка лапка нижняя', '“': 'кавычка лапка верхняя', '—': 'тире длинное', '–': 'тире короткое среднее',
  '№': 'номер', '§': 'параграф раздел', '©': 'копирайт авторское право', '®': 'зарегистрированный знак', '™': 'товарный знак', '°': 'градус',
  '±': 'плюс минус', '×': 'умножение крест', '÷': 'деление', '≈': 'примерно приблизительно', '≠': 'не равно', '≤': 'меньше или равно', '≥': 'больше или равно', '∞': 'бесконечность',
  '½': 'половина дробь', '¼': 'четверть дробь', '¾': 'три четверти дробь', '₽': 'рубль', '€': 'евро', '$': 'доллар', '£': 'фунт', '¥': 'иена юань', '₴': 'гривна', '₸': 'тенге',
  '…': 'многоточие', '•': 'точка маркер', '·': 'точка средняя', '→': 'стрелка вправо', '←': 'стрелка влево', '↑': 'стрелка вверх', '↓': 'стрелка вниз', '↔': 'стрелка в обе стороны',
  '⇒': 'следует двойная стрелка', '⇔': 'равносильно', '✓': 'галочка отметка', '✔': 'галочка жирная', '✗': 'крестик', '✘': 'крестик жирный', '√': 'корень квадратный', '∑': 'сумма',
  '∫': 'интеграл', '∂': 'частная производная', '∆': 'дельта приращение', '∈': 'принадлежит', '∅': 'пустое множество', '‰': 'промилле', '′': 'штрих минута', '″': 'два штриха секунда',
  '²': 'квадрат степень два', '³': 'куб степень три', 'α': 'альфа', 'β': 'бета', 'γ': 'гамма', 'δ': 'дельта', 'ε': 'эпсилон', 'λ': 'лямбда', 'μ': 'мю микро', 'π': 'пи', 'σ': 'сигма',
  'φ': 'фи', 'ω': 'омега', 'Ω': 'омега ом', 'Σ': 'сигма сумма', 'Δ': 'дельта', '℃': 'градус цельсия', '☎': 'телефон', '✉': 'конверт письмо', '★': 'звезда', '☆': 'звезда пустая',
  '♥': 'сердце', '■': 'квадрат чёрный', '□': 'квадрат пустой', '●': 'круг чёрный', '○': 'круг пустой', '¶': 'абзац знак', '†': 'крест', 'Ⅰ': 'римская один', 'Ⅴ': 'римская пять', 'Ⅹ': 'римская десять',
};

const RECENT_KEY = 'ens-text.symbols.recent';
const loadRecent = (): string[] => { try { return JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]'); } catch { return []; } };
const saveRecent = (r: string[]) => { try { localStorage.setItem(RECENT_KEY, JSON.stringify(r.slice(0, 24))); } catch { /* без памяти браузера */ } };

export function symbolsDialog(view: EditorView): void {
  dialog((box, close) => {
    box.classList.add('symbols-dialog');
    let set = 0;
    const search = el('input', { type: 'search', placeholder: 'Поиск по названию: рубль, градус, стрелка…', ariaLabel: 'Поиск символа' });
    const tabs = el('div', { class: 'sym-tabs' });
    const grid = el('div', { class: 'sym-grid' });
    const recentRow = el('div', { class: 'sym-recent' });
    const info = el('p', { class: 'sym-info' });
    const insert = (ch: string) => {
      view.dispatch(view.state.tr.insertText(ch).scrollIntoView());
      saveRecent([ch, ...loadRecent().filter((x) => x !== ch)]);
      drawRecent();
    };
    const btn = (ch: string) => {
      const b = el('button', { class: 'sym', textContent: ch, title: `${NAMES[ch] ?? ''} U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')}`.trim() });
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', () => insert(ch));
      b.addEventListener('mouseenter', () => { info.textContent = `${ch} — ${NAMES[ch] ?? 'знак'} (U+${ch.codePointAt(0)!.toString(16).toUpperCase().padStart(4, '0')})`; });
      return b;
    };
    const drawRecent = () => {
      const r = loadRecent();
      recentRow.replaceChildren(el('b', { textContent: 'Недавние: ' }), ...(r.length ? r.map(btn) : [el('span', { textContent: 'пока нет' })]));
    };
    const draw = () => {
      const q = search.value.trim().toLowerCase();
      tabs.replaceChildren(...SETS.map(([n], i) => el('button', { class: i === set && !q ? 'on' : '', textContent: n, onclick: () => { set = i; search.value = ''; draw(); } })));
      const all = [...new Set(SETS.flatMap(([, s]) => s.split(' ').filter(Boolean)))];
      const chars = q ? all.filter((c) => c === q || (NAMES[c] ?? '').includes(q)) : [...new Set(SETS[set][1].split(' ').filter(Boolean))];
      grid.replaceChildren(...(chars.length ? chars.map(btn) : [el('p', { textContent: 'Ничего не найдено.' })]));
    };
    search.addEventListener('input', draw);
    box.append(el('h2', { textContent: 'Символ' }), search, tabs, grid, recentRow, info,
      el('p', { class: 'hint', textContent: 'Щелчок вставляет знак у курсора; окно можно не закрывать.' }),
      el('div', { class: 'dlg-actions' }, el('button', { class: 'primary', textContent: 'Закрыть', onclick: () => { close(); view.focus(); } })));
    draw();
    drawRecent();
  }, false);
}

const TEMPLATES: [string, string][] = [
  ['a/b', '\\frac{a}{b}'], ['x²', 'x^2'], ['xᵢ', 'x_i'], ['√x', '\\sqrt{x}'], ['ⁿ√x', '\\sqrt[n]{x}'], ['∑', '\\sum_{i=1}^{n}'], ['∫', '\\int_{a}^{b}'],
  ['lim', '\\lim_{x \\to 0}'], ['( )', '\\left( \\right)'], ['α', '\\alpha'], ['π', '\\pi'], ['≤', '\\le'], ['≠', '\\ne'], ['±', '\\pm'], ['·', '\\cdot'], ['∞', '\\infty'],
];

/** Окно формулы: новая (pos не задан) или правка существующей. */
export function mathDialog(view: EditorView, pos?: number): void {
  const node = pos != null ? view.state.doc.nodeAt(pos) : null;
  dialog((box, close) => {
    box.classList.add('math-dialog');
    const input = el('textarea', { rows: 3, value: (node?.attrs.tex as string) ?? '', placeholder: 'Например: x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}', ariaLabel: 'Запись формулы' });
    const preview = el('div', { class: 'math-preview' });
    const sync = () => { preview.innerHTML = input.value.trim() ? toMathML(parseTex(input.value)) : '<span class="hint">Просмотр формулы</span>'; };
    input.addEventListener('input', sync);
    const tpl = el('div', { class: 'math-templates' }, ...TEMPLATES.map(([label, tex]) => {
      const b = el('button', { textContent: label, title: tex });
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', () => { input.setRangeText(tex + ' ', input.selectionStart, input.selectionEnd, 'end'); input.focus(); sync(); });
      return b;
    }));
    const ok = () => {
      const tex = input.value.trim();
      close();
      if (!tex) { view.focus(); return; }
      const m = schema.nodes.math.create({ tex });
      const tr = pos != null && node ? view.state.tr.replaceWith(pos, pos + 1, m) : view.state.tr.replaceSelectionWith(m);
      const at = pos != null ? pos + 1 : tr.mapping.map(view.state.selection.to);
      view.dispatch(tr.setSelection(TextSelection.create(tr.doc, Math.min(at, tr.doc.content.size))).scrollIntoView());
      view.focus();
    };
    input.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); ok(); } });
    box.append(el('h2', { textContent: node ? 'Изменить формулу' : 'Формула' }), tpl, input, preview,
      el('p', { class: 'hint', textContent: 'Запись как в LaTeX: \\frac{a}{b} — дробь, x^2 — степень, x_1 — индекс, \\sqrt{x} — корень, \\alpha — греческие буквы. Ctrl+Enter — вставить. В Word формула откроется как обычная формула.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: () => { close(); view.focus(); } }), el('button', { class: 'primary', textContent: node ? 'Изменить' : 'Вставить', onclick: ok })));
    sync();
    (input as HTMLTextAreaElement & { dataset: DOMStringMap }).dataset.autofocus = '1';
  }, false);
}

/** Двойной щелчок по формуле открывает её для правки. */
export const mathPlugin = new Plugin({
  props: {
    handleDoubleClickOn(view, _pos, node, nodePos) {
      if (node.type !== schema.nodes.math) return false;
      mathDialog(view, nodePos);
      return true;
    },
  },
});
