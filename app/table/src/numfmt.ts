// Числовые форматы, как в Excel: кнопки «Разделитель групп разрядов»,
// «Увеличить/уменьшить разрядность» и окно «Формат ячеек» (Ctrl+1).
// Коды хранятся по-английски, как в xlsx («#,##0.00»), а показываются по
// правилам русской локали («1 234,50»).

import type { Model } from './engine';

/** Палитра Excel для [Red], [Color10] и т. п. (номера 1..56). */
export const PALETTE = [
  '', '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#800000', '#008000',
  '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080', '#9999FF', '#993366', '#FFFFCC', '#CCFFFF',
  '#660066', '#FF8080', '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000',
  '#008080', '#0000FF', '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696', '#003366', '#339966',
  '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];

// --- Разбор кода формата ---

/** Куски кода: в кавычках, экранированные, в скобках — не трогаем. */
function pieces(code: string): { text: string; raw: boolean }[] {
  const out: { text: string; raw: boolean }[] = [];
  let i = 0, plain = '';
  const flush = () => { if (plain) { out.push({ text: plain, raw: false }); plain = ''; } };
  while (i < code.length) {
    const c = code[i];
    if (c === '"') {
      const j = code.indexOf('"', i + 1);
      const end = j < 0 ? code.length : j + 1;
      flush(); out.push({ text: code.slice(i, end), raw: true }); i = end;
    } else if (c === '[') {
      const j = code.indexOf(']', i + 1);
      const end = j < 0 ? code.length : j + 1;
      flush(); out.push({ text: code.slice(i, end), raw: true }); i = end;
    } else if (c === '\\' || c === '_' || c === '*') {
      flush(); out.push({ text: code.slice(i, i + 2), raw: true }); i += 2;
    } else {
      plain += c; i++;
    }
  }
  flush();
  return out;
}

/** Разделы кода через «;» (вне кавычек и скобок). */
function sections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  for (const p of pieces(code)) {
    if (p.raw) { cur += p.text; continue; }
    const parts = p.text.split(';');
    cur += parts[0];
    for (const part of parts.slice(1)) { out.push(cur); cur = part; }
  }
  out.push(cur);
  return out;
}

const plainText = (code: string) => pieces(code).filter((p) => !p.raw).map((p) => p.text).join('');

export const isGeneral = (code: string) => !code || code.toLowerCase() === 'general';
export const isText = (code: string) => code.trim() === '@';
/** Дата или время: есть d, m, y, h, s и нет цифровых мест. */
export const isDateTime = (code: string) => {
  const t = plainText(sections(code)[0]).toLowerCase();
  return /[dmyhs]/.test(t) && !/[#0?]/.test(t.replace(/\.0+/g, ''));
};

/** Число в разделе: первое «0/#/?» с запятыми и дробной частью (вне кавычек). */
function mapNumber(section: string, fn: (int: string, frac: string | null) => string): string {
  let done = false;
  return pieces(section).map((p) => {
    if (p.raw || done) return p.text;
    return p.text.replace(/[#0?][#0?,]*(?:\.[#0?]*)?|\.[#0?]+/, (m) => {
      done = true;
      const dot = m.indexOf('.');
      return fn(dot < 0 ? m : m.slice(0, dot), dot < 0 ? null : m.slice(dot + 1));
    });
  }).join('');
}

const mapSections = (code: string, fn: (s: string) => string) => sections(code).map(fn).join(';');

/** Сколько знаков после запятой показано у числа (для «Общего»). */
export function shownDecimals(shown: string): number {
  const m = /[.,](\d+)/.exec(shown.replace(/\s/g, ''));
  return m && !/E/i.test(shown) ? m[1].length : 0;
}

const zeros = (n: number) => '0'.repeat(Math.max(0, n));

/** По умолчанию — разряды и отрицательные красным (по желанию пользователя). */
export const withRed = (base: string) => `${base};[Red]-${base}`;

/** Простой числовой формат из одной секции (без процента) получает красные отрицательные. */
const redIfPlain = (code: string): string => (code.includes(';') || /[%@"\[]/.test(code) ? code : withRed(code));

/** «Увеличить/уменьшить разрядность»: delta = +1 или −1. null — к этому формату неприменимо. */
export function changeDecimals(code: string, delta: number, shown: string): string | null {
  if (isText(code) || isDateTime(code)) return null;
  if (isGeneral(code)) {
    const n = Math.max(0, shownDecimals(shown) + delta);
    return withRed(n ? `#,##0.${zeros(n)}` : '#,##0');
  }
  return redIfPlain(mapSections(code, (s) => mapNumber(s, (int, frac) => {
    const n = Math.max(0, (frac?.length ?? 0) + delta);
    return n ? `${int || '0'}.${(frac ?? '').replace(/[#?]/g, '0').padEnd(n, '0').slice(0, n)}` : (int || '0');
  })));
}

/** Есть ли разделитель групп разрядов («#,##0»). */
export function hasGrouping(code: string): boolean {
  let found = false;
  mapNumber(sections(code)[0], (int) => { found = int.includes(','); return int; });
  return found;
}

/** «Разделитель групп разрядов»: включить или выключить. */
export function toggleGrouping(code: string, shown: string): string | null {
  if (isText(code) || isDateTime(code)) return null;
  if (isGeneral(code)) {
    const n = shownDecimals(shown);
    return withRed(n ? `#,##0.${zeros(n)}` : '#,##0');
  }
  const on = !hasGrouping(code);
  return redIfPlain(mapSections(code, (s) => mapNumber(s, (int, frac) => {
    const digits = int.replace(/,/g, '');
    const next = on ? `#,##${digits.replace(/^[#?]*/, '').padStart(1, '0')}` : digits.replace(/^#+(?=0)/, '') || '0';
    return frac === null ? next : `${next}.${frac}`;
  })));
}

// --- Окно «Формат ячеек» ---

export interface FormatHost {
  model: () => Model;
  sheet: () => number;
  active: () => { row: number; column: number };
  apply: (code: string) => void;
}

type Category = 'general' | 'number' | 'currency' | 'date' | 'time' | 'percent' | 'fraction' | 'scientific' | 'text' | 'custom';

const CATEGORIES: [Category, string][] = [
  ['general', 'Общий'], ['number', 'Числовой'], ['currency', 'Денежный'], ['date', 'Дата'], ['time', 'Время'],
  ['percent', 'Процентный'], ['fraction', 'Дробный'], ['scientific', 'Экспоненциальный'], ['text', 'Текстовый'], ['custom', 'Все форматы'],
];

const CURRENCIES: [string, string][] = [['₽', '₽ (рубль)'], ['$', '$ (доллар)'], ['€', '€ (евро)'], ['¥', '¥ (юань)'], ['', 'нет']];

const NEGATIVES: [string, string][] = [
  ['minus', '−1 234,10'],
  ['red', '1 234,10 красным'],
  ['redMinus', '−1 234,10 красным'],
];

export const DATE_CODES = ['dd.mm.yyyy', 'dd.mm.yy', 'd mmmm yyyy', 'd mmm yyyy', 'dddd, d mmmm yyyy', 'mmmm yyyy', 'mmm yy', 'dd.mm', 'dd.mm.yyyy hh:mm', 'yyyy-mm-dd'];
export const FRACTION_CODES = ['# ?/?', '# ??/??', '# ???/???', '# ?/2', '# ?/4', '# ?/8', '# ??/16', '# ?/10', '# ??/100'];
export const TIME_CODES = ['hh:mm', 'hh:mm:ss', 'h:mm AM/PM', '[h]:mm', '[h]:mm:ss', 'mm:ss', 'dd.mm.yyyy hh:mm:ss'];

const CUSTOM_CODES = [
  'general', '0', '0.00', '#,##0', '#,##0.00', '#,##0;[Red]-#,##0', '#,##0.00;[Red]-#,##0.00', '#,##0.00 "₽"',
  '0%', '0.00%', '0.00E+00', '@', 'dd.mm.yyyy', 'd mmmm yyyy', 'hh:mm', '[h]:mm',
  '#,##0 "шт."', '#,##0.00 "кг"', '0" %"', '+0;-0;0',
];

interface Num { decimals: number; grouping: boolean; negative: string; symbol: string }

function numberCode(n: Num, kind: 'number' | 'currency'): string {
  let base = `${n.grouping || kind === 'currency' ? '#,##0' : '0'}${n.decimals ? '.' + zeros(n.decimals) : ''}`;
  if (kind === 'currency' && n.symbol) base = n.symbol === '$' ? `"$"${base}` : `${base} "${n.symbol}"`;
  if (n.negative === 'red') return `${base};[Red]${base}`;
  if (n.negative === 'redMinus') return `${base};[Red]-${base}`;
  return base;
}

/** Разбор существующего кода в категорию и настройки окна. */
function classify(code: string): { cat: Category; num: Num } {
  const num: Num = { decimals: 2, grouping: true, negative: 'redMinus', symbol: '₽' };
  if (isGeneral(code)) return { cat: 'general', num };
  if (isText(code)) return { cat: 'text', num };
  if (DATE_CODES.includes(code)) return { cat: 'date', num };
  if (TIME_CODES.includes(code)) return { cat: 'time', num };
  if (FRACTION_CODES.includes(code)) return { cat: 'fraction', num };
  num.negative = 'minus'; // у существующего кода — как в нём записано
  const first = sections(code)[0];
  let decimals = 0;
  mapNumber(first, (_i, frac) => { decimals = frac?.length ?? 0; return ''; });
  num.decimals = decimals;
  num.grouping = hasGrouping(code);
  const secs = sections(code);
  if (secs.length > 1 && /\[red\]/i.test(secs[1])) num.negative = /\]-/.test(secs[1]) ? 'redMinus' : 'red';
  const pct = /^0(\.0+)?%$/.test(code);
  if (pct) return { cat: 'percent', num };
  if (/^0(\.0+)?E\+00$/i.test(code)) return { cat: 'scientific', num };
  const sym = /"([₽$€¥])"/.exec(first)?.[1];
  for (const kind of ['number', 'currency'] as const) {
    const n = { ...num, symbol: sym ?? '' };
    if (numberCode(n, kind) === code) return { cat: kind, num: n };
  }
  return { cat: 'custom', num };
}

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
const opt = (list: [string, string][], v: string) => list.map(([k, t]) => `<option value="${esc(k)}" ${k === v ? 'selected' : ''}>${esc(t)}</option>`).join('');

/** Окно «Формат ячеек» → «Число». */
export function formatCellsDialog(host: FormatHost, current: string): Promise<void> {
  const m = host.model(), sheet = host.sheet(), at = host.active();
  const init = classify(current);
  let cat: Category = init.cat;
  const num = init.num;
  let dateCode = DATE_CODES.includes(current) ? current : DATE_CODES[0];
  let fractionCode = FRACTION_CODES.includes(current) ? current : FRACTION_CODES[0];
  let timeCode = TIME_CODES.includes(current) ? current : TIME_CODES[0];
  let custom = isGeneral(current) ? 'general' : current;

  const o = document.getElementById('recovery')!;
  o.innerHTML = `<div class="modal fmt-modal">
    <h2>Формат ячеек</h2>
    <div class="fmt-body">
      <div class="fmt-cats" role="listbox">${CATEGORIES.map(([k, t]) => `<button data-cat="${k}" role="option">${t}</button>`).join('')}</div>
      <div class="fmt-pane">
        <div class="fmt-sample"><span>Образец</span><b></b></div>
        <div class="fmt-opts"></div>
        <p class="fmt-hint"></p>
      </div>
    </div>
    <div class="rec-actions"><button data-act="ok" class="primary">ОК</button><button data-act="cancel">Отмена</button></div>
  </div>`;
  o.hidden = false;
  const md = o.querySelector<HTMLElement>('.modal')!;
  const opts = md.querySelector<HTMLElement>('.fmt-opts')!;
  const hint = md.querySelector<HTMLElement>('.fmt-hint')!;
  const sample = md.querySelector<HTMLElement>('.fmt-sample b')!;

  const code = (): string => {
    switch (cat) {
      case 'general': return 'general';
      case 'number': return numberCode(num, 'number');
      case 'currency': return numberCode(num, 'currency');
      case 'date': return dateCode;
      case 'time': return timeCode;
      case 'fraction': return fractionCode;
      case 'percent': return `0${num.decimals ? '.' + zeros(num.decimals) : ''}%`;
      case 'scientific': return `0${num.decimals ? '.' + zeros(num.decimals) : ''}E+00`;
      case 'text': return '@';
      case 'custom': return custom.trim() || 'general';
    }
  };
  const showSample = () => {
    let text = '';
    try { text = m.formatCellValueAs(sheet, at.row, at.column, code()); } catch { text = '—'; }
    sample.textContent = text || '(пустая ячейка)';
    // Отрицательные «красным» — показать и в образце.
    sample.classList.toggle('fmt-red', /\[red\]/i.test(sections(code())[1] ?? '') && isNegative());
  };
  const isNegative = () => { try { return m.formatCellValueAs(sheet, at.row, at.column, '0').startsWith('-'); } catch { return false; } };
  const decimalsField = `<label class="fmt-row">Число десятичных знаков <input type="number" min="0" max="15" data-f="decimals" value="${num.decimals}"></label>`;
  const listField = (name: string, codes: string[], value: string) => `<div class="fmt-list" data-f="${name}">${codes.map((c) => {
    let shown = c;
    try { shown = m.formatCellValueAs(sheet, at.row, at.column, c) || c; } catch { /* как есть */ }
    if (!/\d/.test(shown) || m.getCellType(sheet, at.row, at.column) !== 1) shown = c;
    return `<button data-code="${esc(c)}" class="${c === value ? 'active' : ''}"><span>${esc(shown)}</span><small>${esc(c)}</small></button>`;
  }).join('')}</div>`;

  const render = () => {
    md.querySelectorAll<HTMLElement>('[data-cat]').forEach((b) => b.classList.toggle('active', b.dataset.cat === cat));
    const hints: Record<Category, string> = {
      general: 'Как ввели: без разделителей и лишних нулей. Длинные числа сокращаются.',
      number: 'Для обычных чисел. Разделитель групп разрядов: 1 234 567.',
      currency: 'Денежные суммы: разделитель разрядов и знак валюты.',
      date: 'Дата хранится числом (дни с 1900 года), формат меняет только вид.',
      time: 'Время — доля суток. [ч]:мм показывает больше 24 часов (табель, длительность).',
      percent: 'Значение умножается на 100 и показывается со знаком %: 0,15 → 15%.',
      fraction: 'Число простой дробью: 1,5 → 1 1/2. ?/? — знаменатель до 9, ??/?? — до 99, ?/8 — восьмые доли.',
      scientific: 'Очень большие и маленькие числа: 1,23E+06 = 1 230 000.',
      text: 'Ячейка — текст, даже если в ней цифры (номера, коды, ИНН). Действует на новый ввод.',
      custom: 'Свой код в записи Excel: 0 — обязательная цифра, # — необязательная, «,» — разряды, «.» — дробная часть, ; — отрицательные, "текст" — надпись.',
    };
    hint.textContent = hints[cat];
    if (cat === 'number') {
      opts.innerHTML = `${decimalsField}
        <label class="fmt-check"><input type="checkbox" data-f="grouping" ${num.grouping ? 'checked' : ''}> Разделитель групп разрядов (1 234)</label>
        <label class="fmt-row">Отрицательные числа <select data-f="negative">${opt(NEGATIVES, num.negative)}</select></label>`;
    } else if (cat === 'currency') {
      opts.innerHTML = `${decimalsField}
        <label class="fmt-row">Обозначение <select data-f="symbol">${opt(CURRENCIES, num.symbol)}</select></label>
        <label class="fmt-row">Отрицательные числа <select data-f="negative">${opt(NEGATIVES, num.negative)}</select></label>`;
    } else if (cat === 'percent' || cat === 'scientific') {
      opts.innerHTML = decimalsField;
    } else if (cat === 'date') {
      opts.innerHTML = listField('date', DATE_CODES, dateCode);
    } else if (cat === 'fraction') {
      opts.innerHTML = listField('fraction', FRACTION_CODES, fractionCode);
    } else if (cat === 'time') {
      opts.innerHTML = listField('time', TIME_CODES, timeCode);
    } else if (cat === 'custom') {
      opts.innerHTML = `<label class="fmt-row">Код <input data-f="custom" value="${esc(custom)}" spellcheck="false"></label>
        ${listField('customList', CUSTOM_CODES.includes(custom) ? CUSTOM_CODES : [custom, ...CUSTOM_CODES], custom)}`;
    } else {
      opts.innerHTML = '';
    }
    showSample();
  };

  const readInputs = () => {
    const d = md.querySelector<HTMLInputElement>('[data-f="decimals"]');
    if (d) num.decimals = Math.max(0, Math.min(15, Math.round(Number(d.value) || 0)));
    const g = md.querySelector<HTMLInputElement>('[data-f="grouping"]');
    if (g) num.grouping = g.checked;
    const n = md.querySelector<HTMLSelectElement>('[data-f="negative"]');
    if (n) num.negative = n.value;
    const s = md.querySelector<HTMLSelectElement>('[data-f="symbol"]');
    if (s) num.symbol = s.value;
    const c = md.querySelector<HTMLInputElement>('[data-f="custom"]');
    if (c) custom = c.value;
    showSample();
  };
  md.addEventListener('input', readInputs);
  md.addEventListener('change', readInputs);

  return new Promise((resolve) => {
    const close = (ok: boolean) => {
      o.hidden = true;
      if (ok) host.apply(code());
      resolve();
    };
    md.addEventListener('click', (e) => {
      const t = e.target as HTMLElement;
      const c = t.closest<HTMLElement>('[data-cat]');
      if (c) { cat = c.dataset.cat as Category; if (cat === 'custom' && custom === 'general' && !isGeneral(current)) custom = current; render(); return; }
      const item = t.closest<HTMLElement>('[data-code]');
      if (item) {
        const list = item.parentElement!.dataset.f;
        const v = item.dataset.code!;
        if (list === 'date') dateCode = v;
        else if (list === 'time') timeCode = v;
        else if (list === 'fraction') fractionCode = v;
        else { custom = v; }
        render();
        return;
      }
      const act = t.closest<HTMLElement>('[data-act]')?.dataset.act;
      if (act === 'ok') close(true);
      else if (act === 'cancel') close(false);
    });
    md.addEventListener('dblclick', (e) => { if ((e.target as HTMLElement).closest('[data-code]')) close(true); });
    md.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') close(false);
      else if (e.key === 'Enter' && (e.target as HTMLElement).tagName !== 'BUTTON') { e.preventDefault(); close(true); }
    });
    render();
    md.querySelector<HTMLButtonElement>(`[data-cat="${cat}"]`)!.focus();
  });
}
