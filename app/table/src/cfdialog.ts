// Условное форматирование, как в Excel: меню-галерея на «Главной», окно
// «Новое правило» и «Управление правилами». Правила хранит движок (форк
// IronCalc): он же их вычисляет, сохраняет в xlsx и отменяет по Ctrl+Z.

import { columnName, columnNumber, type CfRule, type CfView, type Dxf, type Model } from './engine';

export type Rect = { r1: number; c1: number; r2: number; c2: number };

export interface CfHost {
  model: () => Model;
  sheet: () => number;
  sheetName: () => string;
  /** Выделение; целые строки и столбцы — только до конца данных. */
  selection: () => Rect;
  render: () => void;
  toast: (text: string) => void;
  focus: () => void;
}

let host: CfHost;

export function initConditional(h: CfHost): void {
  host = h;
}

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
import { addRangePicker } from './rangepick';

const overlay = () => document.getElementById('recovery')!;

// --- Диапазоны: «B2:B20 D2:D20» ---

export function rectRef(r: Rect): string {
  const a = `${columnName(r.c1)}${r.r1}`;
  return r.r1 === r.r2 && r.c1 === r.c2 ? a : `${a}:${columnName(r.c2)}${r.r2}`;
}

/** Как пишут люди («$B$2:$B$20; D2») → как хранит Excel («B2:B20 D2»). */
export function normalizeRange(text: string): string {
  return text.toUpperCase().replace(/\$/g, '').replace(/[;,]/g, ' ').trim().replace(/\s+/g, ' ');
}

export function parseRange(text: string): Rect[] | null {
  const out: Rect[] = [];
  for (const part of normalizeRange(text).split(' ')) {
    const m = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(part);
    if (!m) return null;
    const c1 = columnNumber(m[1]), r1 = Number(m[2]);
    const c2 = m[3] ? columnNumber(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
    if (!r1 || !r2) return null;
    out.push({ r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) });
  }
  return out.length ? out : null;
}

const intersects = (a: Rect, b: Rect) => !(b.r2 < a.r1 || b.r1 > a.r2 || b.c2 < a.c1 || b.c1 > a.c2);

/** Прямоугольник без выделенной части: до четырёх кусков (сверху, снизу, слева, справа). */
function subtract(p: Rect, s: Rect): Rect[] {
  if (!intersects(p, s)) return [p];
  const out: Rect[] = [];
  if (p.r1 < s.r1) out.push({ r1: p.r1, c1: p.c1, r2: s.r1 - 1, c2: p.c2 });
  if (p.r2 > s.r2) out.push({ r1: s.r2 + 1, c1: p.c1, r2: p.r2, c2: p.c2 });
  const r1 = Math.max(p.r1, s.r1), r2 = Math.min(p.r2, s.r2);
  if (p.c1 < s.c1) out.push({ r1, c1: p.c1, r2, c2: s.c1 - 1 });
  if (p.c2 > s.c2) out.push({ r1, c1: s.c2 + 1, r2, c2: p.c2 });
  return out;
}

const anchorOf = (rects: Rect[]) => ({ r: Math.min(...rects.map((x) => x.r1)), c: Math.min(...rects.map((x) => x.c1)) });

// --- Готовые форматы, цвета, значки ---

interface Look { fill: string; font: string; b?: boolean; i?: boolean; strike?: boolean; noFill?: boolean }

const PRESETS: [string, string, Look][] = [
  ['red', 'Светло-красная заливка и тёмно-красный текст', { fill: '#FFC7CE', font: '#9C0006' }],
  ['yellow', 'Жёлтая заливка и тёмно-жёлтый текст', { fill: '#FFEB9C', font: '#9C5700' }],
  ['green', 'Зелёная заливка и тёмно-зелёный текст', { fill: '#C6EFCE', font: '#006100' }],
  ['jade', 'Нефритовая заливка и тёмно-нефритовый текст', { fill: '#D5EFE8', font: '#1E5C50' }],
  ['redFill', 'Светло-красная заливка', { fill: '#FFC7CE', font: '#000000' }],
  ['redText', 'Красный текст', { fill: '#FFFFFF', font: '#C00000', noFill: true }],
  ['bold', 'Полужирный текст', { fill: '#FFFFFF', font: '#000000', noFill: true, b: true }],
  ['custom', 'Свой формат', { fill: '#FFF2CC', font: '#000000' }],
];

const BAR_COLORS: [string, string][] = [
  ['#638EC6', 'синяя'], ['#63C384', 'зелёная'], ['#FF555A', 'красная'], ['#FFB628', 'оранжевая'], ['#008AEF', 'голубая'], ['#4FA593', 'нефритовая'],
];

/** Цветовые шкалы: [наименьшее, середина (или ''), наибольшее]. */
const SCALES: [string, string, string, string][] = [
  ['Зелёный — жёлтый — красный (больше — зеленее)', '#F8696B', '#FFEB84', '#63BE7B'],
  ['Красный — жёлтый — зелёный (больше — краснее)', '#63BE7B', '#FFEB84', '#F8696B'],
  ['Зелёный — белый — красный', '#F8696B', '#FCFCFF', '#63BE7B'],
  ['Красный — белый — зелёный', '#63BE7B', '#FCFCFF', '#F8696B'],
  ['Белый — зелёный', '#FCFCFF', '', '#63BE7B'],
  ['Белый — нефритовый', '#FFFFFF', '', '#4FA593'],
];

/** Наборы значков: от наименьших значений к наибольшим (как в движке). */
const ICON_SETS: [string, string, [string, string][]][] = [
  ['3Arrows', '3 стрелки', [['ArrowDown', '#E43400'], ['ArrowRight', '#F2C200'], ['ArrowUp', '#84CB1F']]],
  ['3TrafficLights1', '3 сигнала светофора', [['Circle', '#F8696B'], ['Circle', '#FFD966'], ['Circle', '#63BE7B']]],
  ['3Triangles', '3 треугольника', [['TriangleDown', '#F8696B'], ['FlatRectangle', '#F2C200'], ['TriangleUp', '#63BE7B']]],
  ['3Symbols', '3 символа', [['Cross', '#F8696B'], ['Exclamation', '#F2C200'], ['Check', '#63BE7B']]],
  ['3Flags', '3 флажка', [['Flag', '#F8696B'], ['Flag', '#F2C200'], ['Flag', '#63BE7B']]],
  ['4Arrows', '4 стрелки', [['ArrowDown', '#E43400'], ['ArrowAngleDown', '#F2C200'], ['ArrowAngleUp', '#F2C200'], ['ArrowUp', '#84CB1F']]],
  ['5Arrows', '5 стрелок', [['ArrowDown', '#E43400'], ['ArrowAngleDown', '#F2C200'], ['ArrowRight', '#F2C200'], ['ArrowAngleUp', '#F2C200'], ['ArrowUp', '#84CB1F']]],
  // Рейтинг: в ячейке ряд значков, закрашено тем больше, чем больше значение.
  ['Rating5Star', 'Рейтинг: 5 звёзд', Array.from({ length: 5 }, () => ['Star', '#FFC000'] as [string, string])],
  ['Rating3Star', 'Рейтинг: 3 звезды', Array.from({ length: 3 }, () => ['Star', '#FFC000'] as [string, string])],
  ['Rating5Heart', 'Рейтинг: 5 сердец', Array.from({ length: 5 }, () => ['Heart', '#E8557A'] as [string, string])],
];

const SYMBOL: Record<string, string> = {
  ArrowUp: '↑', ArrowDown: '↓', ArrowRight: '→', ArrowAngleUp: '↗', ArrowAngleDown: '↘', Circle: '●', Flag: '⚑',
  Check: '✔', Cross: '✖', Exclamation: '!', TriangleUp: '▲', TriangleDown: '▼', FlatRectangle: '▬', Star: '★', Heart: '♥',
};

const iconPreview = (icons: [string, string][]) =>
  icons.map(([i, c]) => `<b style="color:${c}">${SYMBOL[i] ?? '●'}</b>`).join('');

const VALUE_OPS: [string, string][] = [
  ['GreaterThan', 'больше'], ['GreaterThanOrEqual', 'больше или равно'], ['LessThan', 'меньше'], ['LessThanOrEqual', 'меньше или равно'],
  ['Equal', 'равно'], ['NotEqual', 'не равно'], ['Between', 'между'], ['NotBetween', 'вне'],
];
const TEXT_OPS: [string, string][] = [
  ['Contains', 'содержит'], ['DoesNotContain', 'не содержит'], ['BeginsWith', 'начинается с'], ['EndsWith', 'заканчивается на'], ['Equals', 'равен'],
];
const PERIODS: [string, string][] = [
  ['Yesterday', 'вчера'], ['Today', 'сегодня'], ['Tomorrow', 'завтра'], ['Last7Days', 'за последние 7 дней'],
  ['LastWeek', 'на прошлой неделе'], ['ThisWeek', 'на этой неделе'], ['NextWeek', 'на следующей неделе'],
  ['LastMonth', 'в прошлом месяце'], ['ThisMonth', 'в этом месяце'], ['NextMonth', 'в следующем месяце'],
  ['LastYear', 'в прошлом году'], ['ThisYear', 'в этом году'], ['NextYear', 'в следующем году'],
  ['Between', 'между датами'], ['NotBetween', 'вне дат'],
];
const ruDate = (iso: unknown) => String(iso ?? '').replace(/^(\d{4})-(\d{2})-(\d{2})$/, '$3.$2.$1');

type Kind = 'cell' | 'text' | 'date' | 'top' | 'avg' | 'dup' | 'blank' | 'formula' | 'bar' | 'scale' | 'icons';

const KINDS: [Kind, string][] = [
  ['cell', 'Значение ячейки'], ['text', 'Текст в ячейке'], ['date', 'Дата'], ['top', 'Первые или последние'],
  ['avg', 'Выше или ниже среднего'], ['dup', 'Повторяющиеся или уникальные'], ['blank', 'Пустые, непустые, ошибки'],
  ['formula', 'Формула (своё условие)'], ['bar', 'Гистограмма'], ['scale', 'Цветовая шкала'], ['icons', 'Набор значков'],
];
const DXF_KINDS = new Set<Kind>(['cell', 'text', 'date', 'top', 'avg', 'dup', 'blank', 'formula']);

const opt = (list: [string, string][], v: string) =>
  list.map(([k, t]) => `<option value="${k}" ${k === v ? 'selected' : ''}>${esc(t)}</option>`).join('');

const label = (list: [string, string][], k: unknown) => list.find((x) => x[0] === k)?.[1] ?? String(k);

/** Цвет из движка (строка или цвет темы) — в «#RRGGBB». */
function hex(color: unknown, fallback: string): string {
  if (color === undefined || color === null) return fallback;
  try {
    const c = host.model().resolveColor(color as string);
    return /^#[0-9a-f]{6}$/i.test(c) ? c.toUpperCase() : fallback;
  } catch {
    return fallback;
  }
}

/** Значение для условия «значение ячейки»: число, ссылка/формула или текст в кавычках. */
function valueFormula(v: string): string {
  const t = v.trim();
  if (t.startsWith('=') || /^-?\d+([.,]\d+)?$/.test(t)) return t;
  return `"${t.replace(/"/g, '""')}"`;
}

function valueText(f: unknown): string {
  const s = String(f ?? '');
  return /^".*"$/.test(s) ? s.slice(1, -1).replace(/""/g, '"') : s;
}

// --- Описание правила для списка ---

export function describe(rule: CfRule): string {
  const r = rule as Record<string, unknown>;
  switch (rule.type) {
    case 'CellIs': {
      const op = r.operator as string;
      if (op === 'Between' || op === 'NotBetween') return `Значение ${op === 'Between' ? 'между' : 'вне'} ${valueText(r.formula)} и ${valueText(r.formula2)}`;
      return `Значение ${label(VALUE_OPS, op)} ${valueText(r.formula)}`;
    }
    case 'Text': return `Текст ${label(TEXT_OPS, r.operator)} «${r.value}»`;
    case 'TimePeriod':
      if (r.time_period === 'Between' || r.time_period === 'NotBetween') {
        return `Дата ${r.time_period === 'Between' ? 'между' : 'вне'} ${ruDate(r.date1)} и ${ruDate(r.date2)}`;
      }
      return `Дата ${label(PERIODS, r.time_period)}`;
    case 'Formula': return `Формула: ${r.formula}`;
    case 'DuplicateValues': return 'Повторяющиеся значения';
    case 'UniqueValues': return 'Уникальные значения';
    case 'Blanks': return 'Пустые ячейки';
    case 'NotBlanks': return 'Непустые ячейки';
    case 'Errors': return 'Ячейки с ошибками';
    case 'NoErrors': return 'Ячейки без ошибок';
    case 'AboveAverage': return 'Выше среднего';
    case 'BelowAverage': return 'Ниже среднего';
    case 'Top10': return `Первые ${r.rank}${r.percent ? ' %' : ''}`;
    case 'Bottom10': return `Последние ${r.rank}${r.percent ? ' %' : ''}`;
    case 'DataBar': return 'Гистограмма';
    case 'ColorScale': return `Цветовая шкала (${(r.thresholds as unknown[]).length} цвета)`;
    case 'IconSet': return 'Набор значков';
    case 'IconRating': return 'Оценка значками';
    default: return rule.type;
  }
}

/** Образец оформления правила для списка. */
function sample(view: CfView): string {
  const r = view.cf_rule as Record<string, unknown>;
  if (view.cf_rule.type === 'DataBar') {
    const c = hex(r.positive_color, '#638EC6');
    return `<span class="cf-sample" style="background:linear-gradient(90deg, ${c} 65%, #fff 65%)"></span>`;
  }
  if (view.cf_rule.type === 'ColorScale') {
    const cols = (r.thresholds as { color: unknown }[]).map((t) => hex(t.color, '#ffffff'));
    return `<span class="cf-sample" style="background:linear-gradient(90deg, ${cols.join(', ')})"></span>`;
  }
  if (view.cf_rule.type === 'IconSet') {
    const icons = (r.thresholds as { icon: string; color: unknown }[]).map((t) => [t.icon, hex(t.color, '#808080')] as [string, string]);
    return `<span class="cf-sample cf-icons">${iconPreview(icons)}</span>`;
  }
  if (view.cf_rule.type === 'IconRating') {
    const sym = SYMBOL[String(r.icon)] ?? '★';
    return `<span class="cf-sample cf-icons"><b style="color:${hex(r.color, '#FFC000')}">${sym.repeat(Math.min(5, (r.thresholds as unknown[]).length || 3))}</b></span>`;
  }
  let dxf: Dxf | null = null;
  try { dxf = host.model().getDxfForConditionalFormatting(host.sheet(), view.index); } catch { dxf = null; }
  return `<span class="cf-sample" style="${dxfCss(dxf)}">АаБб 123</span>`;
}

function dxfCss(dxf: Dxf | null): string {
  if (!dxf) return '';
  const css: string[] = [];
  if (dxf.fill?.color) css.push(`background:${hex(dxf.fill.color, '#ffffff')}`);
  if (dxf.font?.color) css.push(`color:${hex(dxf.font.color, '#000000')}`);
  if (dxf.font?.b) css.push('font-weight:700');
  if (dxf.font?.i) css.push('font-style:italic');
  if (dxf.font?.strike) css.push('text-decoration:line-through');
  return css.join(';');
}

// --- Меню-галерея на ленте ---

let menu: HTMLElement | null = null;

function closeMenu(): void {
  menu?.remove();
  menu = null;
}

document.addEventListener('mousedown', (e) => {
  if (menu && !menu.contains(e.target as Node) && !(e.target as HTMLElement).closest('[data-command="format.conditional"]')) closeMenu();
});
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && menu) { closeMenu(); host.focus(); } }, true);

export function conditionalMenu(x: number, y: number): void {
  if (menu) { closeMenu(); return; }
  const item = (act: string, text: string) => `<button data-act="${act}">${esc(text)}</button>`;
  menu = document.createElement('div');
  menu.className = 'cf-menu';
  menu.innerHTML = `
    <div class="cf-sec">Правила выделения ячеек</div>
    <div class="cf-cols">
      ${item('q:GreaterThan', 'Больше…')}${item('q:LessThan', 'Меньше…')}${item('q:Between', 'Между…')}${item('q:Equal', 'Равно…')}
      ${item('q:text', 'Текст содержит…')}${item('q:date', 'Дата…')}${item('q:dup', 'Повторяющиеся значения…')}
    </div>
    <div class="cf-sec">Первые и последние</div>
    <div class="cf-cols">
      ${item('q:top', 'Первые 10…')}${item('q:topPct', 'Первые 10 %…')}${item('q:bottom', 'Последние 10…')}${item('q:bottomPct', 'Последние 10 %…')}
      ${item('q:above', 'Выше среднего…')}${item('q:below', 'Ниже среднего…')}
    </div>
    <div class="cf-sec">Гистограммы</div>
    <div class="cf-swatches">${BAR_COLORS.map(([c, n], i) => `<button data-act="bar:${i}" title="Гистограмма: ${n}" class="cf-sw" style="background:linear-gradient(90deg, ${c} 70%, #fff 70%)"></button>`).join('')}</div>
    <div class="cf-sec">Цветовые шкалы</div>
    <div class="cf-swatches">${SCALES.map(([n, a, b, c], i) => `<button data-act="scale:${i}" title="${esc(n)}" class="cf-sw" style="background:linear-gradient(90deg, ${[a, b, c].filter(Boolean).join(', ')})"></button>`).join('')}</div>
    <div class="cf-sec">Наборы значков</div>
    <div class="cf-swatches">${ICON_SETS.map(([, n, icons], i) => `<button data-act="icons:${i}" title="${esc(n)}" class="cf-sw cf-icons">${iconPreview(icons)}</button>`).join('')}</div>
    <hr class="menu-sep">
    ${item('new', 'Создать правило…')}
    ${item('clearSel', 'Удалить правила из выделенных ячеек')}
    ${item('clearSheet', 'Удалить правила со всего листа')}
    ${item('manage', 'Управление правилами…')}`;
  document.body.append(menu);
  const { width, height } = menu.getBoundingClientRect();
  menu.style.left = `${Math.max(8, Math.min(x, window.innerWidth - width - 8))}px`;
  menu.style.top = `${Math.max(8, Math.min(y, window.innerHeight - height - 8))}px`;
  menu.addEventListener('mousedown', (e) => e.preventDefault());
  menu.addEventListener('click', (e) => {
    const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
    if (!act) return;
    closeMenu();
    void menuAction(act);
  });
}

async function menuAction(act: string): Promise<void> {
  const [kind, arg] = act.split(':');
  const range = rectRef(host.selection());
  if (kind === 'bar') return quickAdd(range, barRule(BAR_COLORS[Number(arg)][0], true, true));
  if (kind === 'scale') { const [, a, b, c] = SCALES[Number(arg)]; return quickAdd(range, scaleRule(a, b, c)); }
  if (kind === 'icons') return quickAdd(range, iconRule(Number(arg), true));
  if (kind === 'new') { await editRule(null, { range }); host.focus(); return; }
  if (kind === 'manage') return manageRules();
  if (kind === 'clearSel') return clearRules(false);
  if (kind === 'clearSheet') return clearRules(true);
  const quick: Record<string, Partial<FormState>> = {
    GreaterThan: { kind: 'cell', op: 'GreaterThan' }, LessThan: { kind: 'cell', op: 'LessThan' },
    Between: { kind: 'cell', op: 'Between' }, Equal: { kind: 'cell', op: 'Equal' },
    text: { kind: 'text', textOp: 'Contains' }, date: { kind: 'date' }, dup: { kind: 'dup', dup: 'DuplicateValues' },
    top: { kind: 'top', top: 'Top10', rank: 10, percent: false }, topPct: { kind: 'top', top: 'Top10', rank: 10, percent: true },
    bottom: { kind: 'top', top: 'Bottom10', rank: 10, percent: false }, bottomPct: { kind: 'top', top: 'Bottom10', rank: 10, percent: true },
    above: { kind: 'avg', avg: 'AboveAverage' }, below: { kind: 'avg', avg: 'BelowAverage' },
  };
  await editRule(null, { range, ...quick[arg] }, true);
  host.focus();
}

function quickAdd(range: string, rule: CfRule): void {
  try {
    host.model().addConditionalFormatting(host.sheet(), range, rule);
  } catch (e) {
    host.toast(`Не удалось добавить правило: ${e}`);
  }
  host.render();
  host.focus();
}

const barRule = (color: string, gradient: boolean, showValue: boolean): CfRule => ({
  type: 'DataBar', min: null, max: null, positive_color: color, negative_color: '#FF0000', is_gradient: gradient, show_value: showValue,
});

const scaleRule = (low: string, mid: string, high: string): CfRule => ({
  type: 'ColorScale',
  thresholds: mid
    ? [{ cfvo: 'Min', color: low }, { cfvo: { Percentile: 50 }, color: mid }, { cfvo: 'Max', color: high }]
    : [{ cfvo: 'Min', color: low }, { cfvo: 'Max', color: high }],
});

function iconRule(set: number, showValue: boolean): CfRule {
  const icons = ICON_SETS[set][2];
  const n = icons.length;
  if (ICON_SETS[set][0].startsWith('Rating')) {
    return {
      type: 'IconRating', icon: icons[0][0], color: icons[0][1], show_value: showValue,
      thresholds: icons.map((_, i) => [{ Percent: Math.round((100 * i) / n) }, true]),
    };
  }
  return {
    type: 'IconSet',
    show_value: showValue,
    thresholds: icons.map(([icon, color], i) => ({ icon, color, cfvo: { Percent: Math.round((100 * i) / n) }, is_strict: true })),
  };
}

// --- Окно правила ---

interface FormState {
  range: string; kind: Kind; stop: boolean;
  op: string; v1: string; v2: string; textOp: string; text: string; period: string; d1: string; d2: string;
  top: string; rank: number; percent: boolean; avg: string; dup: string; blank: string; formula: string;
  barColor: string; gradient: boolean; showValue: boolean;
  scaleLow: string; scaleMid: string; scaleHigh: string; scale3: boolean; iconSet: number;
  preset: string; look: Look;
}

const DEFAULTS: FormState = {
  range: '', kind: 'cell', stop: false, op: 'GreaterThan', v1: '', v2: '', textOp: 'Contains', text: '', period: 'Today', d1: '', d2: '',
  top: 'Top10', rank: 10, percent: false, avg: 'AboveAverage', dup: 'DuplicateValues', blank: 'Blanks', formula: '',
  barColor: '#638EC6', gradient: true, showValue: true, scaleLow: '#F8696B', scaleMid: '#FFEB84', scaleHigh: '#63BE7B', scale3: true, iconSet: 0,
  preset: 'red', look: { ...PRESETS[0][2] },
};

/** Состояние окна из правила, которое уже есть в книге. */
function stateOf(view: CfView): Partial<FormState> | null {
  const r = view.cf_rule as Record<string, unknown>;
  const base: Partial<FormState> = { range: view.range, stop: !!r.stop_if_true };
  let dxf: Dxf | null = null;
  try { dxf = host.model().getDxfForConditionalFormatting(host.sheet(), view.index); } catch { dxf = null; }
  if (dxf) {
    base.preset = 'custom';
    base.look = {
      fill: hex(dxf.fill?.color, '#FFFFFF'), noFill: !dxf.fill?.color, font: hex(dxf.font?.color, '#000000'),
      b: !!dxf.font?.b, i: !!dxf.font?.i, strike: !!dxf.font?.strike,
    };
  }
  switch (view.cf_rule.type) {
    case 'CellIs': return { ...base, kind: 'cell', op: r.operator as string, v1: valueText(r.formula), v2: valueText(r.formula2) };
    case 'Text': return { ...base, kind: 'text', textOp: r.operator as string, text: String(r.value ?? '') };
    case 'TimePeriod': return PERIODS.some((p) => p[0] === r.time_period) ? { ...base, kind: 'date', period: r.time_period as string, d1: String(r.date1 ?? ''), d2: String(r.date2 ?? '') } : null;
    case 'Formula': return { ...base, kind: 'formula', formula: String(r.formula ?? '') };
    case 'DuplicateValues': case 'UniqueValues': return { ...base, kind: 'dup', dup: view.cf_rule.type };
    case 'Blanks': case 'NotBlanks': case 'Errors': case 'NoErrors': return { ...base, kind: 'blank', blank: view.cf_rule.type };
    case 'AboveAverage': case 'BelowAverage': return { ...base, kind: 'avg', avg: view.cf_rule.type };
    case 'Top10': case 'Bottom10': return { ...base, kind: 'top', top: view.cf_rule.type, rank: Number(r.rank), percent: !!r.percent };
    case 'DataBar': return { ...base, kind: 'bar', barColor: hex(r.positive_color, '#638EC6'), gradient: !!r.is_gradient, showValue: r.show_value !== false };
    case 'ColorScale': {
      const t = (r.thresholds as { color: unknown }[]).map((x) => hex(x.color, '#FFFFFF'));
      return { ...base, kind: 'scale', scale3: t.length > 2, scaleLow: t[0], scaleMid: t.length > 2 ? t[1] : DEFAULTS.scaleMid, scaleHigh: t[t.length - 1] };
    }
    case 'IconSet': {
      const icons = (r.thresholds as { icon: string }[]).map((x) => x.icon).join();
      const i = ICON_SETS.findIndex((s) => s[2].map((x) => x[0]).join() === icons);
      return { ...base, kind: 'icons', iconSet: Math.max(0, i), showValue: r.show_value !== false };
    }
    case 'IconRating': {
      const n = (r.thresholds as unknown[]).length;
      const i = ICON_SETS.findIndex((s) => s[0].startsWith('Rating') && s[2].length === n && s[2][0][0] === r.icon);
      return { ...base, kind: 'icons', iconSet: i < 0 ? ICON_SETS.findIndex((s) => s[0] === 'Rating5Star') : i, showValue: r.show_value !== false };
    }
    default: return null;
  }
}

function dxfOf(look: Look): Dxf {
  return {
    fill: look.noFill ? null : { color: look.fill },
    font: { color: look.font, b: look.b || undefined, i: look.i || undefined, strike: look.strike || undefined },
  };
}

/** Правило из окна или текст ошибки. */
function build(st: FormState): CfRule | string {
  const format = dxfOf(st.look);
  const stop_if_true = st.stop;
  switch (st.kind) {
    case 'cell': {
      if (!st.v1.trim()) return 'Укажите значение.';
      const two = st.op === 'Between' || st.op === 'NotBetween';
      if (two && !st.v2.trim()) return 'Укажите второе значение.';
      return { type: 'CellIs', operator: st.op, formula: valueFormula(st.v1), formula2: two ? valueFormula(st.v2) : null, format, stop_if_true };
    }
    case 'text':
      if (!st.text) return 'Укажите текст.';
      return { type: 'Text', operator: st.textOp, value: st.text, format, stop_if_true };
    case 'date': {
      const two = st.period === 'Between' || st.period === 'NotBetween';
      if (two && (!st.d1 || !st.d2)) return 'Укажите обе даты.';
      return { type: 'TimePeriod', time_period: st.period, date1: two ? st.d1 : null, date2: two ? st.d2 : null, format, stop_if_true };
    }
    case 'top':
      if (!(st.rank >= 1 && st.rank <= (st.percent ? 100 : 1000))) return st.percent ? 'Процент — от 1 до 100.' : 'Число — от 1 до 1000.';
      return { type: st.top, rank: Math.round(st.rank), percent: st.percent, format, stop_if_true };
    case 'avg': return { type: st.avg, format, stop_if_true };
    case 'dup': return { type: st.dup, format, stop_if_true };
    case 'blank': return { type: st.blank, format, stop_if_true };
    case 'formula': {
      const f = st.formula.trim();
      if (!f) return 'Введите формулу, например =$C2>100.';
      return { type: 'Formula', formula: f.startsWith('=') ? f : `=${f}`, format, stop_if_true };
    }
    case 'bar': return barRule(st.barColor, st.gradient, st.showValue);
    case 'scale': return scaleRule(st.scaleLow, st.scale3 ? st.scaleMid : '', st.scaleHigh);
    case 'icons': return iconRule(st.iconSet, st.showValue);
  }
}

/** Окно «Новое правило» / «Изменить правило». true — правило записано. */
export function editRule(view: CfView | null, init: Partial<FormState>, quick = false): Promise<boolean> {
  const st: FormState = { ...DEFAULTS, ...init, look: { ...(init.look ?? DEFAULTS.look) } };
  if (view) {
    const s = stateOf(view);
    if (!s) { host.toast('Это правило из файла можно только удалить или переместить: изменять его здесь пока нельзя.'); return Promise.resolve(false); }
    Object.assign(st, s);
  }
  const rect = parseRange(st.range)?.[0];
  const anchor = rect ? `${columnName(rect.c1)}${rect.r1}` : 'A1';
  const o = overlay();
  o.innerHTML = `<div class="modal cf-modal">
    <h2>${view ? 'Изменить правило' : 'Новое правило форматирования'}</h2>
    <label class="sort-row">Применяется к <input data-f="range" value="${esc(st.range)}" spellcheck="false"></label>
    <label class="sort-row" ${quick ? 'hidden' : ''}>Вид правила <select data-f="kind">${opt(KINDS, st.kind)}</select></label>
    <div class="cf-params" data-k="cell">
      <select data-f="op">${opt(VALUE_OPS, st.op)}</select>
      <input data-f="v1" value="${esc(st.v1)}" placeholder="число, текст или =ссылка">
      <span data-two>и</span><input data-f="v2" data-two value="${esc(st.v2)}" placeholder="второе значение">
    </div>
    <div class="cf-params" data-k="text"><select data-f="textOp">${opt(TEXT_OPS, st.textOp)}</select><input data-f="text" value="${esc(st.text)}" placeholder="текст"></div>
    <div class="cf-params" data-k="date"><span>Дата</span><select data-f="period">${opt(PERIODS, st.period)}</select>
      <input type="date" data-f="d1" data-dates value="${esc(st.d1)}"><span data-dates>и</span><input type="date" data-f="d2" data-dates value="${esc(st.d2)}"></div>
    <div class="cf-params" data-k="top">
      <select data-f="top">${opt([['Top10', 'Первые'], ['Bottom10', 'Последние']], st.top)}</select>
      <input data-f="rank" type="number" min="1" max="1000" value="${st.rank}" class="cf-num">
      <label class="cf-check"><input type="checkbox" data-f="percent" ${st.percent ? 'checked' : ''}> % от всех</label>
    </div>
    <div class="cf-params" data-k="avg"><span>Значения</span><select data-f="avg">${opt([['AboveAverage', 'выше среднего'], ['BelowAverage', 'ниже среднего']], st.avg)}</select></div>
    <div class="cf-params" data-k="dup"><span>Значения</span><select data-f="dup">${opt([['DuplicateValues', 'повторяющиеся'], ['UniqueValues', 'уникальные']], st.dup)}</select></div>
    <div class="cf-params" data-k="blank"><span>Ячейки</span><select data-f="blank">${opt([['Blanks', 'пустые'], ['NotBlanks', 'непустые'], ['Errors', 'с ошибками'], ['NoErrors', 'без ошибок']], st.blank)}</select></div>
    <div class="cf-params cf-col" data-k="formula">
      <input data-f="formula" value="${esc(st.formula)}" placeholder="=$C${rect?.r1 ?? 2}>100" spellcheck="false">
      <small>Условие пишется для левой верхней ячейки диапазона (${anchor}); для остальных ссылки сдвигаются, как при копировании.
      Знак $ закрепляет столбец или строку. Если формула даёт ИСТИНА или не ноль, ячейка выделяется.</small>
    </div>
    <div class="cf-params" data-k="bar">
      <span>Цвет</span><input type="color" data-f="barColor" value="${st.barColor}">
      <label class="cf-check"><input type="checkbox" data-f="gradient" ${st.gradient ? 'checked' : ''}> градиент</label>
    </div>
    <div class="cf-params" data-k="scale">
      <span>Наименьшее</span><input type="color" data-f="scaleLow" value="${st.scaleLow}">
      <label class="cf-check"><input type="checkbox" data-f="scale3" ${st.scale3 ? 'checked' : ''}> середина</label><input type="color" data-f="scaleMid" value="${st.scaleMid}">
      <span>Наибольшее</span><input type="color" data-f="scaleHigh" value="${st.scaleHigh}">
    </div>
    <div class="cf-params" data-k="icons"><select data-f="iconSet">${ICON_SETS.map(([, n], i) => `<option value="${i}" ${i === st.iconSet ? 'selected' : ''}>${esc(n)}</option>`).join('')}</select><span class="cf-icons cf-icon-preview"></span></div>
    <label class="cf-check cf-show" data-k="bar icons"><input type="checkbox" data-f="showValue" ${st.showValue ? 'checked' : ''}> Показывать значение</label>
    <div class="cf-format">
      <label class="sort-row">Формат <select data-f="preset">${opt(PRESETS.map(([k, n]) => [k, n]), st.preset)}</select></label>
      <div class="cf-look">
        <span>Заливка</span><input type="color" data-f="fill" value="${st.look.fill}">
        <label class="cf-check"><input type="checkbox" data-f="noFill" ${st.look.noFill ? 'checked' : ''}> без заливки</label>
        <span>Текст</span><input type="color" data-f="font" value="${st.look.font}">
        <label class="cf-check"><input type="checkbox" data-f="b" ${st.look.b ? 'checked' : ''}> <b>Ж</b></label>
        <label class="cf-check"><input type="checkbox" data-f="i" ${st.look.i ? 'checked' : ''}> <i>К</i></label>
        <label class="cf-check"><input type="checkbox" data-f="strike" ${st.look.strike ? 'checked' : ''}> <s>З</s></label>
      </div>
      <div class="cf-preview">Образец: АаБбВв 123</div>
      <label class="cf-check"><input type="checkbox" data-f="stop" ${st.stop ? 'checked' : ''}> Остановить, если истина (следующие правила для ячейки не проверяются)</label>
    </div>
    <p class="cf-error" hidden></p>
    <div class="rec-actions"><button data-act="ok" class="primary">Применить</button><button data-act="cancel">Отмена</button></div>
  </div>`;
  o.hidden = false;
  const m = o.querySelector<HTMLElement>('.modal')!;
  const f = <T extends HTMLElement = HTMLInputElement>(name: string) => m.querySelector<T>(`[data-f="${name}"]`)!;
  const val = (name: string) => (f(name) as HTMLInputElement | HTMLSelectElement).value;
  const checked = (name: string) => f(name).checked;
  addRangePicker(f('range'));

  const read = () => {
    st.range = normalizeRange(val('range'));
    st.kind = val('kind') as Kind;
    Object.assign(st, {
      op: val('op'), v1: val('v1'), v2: val('v2'), textOp: val('textOp'), text: val('text'), period: val('period'), d1: val('d1'), d2: val('d2'),
      top: val('top'), rank: Number(val('rank')), percent: checked('percent'), avg: val('avg'), dup: val('dup'), blank: val('blank'),
      formula: val('formula'), barColor: val('barColor').toUpperCase(), gradient: checked('gradient'), showValue: checked('showValue'),
      scaleLow: val('scaleLow').toUpperCase(), scaleMid: val('scaleMid').toUpperCase(), scaleHigh: val('scaleHigh').toUpperCase(),
      scale3: checked('scale3'), iconSet: Number(val('iconSet')), preset: val('preset'), stop: checked('stop'),
    });
    st.look = {
      fill: val('fill').toUpperCase(), noFill: checked('noFill'), font: val('font').toUpperCase(),
      b: checked('b'), i: checked('i'), strike: checked('strike'),
    };
  };
  const sync = () => {
    read();
    m.querySelectorAll<HTMLElement>('[data-k]').forEach((el) => { el.hidden = !el.dataset.k!.split(' ').includes(st.kind); });
    m.querySelector<HTMLElement>('.cf-format')!.hidden = !DXF_KINDS.has(st.kind);
    const two = st.op === 'Between' || st.op === 'NotBetween';
    m.querySelectorAll<HTMLElement>('[data-two]').forEach((el) => { el.hidden = !two; });
    const dates = st.period === 'Between' || st.period === 'NotBetween';
    m.querySelectorAll<HTMLElement>('[data-dates]').forEach((el) => { el.hidden = !dates; });
    f('scaleMid').hidden = !st.scale3;
    f('fill').toggleAttribute('disabled', !!st.look.noFill);
    m.querySelector<HTMLElement>('.cf-icon-preview')!.innerHTML = iconPreview(ICON_SETS[st.iconSet][2]);
    m.querySelector<HTMLElement>('.cf-preview')!.style.cssText = dxfCss(dxfOf(st.look));
  };
  m.addEventListener('change', (e) => {
    const t = e.target as HTMLElement;
    if (t.dataset.f === 'preset') {
      const p = PRESETS.find((x) => x[0] === val('preset'))![2];
      f('fill').value = p.fill; f('font').value = p.font;
      f('noFill').checked = !!p.noFill; f('b').checked = !!p.b; f('i').checked = !!p.i; f('strike').checked = !!p.strike;
    } else if (['fill', 'font', 'noFill', 'b', 'i', 'strike'].includes(t.dataset.f ?? '')) {
      f<HTMLSelectElement>('preset').value = 'custom';
    }
    sync();
  });
  m.addEventListener('input', sync);
  sync();

  return new Promise((resolve) => {
    const close = (ok: boolean) => { o.hidden = true; resolve(ok); };
    const error = (text: string) => { const p = m.querySelector<HTMLElement>('.cf-error')!; p.textContent = text; p.hidden = false; };
    const submit = () => {
      read();
      if (!parseRange(st.range)) { error('Диапазон записан непонятно. Пример: B2:B20 или B2:B20 D2:D20.'); f('range').focus(); return; }
      const rule = build(st);
      if (typeof rule === 'string') { error(rule); return; }
      try {
        if (view) host.model().updateConditionalFormatting(host.sheet(), view.index, st.range, rule);
        else host.model().addConditionalFormatting(host.sheet(), st.range, rule);
      } catch (e) {
        error(`Не получилось: ${e}`);
        return;
      }
      host.render();
      close(true);
    };
    m.addEventListener('click', (e) => {
      const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
      if (act === 'ok') submit();
      else if (act === 'cancel') close(false);
    });
    m.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Escape') close(false);
      else if (e.key === 'Enter' && (e.target as HTMLElement).tagName === 'INPUT') { e.preventDefault(); submit(); }
    });
    // Быстрое правило («Больше…») — сразу в поле значения, как в Excel.
    const params = `.cf-params[data-k="${st.kind}"]`;
    const first = quick
      ? m.querySelector<HTMLElement>(`${params} input:not([type="checkbox"])`) ?? m.querySelector<HTMLElement>(`${params} select`)
      : f<HTMLSelectElement>('kind');
    (first ?? f('range')).focus();
  });
}

// --- Удаление ---

/** Правило в виде для записи: формат вместо номера формата. */
function inputOf(view: CfView): CfRule {
  const r = { ...view.cf_rule } as Record<string, unknown>;
  if ('dxf_id' in r) {
    let dxf: Dxf | null = null;
    try { dxf = host.model().getDxfForConditionalFormatting(host.sheet(), view.index); } catch { dxf = null; }
    delete r.dxf_id;
    r.format = dxf ?? {};
  }
  return r as CfRule;
}

const hasFormula = (rule: CfRule) => rule.type === 'Formula' || (rule.type === 'CellIs' && /[A-Za-z]/.test(String(rule.formula ?? '') + String(rule.formula2 ?? '')));

function clearRules(sheet: boolean): void {
  const m = host.model(), sh = host.sheet();
  const sel = host.selection();
  const list = m.getConditionalFormattingList(sh).sort((a, b) => b.index - a.index);
  let removed = 0, kept = 0;
  for (const v of list) {
    try {
      if (sheet) { m.deleteConditionalFormatting(sh, v.index); removed++; continue; }
      const parts = parseRange(v.range);
      if (!parts || !parts.some((p) => intersects(p, sel))) continue;
      const rest = parts.flatMap((p) => subtract(p, sel));
      if (!rest.length) { m.deleteConditionalFormatting(sh, v.index); removed++; continue; }
      const a = anchorOf(parts), b = anchorOf(rest);
      // У правила с формулой ссылки считаются от левой верхней ячейки: если она
      // сдвинется, правило поменяет смысл. Такое правило оставляем как есть.
      if ((a.r !== b.r || a.c !== b.c) && hasFormula(v.cf_rule)) { kept++; continue; }
      m.updateConditionalFormatting(sh, v.index, rest.map(rectRef).join(' '), inputOf(v));
      removed++;
    } catch (e) {
      host.toast(`Не удалось удалить правило: ${e}`);
    }
  }
  host.render();
  host.focus();
  if (kept) host.toast(`Правила с формулой (${kept}) оставлены: их можно удалить в «Управлении правилами».`);
  else if (!removed) host.toast(sheet ? 'На листе нет правил условного форматирования.' : 'В выделенных ячейках нет правил условного форматирования.');
}

// --- Управление правилами ---

export async function manageRules(selectIndex = -1): Promise<void> {
  const m = host.model(), sh = host.sheet();
  const list = m.getConditionalFormattingList(sh);
  const sel = host.selection();
  let current = list.findIndex((v) => v.index === selectIndex);
  if (current < 0) current = list.findIndex((v) => parseRange(v.range)?.some((p) => intersects(p, sel)));
  if (current < 0 && list.length) current = 0;
  const o = overlay();
  o.innerHTML = `<div class="modal cf-modal cf-manager">
    <h2>Управление правилами</h2>
    <p>Лист «${esc(host.sheetName())}». Правила выше в списке главнее: если для ячейки подходят несколько, их форматы
       складываются, а при совпадении берётся верхний.</p>
    <div class="cf-table" role="listbox">
      ${list.length ? list.map((v, i) => `<div class="cf-row" role="option" data-i="${i}">
        <span class="cf-desc">${esc(describe(v.cf_rule))}</span>${sample(v)}<span class="cf-range">${esc(v.range)}</span></div>`).join('')
        : '<div class="cf-empty">На этом листе правил пока нет.</div>'}
    </div>
    <div class="rec-actions cf-actions">
      <button data-act="new">Создать правило…</button>
      <button data-act="edit">Изменить…</button>
      <button data-act="del" class="danger">Удалить</button>
      <button data-act="up" title="Выше (главнее)">▲</button>
      <button data-act="down" title="Ниже">▼</button>
      <span class="cf-spacer"></span>
      <button data-act="close" class="primary">Закрыть</button>
    </div>
  </div>`;
  o.hidden = false;
  const md = o.querySelector<HTMLElement>('.modal')!;
  const mark = () => md.querySelectorAll<HTMLElement>('.cf-row').forEach((r) => r.classList.toggle('active', Number(r.dataset.i) === current));
  mark();
  const reopen = (index: number) => { void manageRules(index); };
  md.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') { o.hidden = true; host.focus(); }
    else if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      current = Math.max(0, Math.min(list.length - 1, current + (e.key === 'ArrowDown' ? 1 : -1)));
      mark();
    }
  });
  md.addEventListener('dblclick', (e) => {
    if ((e.target as HTMLElement).closest('.cf-row')) md.querySelector<HTMLButtonElement>('[data-act="edit"]')!.click();
  });
  md.addEventListener('click', async (e) => {
    const t = e.target as HTMLElement;
    const row = t.closest<HTMLElement>('.cf-row');
    if (row) { current = Number(row.dataset.i); mark(); return; }
    const act = t.closest<HTMLElement>('[data-act]')?.dataset.act;
    if (!act) return;
    const v = list[current];
    try {
      if (act === 'close') { o.hidden = true; host.focus(); return; }
      if (act === 'new') { await editRule(null, { range: rectRef(sel) }); reopen(-1); return; }
      if (!v) return;
      if (act === 'edit') { await editRule(v, {}); reopen(v.index); return; }
      if (act === 'del') {
        m.deleteConditionalFormatting(sh, v.index);
        host.render();
        // после удаления номера правил за ним сдвигаются — выделим соседнее
        const next = list[current + 1] ?? list[current - 1];
        reopen(next ? (next.index > v.index ? next.index - 1 : next.index) : -1);
        return;
      }
      if (act === 'up') m.raiseConditionalFormattingPriority(sh, v.index);
      if (act === 'down') m.lowerConditionalFormattingPriority(sh, v.index);
      host.render();
      reopen(v.index);
    } catch (err) {
      host.toast(String(err));
    }
  });
  md.querySelector<HTMLButtonElement>('[data-act="close"]')!.focus();
}
