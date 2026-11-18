// SPDX-License-Identifier: AGPL-3.0-or-later
// Вставка → Иллюстрации: фигуры, значки, диаграммы, схемы (вместо SmartArt), видео-ссылка.
//
// Всё рисуется в SVG и вставляется картинкой PNG (вдвое чётче экрана): так документ одинаково выглядит в Word,
// LibreOffice и Р7, а в docx нет ничего, кроме обычного рисунка. Данные диаграммы и текст схемы пишутся в
// замещающий текст рисунка — их видно программам чтения с экрана и в свойствах рисунка в Word.

import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { ICON_SET } from './icon-set';
import { schema } from './schema';
import { tableAt } from './tables-ui';
import { dialog, el, toast } from './ui';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const FONT = 'font-family="Arial, Helvetica, sans-serif"';

/** SVG → PNG (data URL) с двойной плотностью. */
export function svgToPng(svg: string, w: number, h: number): Promise<string> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const c = document.createElement('canvas');
      c.width = Math.round(w * 2); c.height = Math.round(h * 2);
      const g = c.getContext('2d')!;
      g.scale(2, 2);
      g.drawImage(img, 0, 0, w, h);
      resolve(c.toDataURL('image/png'));
    };
    img.onerror = () => reject(new Error('Не удалось нарисовать изображение'));
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
  });
}

/** Вставить готовую картинку у курсора; href — ссылка на рисунке (видео). */
async function insertPicture(view: EditorView, svg: string, w: number, h: number, alt: string, href?: string): Promise<void> {
  const src = await svgToPng(svg, w, h);
  const { from, to } = view.state.selection;
  const img = schema.nodes.image.create({ src, width: Math.round(w), height: Math.round(h), alt }, null, href ? [schema.marks.link.create({ href })] : []);
  view.dispatch(view.state.tr.replaceWith(from, to, img).scrollIntoView());
  view.focus();
}

// --- Фигуры

export const SHAPES: Record<string, { label: string; draw: (w: number, h: number, a: string) => string }> = {
  rect: { label: 'Прямоугольник', draw: (w, h, a) => `<rect x="2" y="2" width="${w - 4}" height="${h - 4}" ${a}/>` },
  round: { label: 'Скруглённый прямоугольник', draw: (w, h, a) => `<rect x="2" y="2" width="${w - 4}" height="${h - 4}" rx="${Math.min(w, h) / 6}" ${a}/>` },
  ellipse: { label: 'Овал', draw: (w, h, a) => `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2 - 2}" ry="${h / 2 - 2}" ${a}/>` },
  triangle: { label: 'Треугольник', draw: (w, h, a) => `<polygon points="${w / 2},2 ${w - 2},${h - 2} 2,${h - 2}" ${a}/>` },
  diamond: { label: 'Ромб', draw: (w, h, a) => `<polygon points="${w / 2},2 ${w - 2},${h / 2} ${w / 2},${h - 2} 2,${h / 2}" ${a}/>` },
  arrow: { label: 'Стрелка вправо', draw: (w, h, a) => `<polygon points="2,${h * 0.3} ${w * 0.65},${h * 0.3} ${w * 0.65},2 ${w - 2},${h / 2} ${w * 0.65},${h - 2} ${w * 0.65},${h * 0.7} 2,${h * 0.7}" ${a}/>` },
  star: { label: 'Звезда', draw: (w, h, a) => {
    const pts: string[] = [];
    for (let i = 0; i < 10; i++) { const r = i % 2 ? 0.4 : 1, t = -Math.PI / 2 + i * Math.PI / 5; pts.push(`${w / 2 + Math.cos(t) * (w / 2 - 2) * r},${h / 2 + Math.sin(t) * (h / 2 - 2) * r}`); }
    return `<polygon points="${pts.join(' ')}" ${a}/>`;
  } },
  callout: { label: 'Выноска', draw: (w, h, a) => `<path d="M4 2 H${w - 4} Q${w - 2} 2 ${w - 2} 6 V${h * 0.72} Q${w - 2} ${h * 0.75} ${w - 6} ${h * 0.75} H${w * 0.35} L${w * 0.18} ${h - 2} L${w * 0.22} ${h * 0.75} H6 Q2 ${h * 0.75} 2 ${h * 0.72} V6 Q2 2 6 2 Z" ${a}/>` },
  line: { label: 'Линия', draw: (w, h, a) => `<line x1="2" y1="${h / 2}" x2="${w - 2}" y2="${h / 2}" ${a.replace(/fill="[^"]*"/, '')}/>` },
};

export function shapeSvg(kind: string, w: number, h: number, fill: string, line: string, text: string): string {
  const a = `fill="${kind === 'line' ? 'none' : fill}" stroke="${line}" stroke-width="2"`;
  const t = text.trim() ? `<text x="${w / 2}" y="${(kind === 'callout' ? h * 0.38 : h / 2)}" text-anchor="middle" dominant-baseline="middle" font-size="${Math.max(10, Math.min(18, h / 4))}" ${FONT} fill="${contrast(fill)}">${esc(text.trim())}</text>` : '';
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${SHAPES[kind].draw(w, h, a)}${t}</svg>`;
}
const contrast = (hex: string) => { const n = parseInt(hex.replace('#', ''), 16); const l = ((n >> 16) & 255) * 0.3 + ((n >> 8) & 255) * 0.59 + (n & 255) * 0.11; return l < 140 ? '#FFFFFF' : '#000000'; };

export function shapeDialog(view: EditorView): void {
  dialog((box, close) => {
    box.classList.add('graphics-dialog');
    let kind = 'rect';
    const fill = el('input', { type: 'color', value: '#4472C4', ariaLabel: 'Заливка' });
    const line = el('input', { type: 'color', value: '#2F528F', ariaLabel: 'Контур' });
    const wIn = el('input', { type: 'number', value: '4', min: '0.5', max: '17', step: '0.5', ariaLabel: 'Ширина, см' });
    const hIn = el('input', { type: 'number', value: '2.5', min: '0.5', max: '25', step: '0.5', ariaLabel: 'Высота, см' });
    const text = el('input', { type: 'text', placeholder: 'Текст внутри (необязательно)', ariaLabel: 'Текст фигуры' });
    const preview = el('div', { class: 'gfx-preview' });
    const grid = el('div', { class: 'gfx-grid' });
    const size = () => [Math.max(20, Number(wIn.value) * 37.8), Math.max(10, Number(hIn.value) * 37.8)];
    const draw = () => {
      grid.replaceChildren(...Object.entries(SHAPES).map(([k, s]) => el('button', { class: k === kind ? 'on' : '', title: s.label, ariaLabel: s.label,
        innerHTML: shapeSvg(k, 40, 28, fill.value, line.value, ''), onclick: () => { kind = k; draw(); } })));
      const [w, h] = size();
      preview.innerHTML = shapeSvg(kind, w, h, fill.value, line.value, text.value);
    };
    for (const c of [fill, line, wIn, hIn, text]) c.addEventListener('input', draw);
    box.append(el('h2', { textContent: 'Фигура' }), grid,
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Заливка и контур' }), fill, line),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Размер, см' }), wIn, el('span', { textContent: '×' }), hIn),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Текст' }), text), preview,
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Вставить', onclick: () => {
        const [w, h] = size();
        close();
        void insertPicture(view, shapeSvg(kind, w, h, fill.value, line.value, text.value), w, h, `Фигура: ${SHAPES[kind].label.toLowerCase()}${text.value.trim() ? ` — ${text.value.trim()}` : ''}`);
      } })));
    draw();
  }, false);
}

// --- Значки

export function iconSvg(raw: string, color: string, px: number): string {
  // Свои размер и цвет: у исходного значка уже есть width/height — заменяем, а не дублируем (иначе SVG не разберётся).
  const open = /<svg\b[^>]*>/.exec(raw)?.[0] ?? '<svg>';
  const attrs = open.replace(/\s(width|height|fill)="[^"]*"/g, '').replace(/^<svg/, `<svg width="${px}" height="${px}" fill="${color}"`);
  return raw.replace(open, attrs).replace(/fill="currentColor"/g, `fill="${color}"`);
}

export function iconsDialog(view: EditorView): void {
  dialog((box, close) => {
    box.classList.add('graphics-dialog');
    const search = el('input', { type: 'search', placeholder: 'Поиск: дом, почта, срок, закон…', ariaLabel: 'Поиск значка' });
    const color = el('input', { type: 'color', value: '#2F5496', ariaLabel: 'Цвет значка' });
    const size = el('select', { ariaLabel: 'Размер' }, new Option('Маленький (1 см)', '38'), new Option('Средний (2 см)', '76'), new Option('Крупный (3 см)', '113'));
    size.value = '76';
    const grid = el('div', { class: 'gfx-grid icons' });
    const draw = () => {
      const q = search.value.trim().toLowerCase();
      const list = ICON_SET.filter((i) => !q || i.words.includes(q) || i.name.includes(q));
      grid.replaceChildren(...(list.length ? list.map((i) => el('button', { title: i.words, ariaLabel: i.words, innerHTML: iconSvg(i.svg, color.value, 28), onclick: () => {
        const px = Number(size.value);
        close();
        void insertPicture(view, iconSvg(i.svg, color.value, px), px, px, `Значок: ${i.words.split(' ')[0]}`);
      } })) : [el('p', { textContent: 'Ничего не найдено.' })]));
    };
    search.addEventListener('input', draw);
    color.addEventListener('input', draw);
    box.append(el('h2', { textContent: 'Значки' }), search,
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Цвет и размер' }), color, size), grid,
      el('p', { class: 'hint', textContent: 'Щелчок вставляет значок у курсора. Набор Fluent UI (MIT), работает без интернета.' }),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Закрыть', onclick: close })));
    draw();
  }, false);
}

// --- Диаграммы

export interface ChartData { labels: string[]; series: { name: string; values: number[] }[] }
const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E'];
const num = (s: string) => { const v = Number(s.replace(/\s/g, '').replace(',', '.').replace(/[%₽$€]/g, '')); return isNaN(v) ? null : v; };

/** Данные из таблицы: первая строка — названия рядов, первый столбец — подписи, остальное — числа. */
export function chartDataFromTable(view: EditorView): ChartData | null {
  const t = tableAt(view.state);
  if (!t) return null;
  const rows: string[][] = [];
  t.node.forEach((r) => { const cells: string[] = []; r.forEach((c) => cells.push(c.textContent.trim())); rows.push(cells); });
  return parseRows(rows);
}
/** Текст «Метка; Ряд1; Ряд2» по строкам (первая строка — заголовки) → данные. */
export function chartDataFromText(text: string): ChartData | null {
  return parseRows(text.split('\n').map((l) => l.split(/\t|;/).map((c) => c.trim())).filter((r) => r.some(Boolean)));
}
function parseRows(rows: string[][]): ChartData | null {
  if (rows.length < 2) return null;
  const head = rows[0];
  const body = rows.slice(1).filter((r) => r.slice(1).some((c) => num(c) != null));
  const cols = head.slice(1).map((_, i) => i + 1).filter((c) => body.some((r) => num(r[c] ?? '') != null));
  if (!body.length || !cols.length) return null;
  return { labels: body.map((r) => r[0] ?? ''), series: cols.map((c) => ({ name: head[c] || `Ряд ${c}`, values: body.map((r) => num(r[c] ?? '') ?? 0) })) };
}

export function chartSvg(kind: 'bar' | 'line' | 'pie', d: ChartData, title: string, W = 520, H = 300): string {
  const t = title.trim() ? `<text x="${W / 2}" y="20" text-anchor="middle" font-size="15" font-weight="bold" ${FONT}>${esc(title.trim())}</text>` : '';
  const top = title.trim() ? 34 : 12;
  if (kind === 'pie') {
    const s = d.series[0], total = s.values.reduce((a, b) => a + Math.max(0, b), 0) || 1;
    const cx = W * 0.36, cy = top + (H - top) / 2, r = Math.min(W * 0.3, (H - top) / 2 - 8);
    let a0 = -Math.PI / 2, parts = '', legend = '';
    s.values.forEach((v, i) => {
      const a1 = a0 + (Math.max(0, v) / total) * Math.PI * 2;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      const p = (a: number) => `${cx + Math.cos(a) * r},${cy + Math.sin(a) * r}`;
      parts += s.values.length === 1 ? `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${PALETTE[0]}"/>` : `<path d="M${cx},${cy} L${p(a0)} A${r},${r} 0 ${large} 1 ${p(a1)} Z" fill="${PALETTE[i % PALETTE.length]}" stroke="#fff" stroke-width="1.5"/>`;
      legend += `<rect x="${W * 0.7}" y="${top + 8 + i * 20}" width="12" height="12" fill="${PALETTE[i % PALETTE.length]}"/><text x="${W * 0.7 + 18}" y="${top + 18 + i * 20}" font-size="12" ${FONT}>${esc(d.labels[i])} — ${Math.round((v / total) * 100)}%</text>`;
      a0 = a1;
    });
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><rect width="${W}" height="${H}" fill="#fff"/>${t}${parts}${legend}</svg>`;
  }
  const L = 52, R = d.series.length > 1 ? 130 : 16, B = 40, plotW = W - L - R, plotH = H - top - B;
  const all = d.series.flatMap((s) => s.values);
  const max = Math.max(0, ...all), min = Math.min(0, ...all);
  const step = niceStep((max - min) / 5 || 1);
  const hi = Math.ceil(max / step) * step, lo = Math.floor(min / step) * step;
  const y = (v: number) => top + plotH - ((v - lo) / (hi - lo || 1)) * plotH;
  let grid = '';
  for (let v = lo; v <= hi + step / 2; v += step) grid += `<line x1="${L}" x2="${L + plotW}" y1="${y(v)}" y2="${y(v)}" stroke="#E0E0E0"/><text x="${L - 6}" y="${y(v) + 4}" text-anchor="end" font-size="11" ${FONT}>${fmt(v)}</text>`;
  const n = d.labels.length, band = plotW / n;
  let marks = '';
  if (kind === 'bar') {
    const bw = (band * 0.7) / d.series.length;
    d.series.forEach((s, k) => s.values.forEach((v, i) => {
      const x = L + i * band + band * 0.15 + k * bw;
      marks += `<rect x="${x}" y="${Math.min(y(v), y(0))}" width="${bw - 1}" height="${Math.abs(y(v) - y(0))}" fill="${PALETTE[k % PALETTE.length]}"/>`;
    }));
  } else {
    d.series.forEach((s, k) => {
      const pts = s.values.map((v, i) => `${L + i * band + band / 2},${y(v)}`);
      marks += `<polyline points="${pts.join(' ')}" fill="none" stroke="${PALETTE[k % PALETTE.length]}" stroke-width="2.5"/>` + pts.map((p) => `<circle cx="${p.split(',')[0]}" cy="${p.split(',')[1]}" r="3.5" fill="${PALETTE[k % PALETTE.length]}"/>`).join('');
    });
  }
  const labels = d.labels.map((l, i) => `<text x="${L + i * band + band / 2}" y="${top + plotH + 16}" text-anchor="middle" font-size="11" ${FONT}>${esc(l.length > 14 ? l.slice(0, 13) + '…' : l)}</text>`).join('');
  const legend = d.series.length > 1 ? d.series.map((s, k) => `<rect x="${W - R + 10}" y="${top + 8 + k * 20}" width="12" height="12" fill="${PALETTE[k % PALETTE.length]}"/><text x="${W - R + 28}" y="${top + 18 + k * 20}" font-size="12" ${FONT}>${esc(s.name)}</text>`).join('') : '';
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><rect width="${W}" height="${H}" fill="#fff"/>${t}${grid}<line x1="${L}" x2="${L + plotW}" y1="${y(0)}" y2="${y(0)}" stroke="#808080"/>${marks}${labels}${legend}</svg>`;
}
function niceStep(raw: number): number { const p = 10 ** Math.floor(Math.log10(raw)); const f = raw / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p; }
const fmt = (v: number) => (Math.abs(v) >= 1000 ? v.toLocaleString('ru-RU') : String(Math.round(v * 100) / 100).replace('.', ','));

/** Замещающий текст диаграммы: вид, заголовок и все данные строками — их видно в Word и программам чтения. */
export const chartAlt = (kind: string, title: string, d: ChartData) =>
  `Диаграмма (${{ bar: 'гистограмма', line: 'график', pie: 'круговая' }[kind]})${title ? `: ${title}` : ''}. ` +
  `${['', ...d.series.map((s) => s.name)].join('; ')}\n` + d.labels.map((l, i) => [l, ...d.series.map((s) => String(s.values[i]))].join('; ')).join('\n');

export function chartDialog(view: EditorView): void {
  const fromTable = chartDataFromTable(view);
  dialog((box, close) => {
    box.classList.add('graphics-dialog');
    const kind = el('select', { ariaLabel: 'Вид диаграммы' }, new Option('Гистограмма', 'bar'), new Option('График', 'line'), new Option('Круговая', 'pie'));
    const title = el('input', { type: 'text', placeholder: 'Название диаграммы', ariaLabel: 'Название диаграммы' });
    const data = el('textarea', { rows: 6, ariaLabel: 'Данные', value: fromTable ? '' : 'Квартал; План; Факт\nI; 100; 92\nII; 120; 131\nIII; 110; 105\nIV; 140; 150' });
    const source = el('p', { class: 'hint' });
    const preview = el('div', { class: 'gfx-preview chart' });
    const current = (): ChartData | null => (data.value.trim() ? chartDataFromText(data.value) : fromTable);
    const draw = () => {
      const d = current();
      source.textContent = data.value.trim() ? 'Данные из поля ниже: первая строка — названия рядов, первый столбец — подписи; разделитель «;» или табуляция.' : 'Данные из таблицы под курсором. Чтобы ввести свои, заполните поле.';
      preview.innerHTML = d ? chartSvg(kind.value as 'bar', d, title.value) : '<p class="hint">Нет чисел для диаграммы.</p>';
    };
    for (const c of [kind, title, data]) { c.addEventListener('input', draw); c.addEventListener('change', draw); }
    box.append(el('h2', { textContent: 'Диаграмма' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Вид' }), kind),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Название' }), title), source, data, preview,
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Вставить', onclick: () => {
        const d = current();
        if (!d) { toast('Нет чисел для диаграммы.'); return; }
        close();
        // Из таблицы — диаграмма встаёт сразу под таблицей.
        const t = !data.value.trim() ? tableAt(view.state) : null;
        if (t) {
          const tr = view.state.tr.insert(t.pos + t.node.nodeSize, schema.nodes.paragraph.create({ align: 'center' }));
          view.dispatch(tr.setSelection(TextSelection.create(tr.doc, t.pos + t.node.nodeSize + 1)));
        }
        void insertPicture(view, chartSvg(kind.value as 'bar', d, title.value), 520, 300, chartAlt(kind.value, title.value.trim(), d));
      } })));
    draw();
  }, false);
}

// --- Схемы (вместо SmartArt): список строк → блоки

interface Item { text: string; level: number; kids: Item[] }
export function parseOutline(text: string): Item[] {
  const root: Item = { text: '', level: -1, kids: [] };
  const stack = [root];
  for (const raw of text.split('\n')) {
    if (!raw.trim()) continue;
    const level = Math.floor((raw.match(/^\s*/)![0].replace(/\t/g, '  ').length) / 2);
    const item: Item = { text: raw.trim().replace(/^[-•*]\s*/, ''), level, kids: [] };
    while (stack.length > 1 && stack[stack.length - 1].level >= level) stack.pop();
    stack[stack.length - 1].kids.push(item);
    stack.push(item);
  }
  return root.kids;
}

const box = (x: number, y: number, w: number, h: number, text: string, fill: string) =>
  `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="6" fill="${fill}" stroke="#2F528F" stroke-width="1.5"/>` +
  wrap(text, w - 10, 12).map((l, i, all) => `<text x="${x + w / 2}" y="${y + h / 2 + (i - (all.length - 1) / 2) * 15 + 4}" text-anchor="middle" font-size="12" ${FONT} fill="${contrast(fill)}">${esc(l)}</text>`).join('');
function wrap(t: string, w: number, size: number): string[] {
  const max = Math.max(4, Math.floor(w / (size * 0.55)));
  const out: string[] = [];
  let line = '';
  for (const word of t.split(/\s+/)) { if ((line + ' ' + word).trim().length > max && line) { out.push(line); line = word; } else line = (line + ' ' + word).trim(); }
  if (line) out.push(line);
  return out.slice(0, 3);
}

export function schemeSvg(kind: 'process' | 'hierarchy' | 'list', items: Item[]): { svg: string; w: number; h: number } {
  const fills = ['#4472C4', '#8FAADC', '#DAE3F3'];
  if (kind === 'process') {
    const n = items.length, bw = 120, bh = 60, gap = 34, w = n * bw + (n - 1) * gap + 8, h = bh + 8;
    let s = '';
    items.forEach((it, i) => {
      const x = 4 + i * (bw + gap);
      s += box(x, 4, bw, bh, it.text, fills[0]);
      if (i < n - 1) s += `<polygon points="${x + bw + 6},${4 + bh / 2 - 8} ${x + bw + gap - 6},${4 + bh / 2} ${x + bw + 6},${4 + bh / 2 + 8}" fill="#A5A5A5"/>`;
    });
    return { svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${s}</svg>`, w, h };
  }
  if (kind === 'list') {
    const rows: { it: Item; depth: number }[] = [];
    const walk = (list: Item[], d: number) => { for (const it of list) { rows.push({ it, depth: d }); walk(it.kids, d + 1); } };
    walk(items, 0);
    const w = 420, rh = 40, h = rows.length * (rh + 6) + 4;
    const s = rows.map((r, i) => box(4 + r.depth * 30, 4 + i * (rh + 6), w - 8 - r.depth * 30, rh, r.it.text, fills[Math.min(r.depth, 2)])).join('');
    return { svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${s}</svg>`, w, h };
  }
  // Иерархия (оргструктура): ширина поддерева — по числу листьев.
  const bw = 120, bh = 50, gx = 14, gy = 34;
  const leaves = (it: Item): number => (it.kids.length ? it.kids.reduce((a, k) => a + leaves(k), 0) : 1);
  const depth = (it: Item): number => 1 + Math.max(0, ...it.kids.map(depth));
  const roots = items.length === 1 ? items : [{ text: '', level: -1, kids: items }];
  const total = roots.reduce((a, r) => a + leaves(r), 0), levels = Math.max(...roots.map(depth));
  const w = total * (bw + gx) + gx, h = levels * (bh + gy) + 8;
  let s = '';
  const place = (it: Item, x0: number, d: number, px?: number, py?: number) => {
    const span = leaves(it) * (bw + gx);
    const cx = x0 + span / 2, y = 4 + d * (bh + gy);
    if (it.text) {
      if (px != null) s += `<path d="M${px},${py} V${py! + gy / 2} H${cx} V${y}" fill="none" stroke="#7F7F7F" stroke-width="1.5"/>`;
      s += box(cx - bw / 2, y, bw, bh, it.text, fills[Math.min(d, 2)]);
    }
    let x = x0;
    for (const k of it.kids) { place(k, x, it.text ? d + 1 : d, it.text ? cx : undefined, it.text ? y + bh : undefined); x += leaves(k) * (bw + gx); }
  };
  let x = gx / 2;
  for (const r of roots) { place(r, x, 0); x += leaves(r) * (bw + gx); }
  return { svg: `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}">${s}</svg>`, w, h };
}

export function schemeDialog(view: EditorView): void {
  dialog((box2, close) => {
    box2.classList.add('graphics-dialog');
    const kind = el('select', { ariaLabel: 'Вид схемы' }, new Option('Процесс (шаги слева направо)', 'process'), new Option('Иерархия (оргструктура)', 'hierarchy'), new Option('Список блоков', 'list'));
    const text = el('textarea', { rows: 7, ariaLabel: 'Пункты схемы', value: 'Заявка\nПроверка\nСогласование\nИсполнение' });
    const preview = el('div', { class: 'gfx-preview' });
    const draw = () => { const items = parseOutline(text.value); preview.innerHTML = items.length ? schemeSvg(kind.value as 'process', items).svg : ''; };
    kind.addEventListener('change', () => {
      if (kind.value === 'hierarchy' && !/\n\s+/.test(text.value)) text.value = 'Директор\n  Заместитель по экономике\n    Бухгалтерия\n    Плановый отдел\n  Заместитель по кадрам\n    Отдел кадров';
      draw();
    });
    text.addEventListener('input', draw);
    box2.append(el('h2', { textContent: 'Схема' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Вид' }), kind),
      el('p', { class: 'hint', textContent: 'Каждый пункт — с новой строки; подчинённые пункты — с отступом в два пробела.' }), text, preview,
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Вставить', onclick: () => {
        const items = parseOutline(text.value);
        if (!items.length) return;
        const r = schemeSvg(kind.value as 'process', items);
        // Шире полосы набора — уменьшаем пропорционально (около 17 см).
        const k = Math.min(1, 640 / r.w);
        close();
        void insertPicture(view, r.svg.replace(`width="${r.w}" height="${r.h}"`, `width="${r.w * k}" height="${r.h * k}"`), r.w * k, r.h * k, `Схема (${kind.selectedOptions[0].text.split(' (')[0].toLowerCase()}): ${text.value.trim().replace(/\n\s*/g, '; ')}`);
      } })));
    draw();
  }, false);
}

// --- Видео: обложка со ссылкой

export function videoSvg(title: string, url: string): string {
  const W = 480, H = 270;
  let host = '';
  try { host = new URL(url).hostname.replace(/^www\./, ''); } catch { host = ''; }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><rect width="${W}" height="${H}" rx="10" fill="#1F2937"/>` +
    `<circle cx="${W / 2}" cy="${H / 2 - 18}" r="38" fill="#FFFFFF" fill-opacity="0.92"/><polygon points="${W / 2 - 12},${H / 2 - 40} ${W / 2 + 22},${H / 2 - 18} ${W / 2 - 12},${H / 2 + 4}" fill="#C00000"/>` +
    `<text x="${W / 2}" y="${H - 52}" text-anchor="middle" font-size="18" font-weight="bold" fill="#FFFFFF" ${FONT}>${esc(title.slice(0, 48))}</text>` +
    `<text x="${W / 2}" y="${H - 26}" text-anchor="middle" font-size="13" fill="#D1D5DB" ${FONT}>${esc(host || 'видео')}</text></svg>`;
}

export function videoDialog(view: EditorView): void {
  dialog((box2, close) => {
    const url = el('input', { type: 'url', placeholder: 'https://…', ariaLabel: 'Адрес видео' });
    const title = el('input', { type: 'text', placeholder: 'Название видео', ariaLabel: 'Название видео' });
    const status = el('p', { role: 'status' });
    box2.append(el('h2', { textContent: 'Видео из интернета' }),
      el('p', { class: 'hint', textContent: 'В документ вставляется обложка со ссылкой: Ctrl+щелчок открывает видео. Так файл остаётся небольшим и открывается в любой программе.' }),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Адрес' }), url),
      el('div', { class: 'dlg-row' }, el('label', { textContent: 'Название' }), title), status,
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), el('button', { class: 'primary', textContent: 'Вставить', onclick: () => {
        const u = url.value.trim();
        if (!/^https?:\/\/\S+\.\S+/.test(u)) { status.textContent = 'Введите адрес, начинающийся с http:// или https://'; return; }
        const t = title.value.trim() || 'Видео';
        close();
        void insertPicture(view, videoSvg(t, u), 480, 270, `Видео: ${t} — ${u}`, u);
      } })));
  });
}
