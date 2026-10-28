// Рисование диаграмм (столбцы, линии, круговая) на canvas по данным ячеек.

import type { Chart, Model } from './engine';

export interface ChartData { cats: string[]; names: string[]; series: number[][] }

/** Оформление диаграммы: цвета рядов (у круговой — долей), легенда (0 — снизу, 1 — справа, 2 — нет), подписи значений. */
export interface ChartLook { colors: string[]; legend: number; labels: boolean }
export interface ChartTemplate extends ChartLook { name: string }

/** Встроенные шаблоны; свои пользователь сохраняет в настройках. */
export const CHART_TEMPLATES: ChartTemplate[] = [
  { name: 'Классика', colors: ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'], legend: 0, labels: false },
  { name: 'Мягкая', colors: ['#8faadc', '#f4b183', '#c9c9c9', '#ffd966', '#9dc3e6', '#a9d18e', '#b4a7d6', '#e6b8b7'], legend: 0, labels: false },
  { name: 'Контраст', colors: ['#c00000', '#0070c0', '#00b050', '#ffc000', '#7030a0', '#000000', '#ff6600', '#00b0f0'], legend: 1, labels: true },
  { name: 'Зелёная гамма', colors: ['#1e5e3b', '#2e8b57', '#6bbf8a', '#a8dcb9', '#d5efdd', '#8a9a5b', '#4f7942', '#c5e1a5'], legend: 0, labels: false },
  { name: 'Синяя гамма', colors: ['#0b2a5b', '#1f4e9c', '#4472c4', '#7fa1dd', '#b4c7e7', '#2f5597', '#8faadc', '#d9e2f3'], legend: 0, labels: false },
  { name: 'Строгая (ч/б)', colors: ['#222222', '#666666', '#999999', '#bbbbbb', '#dddddd', '#444444', '#808080', '#aaaaaa'], legend: 1, labels: true },
];

export const CHART_COLORS = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];

/** «A1:C5» → [строка1, столбец1, строка2, столбец2] или null. */
export function parseRange(range: string): [number, number, number, number] | null {
  const m = /^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/.exec(range.trim().toUpperCase());
  if (!m) return null;
  const col = (s: string) => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const [r1, r2, c1, c2] = [Number(m[2]), Number(m[4]), col(m[1]), col(m[3])];
  return r2 > r1 && c2 > c1 ? [Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2)] : null;
}

/** Круглые деления шкалы: [минимум, максимум, шаг]. */
function niceScale(min: number, max: number): [number, number, number] {
  if (min === max) { max = min + 1; }
  const raw = (max - min) / 5, pow = 10 ** Math.floor(Math.log10(raw));
  const step = ([1, 2, 5, 10].find((k) => k * pow >= raw) ?? 10) * pow;
  return [Math.floor(min / step) * step, Math.ceil(max / step) * step, step];
}

const fmt = (n: number) => (Math.abs(n) >= 1e6 ? `${+(n / 1e6).toPrecision(3)}M` : String(+n.toPrecision(6))).replace('.', ',');

export function drawChart(ctx: CanvasRenderingContext2D, r: { x: number; y: number; w: number; h: number },
  kind: number, title: string, d: ChartData, look: Partial<ChartLook> = {}, textColor = '#333'): void {
  const color = (i: number) => look.colors?.[i] ?? CHART_COLORS[i % CHART_COLORS.length];
  const legendPos = look.legend ?? 0, labels = !!look.labels;
  ctx.save();
  ctx.beginPath(); ctx.rect(r.x, r.y, r.w, r.h); ctx.clip();
  ctx.fillStyle = '#fff'; ctx.fillRect(r.x, r.y, r.w, r.h);
  ctx.strokeStyle = '#c8c8c8'; ctx.lineWidth = 1; ctx.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 1, r.h - 1);
  ctx.fillStyle = textColor; ctx.textBaseline = 'middle';
  let top = r.y + 8;
  if (title) {
    ctx.font = 'bold 14px sans-serif'; ctx.textAlign = 'center';
    ctx.fillText(title, r.x + r.w / 2, r.y + 16, r.w - 16);
    top = r.y + 32;
  }
  // Легенда: снизу, справа или без неё; ряды (для круговой — категории).
  const legend = kind === 2 ? d.cats : d.names;
  ctx.font = '11px sans-serif'; ctx.textAlign = 'left';
  const legendRows: [string, number, number, number][] = [];
  let right = r.x + r.w, bottom = r.y + r.h - 8, extra = 0;
  if (legendPos === 1) {
    const wMax = Math.min(r.w * 0.35, Math.max(0, ...legend.map((t) => ctx.measureText(t).width)) + 26);
    legend.forEach((name, i) => legendRows.push([name, i, r.x + r.w - wMax - 4, top + 8 + i * 17]));
    right = r.x + r.w - wMax - 10;
  } else if (legendPos === 0) {
    let lx = r.x + 10, ly = r.y + r.h - 12;
    for (const [i, name] of legend.entries()) {
      const w = ctx.measureText(name).width + 24;
      if (lx + w > r.x + r.w - 6 && lx > r.x + 10) { lx = r.x + 10; ly += 16; }
      legendRows.push([name, i, lx, ly]);
      lx += w;
    }
    extra = ly - (r.y + r.h - 12);
    bottom = r.y + r.h - 28 - extra;
  }
  for (const [name, i, x, y] of legendRows) {
    const yy = legendPos === 0 ? y - extra : y;
    ctx.fillStyle = color(i); ctx.fillRect(x, yy - 5, 10, 10);
    ctx.fillStyle = textColor; ctx.fillText(name, x + 14, yy, Math.max(20, r.x + r.w - x - 18));
  }
  const n = d.cats.length;
  if (kind === 2) {
    const vals = (d.series[0] ?? []).map((v) => (Number.isFinite(v) && v > 0 ? v : 0));
    const total = vals.reduce((a, b) => a + b, 0);
    const cx = (r.x + right) / 2, cy = (top + bottom) / 2, rad = Math.max(10, Math.min((right - r.x) / 2 - 20, (bottom - top) / 2 - 4));
    if (total > 0) {
      let a = -Math.PI / 2;
      vals.forEach((v, i) => {
        const da = (v / total) * Math.PI * 2;
        ctx.fillStyle = color(i);
        ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, rad, a, a + da); ctx.closePath(); ctx.fill();
        ctx.strokeStyle = '#fff'; ctx.stroke();
        if (labels && v / total > 0.04) {
          ctx.fillStyle = '#fff'; ctx.font = 'bold 11px sans-serif'; ctx.textAlign = 'center';
          ctx.fillText(`${Math.round((v / total) * 100)}%`, cx + Math.cos(a + da / 2) * rad * 0.65, cy + Math.sin(a + da / 2) * rad * 0.65);
        }
        a += da;
      });
    }
    ctx.restore();
    return;
  }
  const all = d.series.flat().filter(Number.isFinite);
  const [lo, hi, step] = niceScale(Math.min(0, ...all, 0), Math.max(0, ...all, 0));
  ctx.font = '11px sans-serif';
  const labelW = Math.max(...[lo, hi].map((v) => ctx.measureText(fmt(v)).width)) + 8;
  const px0 = r.x + labelW + 8, px1 = right - 12, py0 = top + 4, py1 = bottom - 18;
  const yOf = (v: number) => py1 - ((v - lo) / (hi - lo)) * (py1 - py0);
  ctx.textAlign = 'right';
  for (let v = lo; v <= hi + step / 2; v += step) {
    const y = Math.round(yOf(v)) + 0.5;
    ctx.strokeStyle = '#e3e3e3'; ctx.beginPath(); ctx.moveTo(px0, y); ctx.lineTo(px1, y); ctx.stroke();
    ctx.fillStyle = textColor; ctx.fillText(fmt(v), px0 - 6, y);
  }
  const slot = (px1 - px0) / Math.max(1, n);
  ctx.textAlign = 'center';
  d.cats.forEach((c, i) => { ctx.fillStyle = textColor; ctx.fillText(c, px0 + slot * (i + 0.5), py1 + 12, slot - 2); });
  const zero = yOf(0);
  if (kind === 0) {
    const k = Math.max(1, d.series.length), bw = Math.min(40, (slot * 0.7) / k);
    d.series.forEach((vals, s) => {
      ctx.fillStyle = color(s);
      vals.forEach((v, i) => {
        if (!Number.isFinite(v)) return;
        const x = px0 + slot * (i + 0.5) - (bw * k) / 2 + bw * s, y = yOf(v);
        ctx.fillRect(x, Math.min(y, zero), bw - 1, Math.abs(zero - y));
        if (labels) { ctx.fillStyle = textColor; ctx.font = '10px sans-serif'; ctx.textAlign = 'center'; ctx.fillText(fmt(v), x + bw / 2, v >= 0 ? y - 7 : y + 8); ctx.fillStyle = color(s); }
      });
    });
  } else {
    d.series.forEach((vals, s) => {
      const col = color(s);
      ctx.strokeStyle = col; ctx.fillStyle = col; ctx.lineWidth = 2; ctx.beginPath();
      let pen = false;
      vals.forEach((v, i) => {
        if (!Number.isFinite(v)) { pen = false; return; }
        const x = px0 + slot * (i + 0.5), y = yOf(v);
        if (pen) ctx.lineTo(x, y); else ctx.moveTo(x, y);
        pen = true;
      });
      ctx.stroke(); ctx.lineWidth = 1;
      vals.forEach((v, i) => { if (Number.isFinite(v)) { ctx.beginPath(); ctx.arc(px0 + slot * (i + 0.5), yOf(v), 3, 0, Math.PI * 2); ctx.fill(); if (labels) { ctx.fillStyle = textColor; ctx.font = '10px sans-serif'; ctx.textAlign = 'center'; ctx.fillText(fmt(v), px0 + slot * (i + 0.5), yOf(v) - 9); ctx.fillStyle = col; } } });
    });
  }
  ctx.strokeStyle = '#888'; ctx.beginPath(); ctx.moveTo(px0, Math.round(zero) + 0.5); ctx.lineTo(px1, Math.round(zero) + 0.5); ctx.stroke();
  ctx.restore();
}

/** Данные диаграммы из ячеек её диапазона (первая строка — имена рядов, первый столбец — подписи). */
export function readChartData(m: Model, sheet: number, c: Chart): ChartData {
  const empty: ChartData = { cats: [], names: [], series: [] };
  const rg = parseRange(c.range);
  if (!rg) return empty;
  const [r1, c1, r2, c2] = rg;
  const cats: string[] = [];
  for (let r = r1 + 1; r <= r2; r++) cats.push(m.getFormattedCellValue(sheet, r, c1));
  const names: string[] = [], series: number[][] = [];
  const cols = c.kind === 2 ? 1 : c2 - c1;
  for (let k = 1; k <= cols; k++) {
    names.push(m.getFormattedCellValue(sheet, r1, c1 + k));
    series.push(Array.from(m.getRangeNumbers(sheet, r1 + 1, c1 + k, r2, c1 + k)));
  }
  return { cats, names, series };
}
