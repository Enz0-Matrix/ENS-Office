// Рисование фигур и надписей на canvas (прямоугольники, эллипс, линии, стрелки, текст внутри).

import type { Shape } from './engine';

export const SHAPE_NAMES = ['Прямоугольник', 'Скруглённый прямоугольник', 'Эллипс', 'Линия', 'Стрелка', 'Надпись'];

/** Новая фигура с настройками по умолчанию: заливка и контур в цвет Excel, у надписи — без них. */
export function newShape(kind: number, row: number, column: number): Shape {
  const line = kind === 3 || kind === 4;
  return {
    kind, text: '', fill: kind === 5 || line ? '' : '#DDEBF7', line_color: kind === 5 ? '' : line ? '#000000' : '#2F75B5',
    line_width: line ? 2 : 1, text_color: '', font_size: 0, bold: false, align: kind === 5 ? 0 : 1, flip: false,
    row, column, dx: 4, dy: 4, width: line ? 140 : kind === 5 ? 160 : 140, height: line ? 0 : kind === 5 ? 50 : 80,
  };
}

function words(ctx: CanvasRenderingContext2D, text: string, maxW: number): string[] {
  const out: string[] = [];
  for (const para of text.split('\n')) {
    let cur = '';
    for (const w of para.split(/(\s+)/)) {
      if (cur && ctx.measureText(cur + w).width > maxW && w.trim()) { out.push(cur.trimEnd()); cur = w; } else cur += w;
    }
    out.push(cur.trimEnd());
  }
  return out;
}

export function drawShape(ctx: CanvasRenderingContext2D, r: { x: number; y: number; w: number; h: number }, s: Shape, defaultText = '#000000'): void {
  ctx.save();
  const lw = Math.max(1, s.line_width || 1);
  if (s.kind === 3 || s.kind === 4) {
    const x1 = r.x, x2 = r.x + r.w, y1 = s.flip ? r.y + r.h : r.y, y2 = s.flip ? r.y : r.y + r.h;
    ctx.strokeStyle = s.line_color || '#000000'; ctx.lineWidth = lw; ctx.lineCap = 'round';
    ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke();
    if (s.kind === 4) {
      const a = Math.atan2(y2 - y1, x2 - x1), len = 8 + lw * 2;
      ctx.fillStyle = s.line_color || '#000000';
      ctx.beginPath();
      ctx.moveTo(x2, y2);
      ctx.lineTo(x2 - len * Math.cos(a - 0.4), y2 - len * Math.sin(a - 0.4));
      ctx.lineTo(x2 - len * Math.cos(a + 0.4), y2 - len * Math.sin(a + 0.4));
      ctx.closePath(); ctx.fill();
    }
    ctx.restore();
    return;
  }
  ctx.beginPath();
  if (s.kind === 2) ctx.ellipse(r.x + r.w / 2, r.y + r.h / 2, Math.max(1, r.w / 2), Math.max(1, r.h / 2), 0, 0, Math.PI * 2);
  else if (s.kind === 1) {
    const k = Math.min(r.w, r.h) * 0.1667;
    ctx.moveTo(r.x + k, r.y); ctx.arcTo(r.x + r.w, r.y, r.x + r.w, r.y + r.h, k); ctx.arcTo(r.x + r.w, r.y + r.h, r.x, r.y + r.h, k);
    ctx.arcTo(r.x, r.y + r.h, r.x, r.y, k); ctx.arcTo(r.x, r.y, r.x + r.w, r.y, k); ctx.closePath();
  } else ctx.rect(r.x, r.y, r.w, r.h);
  if (s.fill) { ctx.fillStyle = s.fill; ctx.fill(); }
  if (s.line_color) { ctx.strokeStyle = s.line_color; ctx.lineWidth = lw; ctx.stroke(); }
  if (s.text) {
    const px = Math.round((s.font_size || 11) * 1.33);
    ctx.font = `${s.bold ? 'bold ' : ''}${px}px sans-serif`;
    ctx.fillStyle = s.text_color || defaultText;
    ctx.textBaseline = 'middle';
    const pad = 6, lines = words(ctx, s.text, Math.max(10, r.w - pad * 2)), lh = Math.round(px * 1.25);
    ctx.beginPath(); ctx.rect(r.x + 1, r.y + 1, r.w - 2, r.h - 2); ctx.clip();
    let y = s.kind === 5 ? r.y + pad + lh / 2 : r.y + (r.h - lines.length * lh) / 2 + lh / 2;
    ctx.textAlign = s.align === 2 ? 'right' : s.align === 1 ? 'center' : 'left';
    const x = s.align === 2 ? r.x + r.w - pad : s.align === 1 ? r.x + r.w / 2 : r.x + pad;
    for (const l of lines) { ctx.fillText(l, x, y); y += lh; }
  }
  ctx.restore();
}
