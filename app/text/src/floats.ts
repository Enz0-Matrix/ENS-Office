// Плавающие рисунки: «за текстом», «перед текстом» (положение от страницы, поля или абзаца) и с обтеканием справа или слева.
// Они не занимают места в строке, поэтому колонтитулы с логотипом не раздуваются. Положение считается после вёрстки страниц.

import type { PageBox } from './pagination';
import { PX } from './pagination';

export interface Float {
  /** behind — за текстом, front — перед текстом, wrap — обтекание (рисунок прижат к краю абзаца). */
  mode: 'behind' | 'front' | 'wrap';
  /** Для wrap: с какой стороны рисунок. */
  side?: 'left' | 'right';
  hRel: 'page' | 'margin' | 'column' | 'character';
  /** Смещение от точки отсчёта, px; null — по выравниванию. */
  hOff: number | null;
  hAlign: 'left' | 'center' | 'right' | null;
  vRel: 'page' | 'margin' | 'paragraph' | 'line';
  vOff: number | null;
  vAlign: 'top' | 'center' | 'bottom' | null;
  /** Отступы текста от рисунка: сверху, снизу, слева, справа, px. */
  dist: [number, number, number, number];
}

export const FLOAT_CLASS: Record<Float['mode'], string> = { behind: 'fl fl-behind', front: 'fl fl-front', wrap: 'fl fl-wrap' };

/** Начальный CSS рисунка (до расчёта положения). */
export function floatStyle(f: Float, w: number | null, h: number | null): string {
  const s: string[] = ['max-width:none'];
  if (w && h) s.push(`width:${Math.round(w)}px`, `height:${Math.round(h)}px`);
  if (f.mode === 'wrap') {
    s.push(`float:${f.side === 'right' ? 'right' : 'left'}`, `margin:${f.dist.map((x) => `${Math.round(x)}px`).join(' ')}`);
  }
  return s.join(';');
}

interface Ctx {
  /** Корень, внутри которого лежит рисунок, и у него position: relative/absolute (его границы — начало координат). */
  root: HTMLElement;
  boxes: PageBox[];
  sheetW: number;
  /** Масштаб экрана (zoom листа). */
  scale: number;
  /** Для колонтитулов: страница известна заранее и начало координат — угол листа. */
  fixed?: { box: PageBox; originX: number; originY: number };
}

const num = (v: number | null | undefined) => (typeof v === 'number' && Number.isFinite(v) ? v : 0);

/** Расставить плавающие рисунки (за и перед текстом) внутри root по параметрам их страницы. */
export function placeFloats(ctx: Ctx): void {
  const rootRect = ctx.root.getBoundingClientRect();
  for (const img of ctx.root.querySelectorAll<HTMLImageElement>('img.fl-behind, img.fl-front')) {
    let f: Float;
    try { f = JSON.parse(img.dataset.float ?? 'null') as Float; } catch { continue; }
    if (!f) continue;
    const w = parseFloat(img.style.width) || img.offsetWidth;
    const h = parseFloat(img.style.height) || img.offsetHeight;
    const para = img.closest('p, h1, h2, h3, h4, h5, h6, li') as HTMLElement | null;
    const paraRect = (para ?? img).getBoundingClientRect();
    const paraY = (paraRect.top - rootRect.top) / ctx.scale;
    const paraX = (paraRect.left - rootRect.left) / ctx.scale;

    let box = ctx.fixed?.box;
    let ox = ctx.fixed?.originX ?? 0;
    let oy = ctx.fixed?.originY ?? 0;
    if (!box) {
      let idx = 0;
      for (let i = 0; i < ctx.boxes.length; i++) if (ctx.boxes[i].top <= paraY + 1) idx = i;
      box = ctx.boxes[idx];
      if (!box) continue;
      ox = (ctx.sheetW - box.W) / 2;
      oy = box.top;
    }
    const pg = box.page;
    const L = pg.left * PX, R = pg.right * PX, T = pg.top * PX, B = pg.bottom * PX;

    // По горизонтали: от страницы, поля (колонки) или от абзаца.
    let b0 = 0, b1 = box.W;
    if (f.hRel === 'margin' || f.hRel === 'column') { b0 = L; b1 = box.W - R; }
    let x: number;
    if (f.hRel === 'character') x = paraX - ox + num(f.hOff);
    else if (f.hAlign === 'left') x = b0;
    else if (f.hAlign === 'center') x = (b0 + b1 - w) / 2;
    else if (f.hAlign === 'right') x = b1 - w;
    else x = b0 + num(f.hOff);

    // По вертикали: от страницы, поля или абзаца.
    let y: number;
    if (f.vRel === 'paragraph' || f.vRel === 'line') y = paraY - oy + num(f.vOff);
    else {
      let c0 = 0, c1 = box.H;
      if (f.vRel === 'margin') { c0 = T; c1 = box.H - B; }
      if (f.vAlign === 'top') y = c0;
      else if (f.vAlign === 'center') y = (c0 + c1 - h) / 2;
      else if (f.vAlign === 'bottom') y = c1 - h;
      else y = c0 + num(f.vOff);
    }
    img.style.left = `${ox + x}px`;
    img.style.top = `${oy + y}px`;
  }
}
