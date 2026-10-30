// Работа с рисунками: ручки размера и обрезки мышью, перетаскивание плавающих рисунков, окно «Формат рисунка», подписи, замена.
// Рисунок выделяется щелчком (Alt+щелчок — и рисунок «за текстом», до которого иначе не добраться).

import { NodeSelection, TextSelection } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import type { PageBox } from './pagination';
import { PX } from './pagination';
import type { Float } from './floats';
import { schema } from './schema';
import { dialog, el, toast } from './ui';

export interface ImgHost {
  view: () => EditorView;
  sheet: HTMLElement;
  boxes: () => PageBox[];
  /** Увеличение листа (zoom). */
  scale: () => number;
  /** Пересчитать положение плавающих рисунков. */
  relayout: () => void;
  /** Экран «правка колонтитула»: рисунки основного текста не трогаем. */
  locked: () => boolean;
  /** Ширина полосы текста текущей страницы, px. */
  textWidth: () => number;
  update: () => void;
}

type Handle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';
const HANDLES: Handle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const CM = 96 / 2.54;

let host: ImgHost;
let overlay: HTMLElement;
let cropMode = false;

export const isCropMode = (): boolean => cropMode;

export function selectedImage(v: EditorView = host.view()): { pos: number; node: PMNode; dom: HTMLImageElement | null } | null {
  const sel = v.state.selection;
  if (sel instanceof NodeSelection && sel.node.type === schema.nodes.image) {
    const dom = v.nodeDOM(sel.from);
    return { pos: sel.from, node: sel.node, dom: dom instanceof HTMLImageElement ? dom : null };
  }
  return null;
}

const floatOf = (n: PMNode): Float | null => n.attrs.float as Float | null;
const isAbs = (f: Float | null) => !!f && (f.mode === 'behind' || f.mode === 'front');

/** Исходный размер рисунка, px (по самой картинке). */
export function naturalSize(src: string): Promise<{ w: number; h: number }> {
  return new Promise((res) => {
    const i = new Image();
    i.onload = () => res({ w: i.naturalWidth || 200, h: i.naturalHeight || 150 });
    i.onerror = () => res({ w: 200, h: 150 });
    i.src = src;
  });
}

function setAttrs(pos: number, node: PMNode, next: Record<string, unknown>): void {
  const v = host.view();
  const tr = v.state.tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...next });
  // Рисунок остаётся выделенным: рамка с ручками не пропадает.
  v.dispatch(tr.setSelection(NodeSelection.create(tr.doc, pos)));
}

// ---------------------------------------------------------------- рамка с ручками

/** Положение рамки: по прямоугольнику выделенного рисунка в координатах листа. */
export function updateOverlay(): void {
  if (!overlay) return;
  const img = !host.locked() ? selectedImage() : null;
  if (!img?.dom) { overlay.hidden = true; return; }
  const s = host.sheet.getBoundingClientRect();
  const r = img.dom.getBoundingClientRect();
  const k = host.scale();
  overlay.hidden = false;
  overlay.classList.toggle('crop', cropMode);
  overlay.classList.toggle('abs', isAbs(floatOf(img.node)));
  Object.assign(overlay.style, { left: `${(r.left - s.left) / k}px`, top: `${(r.top - s.top) / k}px`, width: `${r.width / k}px`, height: `${r.height / k}px` });
}

interface Drag {
  kind: 'size' | 'crop' | 'move';
  handle?: Handle;
  sx: number;
  sy: number;
  w: number;
  h: number;
  crop: number[];
  /** Положение левого верхнего угла рисунка на листе в начале перетаскивания, px. */
  left: number;
  top: number;
  pos: number;
  node: PMNode;
  dom: HTMLImageElement;
  /** Последнее значение: что записать при отпускании. */
  out?: { w: number; h: number; crop: number[] };
  moved?: { left: number; top: number };
}

let drag: Drag | null = null;

function startDrag(e: MouseEvent, kind: Drag['kind'], handle?: Handle): void {
  const img = selectedImage();
  if (!img?.dom) return;
  e.preventDefault();
  e.stopPropagation();
  const s = host.sheet.getBoundingClientRect();
  const r = img.dom.getBoundingClientRect();
  const k = host.scale();
  drag = {
    kind, handle, sx: e.clientX, sy: e.clientY, w: r.width / k, h: r.height / k,
    crop: (img.node.attrs.crop as number[] | null) ?? [0, 0, 0, 0],
    left: (r.left - s.left) / k, top: (r.top - s.top) / k, pos: img.pos, node: img.node, dom: img.dom,
  };
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp, { once: true });
}

const clamp = (v: number, a: number, b: number) => Math.min(b, Math.max(a, v));

function onMove(e: MouseEvent): void {
  const d = drag;
  if (!d) return;
  const k = host.scale();
  const dx = (e.clientX - d.sx) / k;
  const dy = (e.clientY - d.sy) / k;
  const dom = d.dom;
  const h = d.handle!;
  const west = h.includes('w'), east = h.includes('e'), north = h.includes('n'), south = h.includes('s');
  const crop = [...d.crop];
  let w = d.w, hh = d.h;
  if (d.kind === 'size') {
    if (east) w = d.w + dx; else if (west) w = d.w - dx;
    if (south) hh = d.h + dy; else if (north) hh = d.h - dy;
    const corner = (west || east) && (north || south);
    if (corner && !e.shiftKey) {
      const f = Math.max(w / d.w, hh / d.h);
      w = d.w * f;
      hh = d.h * f;
    }
    w = Math.max(12, w);
    hh = Math.max(12, hh);
  } else {
    // Обрезка: двигаем край, рисунок «вырезается» без изменения масштаба.
    const full = { w: d.w / (1 - crop[0] - crop[2]), h: d.h / (1 - crop[1] - crop[3]) };
    if (west) { const c = clamp(d.crop[0] + dx / full.w, 0, 0.95 - d.crop[2]); crop[0] = c; w = full.w * (1 - crop[0] - crop[2]); }
    if (east) { const c = clamp(d.crop[2] - dx / full.w, 0, 0.95 - d.crop[0]); crop[2] = c; w = full.w * (1 - crop[0] - crop[2]); }
    if (north) { const c = clamp(d.crop[1] + dy / full.h, 0, 0.95 - d.crop[3]); crop[1] = c; hh = full.h * (1 - crop[1] - crop[3]); }
    if (south) { const c = clamp(d.crop[3] - dy / full.h, 0, 0.95 - d.crop[1]); crop[3] = c; hh = full.h * (1 - crop[1] - crop[3]); }
  }
  d.out = { w, h: hh, crop };
  dom.style.width = `${w}px`;
  dom.style.height = `${hh}px`;
  dom.style.aspectRatio = `${w} / ${hh}`;
  if (d.kind === 'crop') dom.style.setProperty('object-view-box', `inset(${crop[1] * 100}% ${crop[2] * 100}% ${crop[3] * 100}% ${crop[0] * 100}%)`);
  updateOverlay();
}

function onUp(): void {
  document.removeEventListener('mousemove', onMove);
  const d = drag;
  drag = null;
  if (!d) return;
  if (!d.out) return;
  const next: Record<string, unknown> = { width: Math.round(d.out.w), height: Math.round(d.out.h) };
  if (d.kind === 'crop') next.crop = d.out.crop.some((x) => x > 0.0005) ? d.out.crop.map((x) => Math.round(x * 1e5) / 1e5) : null;
  setAttrs(d.pos, d.node, next);
  host.relayout();
}

// ---------------------------------------------------------------- перетаскивание плавающего рисунка

/** Рисунок «за/перед текстом» перетаскивается целиком; запоминается положение от угла страницы, на которую он попал. */
function finishMove(d: Drag): void {
  const k = host.scale();
  const m = d.moved;
  if (!m) { host.relayout(); return; }
  const boxes = host.boxes();
  const sheetW = host.sheet.offsetWidth;
  const cy = m.top + d.h / 2;
  let idx = 0;
  for (let i = 0; i < boxes.length; i++) if (boxes[i].top <= cy) idx = i;
  const box = boxes[idx];
  if (!box) { host.relayout(); return; }
  const side = (sheetW - box.W) / 2;
  const f = floatOf(d.node)!;
  const nf: Float = { ...f, hRel: 'page', hOff: Math.round(m.left - side), hAlign: null, vRel: 'page', vOff: Math.round(m.top - box.top), vAlign: null };
  const v = host.view();
  // Если рисунок перенесли на другую страницу, он «привязывается» к абзацу рядом с новым местом.
  const cur = (() => {
    const para = v.nodeDOM(d.pos)?.parentElement;
    const r = para?.getBoundingClientRect();
    const s = host.sheet.getBoundingClientRect();
    return r ? (r.top - s.top) / k : 0;
  })();
  let curIdx = 0;
  for (let i = 0; i < boxes.length; i++) if (boxes[i].top <= cur + 1) curIdx = i;
  let tr = v.state.tr;
  if (curIdx !== idx) {
    const s = host.sheet.getBoundingClientRect();
    const at = v.posAtCoords({ left: s.left + (side + box.page.left * PX + 20) * k, top: s.top + (m.top + 4) * k });
    if (at) {
      const target = at.pos;
      const node = d.node.type.create({ ...d.node.attrs, float: nf });
      tr = tr.delete(d.pos, d.pos + 1);
      const mapped = tr.mapping.map(target);
      tr = tr.insert(mapped, node);
      v.dispatch(tr.setSelection(NodeSelection.create(tr.doc, mapped)));
      host.relayout();
      return;
    }
  }
  setAttrs(d.pos, d.node, { float: nf });
  host.relayout();
}

// ---------------------------------------------------------------- события

function build(): void {
  overlay = document.createElement('div');
  overlay.className = 'img-ov';
  overlay.hidden = true;
  const body = document.createElement('div');
  body.className = 'img-ov-body';
  body.addEventListener('mousedown', (e) => {
    const img = selectedImage();
    if (img && isAbs(floatOf(img.node)) && img.dom) {
      // Плавающий рисунок: перетаскиваем «живьём» — сдвигаем рисунок и рамку.
      e.preventDefault();
      e.stopPropagation();
      const k = host.scale();
      const r = img.dom.getBoundingClientRect();
      const s = host.sheet.getBoundingClientRect();
      const start = { x: e.clientX, y: e.clientY, l: parseFloat(img.dom.style.left || '0'), t: parseFloat(img.dom.style.top || '0') };
      const d: Drag = {
        kind: 'move', sx: e.clientX, sy: e.clientY, w: r.width / k, h: r.height / k, crop: [0, 0, 0, 0],
        left: (r.left - s.left) / k, top: (r.top - s.top) / k, pos: img.pos, node: img.node, dom: img.dom,
      };
      drag = d;
      const mv = (ev: MouseEvent) => {
        const dx = (ev.clientX - start.x) / k, dy = (ev.clientY - start.y) / k;
        img.dom!.style.left = `${start.l + dx}px`;
        img.dom!.style.top = `${start.t + dy}px`;
        d.moved = { left: d.left + dx, top: d.top + dy };
        updateOverlay();
      };
      const up = () => { document.removeEventListener('mousemove', mv); drag = null; finishMove(d); };
      document.addEventListener('mousemove', mv);
      document.addEventListener('mouseup', up, { once: true });
    }
  });
  overlay.append(body);
  for (const h of HANDLES) {
    const n = document.createElement('div');
    n.className = `img-h img-h-${h}`;
    n.addEventListener('mousedown', (e) => startDrag(e, cropMode ? 'crop' : 'size', h));
    overlay.append(n);
  }
  host.sheet.append(overlay);
}

export function initImageTools(h: ImgHost): void {
  host = h;
  build();
  // Alt+щелчок: выбрать рисунок под курсором, в том числе «за текстом».
  host.sheet.addEventListener('mousedown', (e) => {
    if (!e.altKey || host.locked()) return;
    const v = host.view();
    const imgs = [...v.dom.querySelectorAll<HTMLImageElement>('img[src]')].filter((i) => {
      const r = i.getBoundingClientRect();
      return e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom;
    });
    if (!imgs.length) return;
    imgs.sort((a, b) => a.width * a.height - b.width * b.height);
    const pos = v.posAtDOM(imgs[0], 0);
    if (!(v.state.doc.nodeAt(pos)?.type === schema.nodes.image)) return;
    e.preventDefault();
    e.stopPropagation();
    v.dispatch(v.state.tr.setSelection(NodeSelection.create(v.state.doc, pos)));
    v.focus();
  }, true);
  // Alt+стрелки сдвигают плавающий рисунок на 1 px (с Shift — на 10).
  host.sheet.addEventListener('keydown', (e) => {
    if (!e.altKey || !e.key.startsWith('Arrow')) return;
    const img = selectedImage();
    const f = img ? floatOf(img.node) : null;
    if (!img || !f || !isAbs(f)) return;
    e.preventDefault();
    const step = e.shiftKey ? 10 : 1;
    const dx = e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0;
    const dy = e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0;
    setAttrs(img.pos, img.node, { float: { ...f, hAlign: null, vAlign: null, hOff: (f.hOff ?? 0) + dx, vOff: (f.vOff ?? 0) + dy } });
    host.relayout();
  }, true);
  window.addEventListener('resize', updateOverlay);
  host.sheet.parentElement?.addEventListener('scroll', updateOverlay);
}

export function setCropMode(on: boolean): void {
  cropMode = on && !!selectedImage();
  updateOverlay();
  host.update();
}

// ---------------------------------------------------------------- команды

/** Размер в px после изменения масштаба на k (с учётом обрезки пропорции сохраняются). */
export function scaleImage(k: number): void {
  const img = selectedImage();
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  const w = (img.node.attrs.width as number | null) ?? img.dom?.offsetWidth ?? 200;
  const h = (img.node.attrs.height as number | null) ?? img.dom?.offsetHeight ?? 150;
  setAttrs(img.pos, img.node, { width: Math.max(12, Math.round(w * k)), height: Math.max(12, Math.round(h * k)) });
  host.relayout();
}

export async function resetImageSize(): Promise<void> {
  const img = selectedImage();
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  const n = await naturalSize(img.node.attrs.src as string);
  const c = (img.node.attrs.crop as number[] | null) ?? [0, 0, 0, 0];
  setAttrs(img.pos, img.node, { width: Math.round(n.w * (1 - c[0] - c[2])), height: Math.round(n.h * (1 - c[1] - c[3])) });
  host.relayout();
}

/** По ширине полосы текста. */
export function fitImageWidth(): void {
  const img = selectedImage();
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  const w = (img.node.attrs.width as number | null) ?? img.dom?.offsetWidth ?? 200;
  const h = (img.node.attrs.height as number | null) ?? img.dom?.offsetHeight ?? 150;
  const tw = host.textWidth();
  setAttrs(img.pos, img.node, { width: Math.round(tw), height: Math.round((h * tw) / w) });
  host.relayout();
}

export function clearCrop(): void {
  const img = selectedImage();
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  const c = img.node.attrs.crop as number[] | null;
  if (!c) return;
  const w = (img.node.attrs.width as number) / (1 - c[0] - c[2]);
  const h = (img.node.attrs.height as number) / (1 - c[1] - c[3]);
  setAttrs(img.pos, img.node, { crop: null, width: Math.round(w), height: Math.round(h) });
  host.relayout();
}

/** Следующий рисунок документа (по кругу). */
export function selectNextImage(): void {
  const v = host.view();
  const list: number[] = [];
  v.state.doc.descendants((n, pos) => { if (n.type === schema.nodes.image) list.push(pos); });
  if (!list.length) { toast('В документе нет рисунков.'); return; }
  const cur = selectedImage()?.pos ?? v.state.selection.from - 1;
  const next = list.find((p) => p > cur) ?? list[0];
  v.dispatch(v.state.tr.setSelection(NodeSelection.create(v.state.doc, next)).scrollIntoView());
  v.focus();
}

/** Рисунок «на передний/задний план»: перед текстом или за текстом (положение сохраняется). */
export function arrangeImage(front: boolean): void {
  const img = selectedImage();
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  const f = floatOf(img.node);
  const base: Float = f ?? { mode: 'front', hRel: 'column', hOff: 0, hAlign: null, vRel: 'paragraph', vOff: 0, vAlign: null, dist: [0, 0, 8, 8] };
  setAttrs(img.pos, img.node, { float: { ...base, mode: front ? 'front' : 'behind' } });
  host.relayout();
}

/** Выровнять плавающий рисунок относительно страницы. */
export function alignImage(how: string): void {
  const img = selectedImage();
  const f = img ? floatOf(img.node) : null;
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  if (!f || !isAbs(f)) { toast('Выравнивание по странице работает для рисунков «за текстом» и «перед текстом».'); return; }
  const next: Float = { ...f };
  if (how === 'left' || how === 'center' || how === 'right') { next.hRel = 'page'; next.hAlign = how; next.hOff = null; }
  else { next.vRel = 'page'; next.vAlign = (how === 'middle' ? 'center' : how) as 'top' | 'center' | 'bottom'; next.vOff = null; }
  setAttrs(img.pos, img.node, { float: next });
  host.relayout();
}

export function deleteImage(): void {
  const img = selectedImage();
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  const v = host.view();
  v.dispatch(v.state.tr.delete(img.pos, img.pos + 1).scrollIntoView());
  v.focus();
}

/** Заменить картинку, сохранив размер, положение и обтекание. */
export function replaceImage(): void {
  const img = selectedImage();
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = 'image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml';
  input.addEventListener('change', async () => {
    const f = input.files?.[0];
    if (!f) return;
    const src = await new Promise<string>((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result as string); r.onerror = () => rej(r.error); r.readAsDataURL(f); });
    const cur = selectedImage();
    if (cur) setAttrs(cur.pos, cur.node, { src, crop: null });
    host.relayout();
  });
  input.click();
}

/** Подпись под рисунком: «Рисунок N — » по центру отдельным абзацем; курсор в конце. */
export function addCaption(): void {
  const img = selectedImage();
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  const v = host.view();
  const $p = v.state.doc.resolve(img.pos);
  const after = $p.after($p.depth);
  let n = 1;
  v.state.doc.descendants((node) => { if (node.isTextblock && /^Рисунок\s+\d+/.test(node.textContent)) n++; });
  const text = `Рисунок ${n} — `;
  const para = schema.nodes.paragraph.create({ align: 'center', before: 6, after: 12 }, schema.text(text));
  let tr = v.state.tr.insert(after, para);
  tr = tr.setSelection(TextSelection.create(tr.doc, after + 1 + text.length));
  v.dispatch(tr.scrollIntoView());
  v.focus();
}

// ---------------------------------------------------------------- окно «Формат рисунка»

const toCm = (px: number | null | undefined) => (px == null ? '' : String(Math.round((px / CM) * 100) / 100));
const fromCm = (s: string) => { const v = parseFloat(s.replace(',', '.')); return Number.isFinite(v) && v > 0 ? v * CM : null; };

export async function imageDialog(): Promise<void> {
  const img = selectedImage();
  if (!img) { toast('Сначала выделите рисунок (щёлкните по нему).'); return; }
  const natural = await naturalSize(img.node.attrs.src as string);
  const a = img.node.attrs;
  const w0 = (a.width as number | null) ?? natural.w;
  const h0 = (a.height as number | null) ?? natural.h;
  const crop0 = (a.crop as number[] | null) ?? [0, 0, 0, 0];
  const f0 = a.float as Float | null;
  dialog((box, close) => {
    const row = (label: string, ...k: Node[]) => el('div', { class: 'dlg-row' }, el('label', { textContent: label }), ...k);
    const num = (v: string) => el('input', { type: 'text', value: v, size: 7 });
    const width = num(toCm(w0)), height = num(toCm(h0));
    const lock = el('input', { type: 'checkbox', checked: true });
    const ratio = w0 / h0;
    width.addEventListener('input', () => { if (lock.checked) { const w = fromCm(width.value); if (w) height.value = toCm(w / ratio); } });
    height.addEventListener('input', () => { if (lock.checked) { const h = fromCm(height.value); if (h) width.value = toCm(h * ratio); } });
    const cs = crop0.map((c) => num(String(Math.round(c * 1000) / 10)));
    const wrap = el('select', {},
      ...[['inline', 'В тексте'], ['left', 'Слева, текст справа'], ['right', 'Справа, текст слева'], ['behind', 'За текстом'], ['front', 'Перед текстом']].map(([v, t]) => new Option(t, v)));
    wrap.value = !f0 ? 'inline' : f0.mode === 'wrap' ? (f0.side ?? 'left') : f0.mode;
    const hRel = el('select', {}, ...[['page', 'страницы'], ['margin', 'поля'], ['character', 'абзаца']].map(([v, t]) => new Option(t, v)));
    hRel.value = f0?.hRel === 'column' ? 'margin' : f0?.hRel ?? 'page';
    const hAl = el('select', {}, ...[['', 'смещение, см →'], ['left', 'слева'], ['center', 'по центру'], ['right', 'справа']].map(([v, t]) => new Option(t, v)));
    hAl.value = f0?.hAlign ?? '';
    const hOff = num(toCm(f0?.hOff ?? 0));
    const vRel = el('select', {}, ...[['page', 'страницы'], ['margin', 'поля'], ['paragraph', 'абзаца']].map(([v, t]) => new Option(t, v)));
    vRel.value = f0?.vRel === 'line' ? 'paragraph' : f0?.vRel ?? 'page';
    const vAl = el('select', {}, ...[['', 'смещение, см →'], ['top', 'вверху'], ['center', 'по центру'], ['bottom', 'внизу']].map(([v, t]) => new Option(t, v)));
    vAl.value = f0?.vAlign ?? '';
    const vOff = num(toCm(f0?.vOff ?? 0));
    const alt = el('input', { type: 'text', value: String(a.alt ?? ''), placeholder: 'Описание для программ экранного чтения' });
    const posBox = el('div', {},
      row('По горизонтали от', hRel, hAl, hOff),
      row('По вертикали от', vRel, vAl, vOff));
    const showPos = () => { posBox.hidden = wrap.value !== 'behind' && wrap.value !== 'front'; };
    wrap.addEventListener('change', showPos);
    showPos();
    const natBtn = el('button', { textContent: 'Исходный размер', onclick: () => { const k = [1 - crop0[0] - crop0[2], 1 - crop0[1] - crop0[3]]; width.value = toCm(natural.w * k[0]); height.value = toCm(natural.h * k[1]); } });
    const ok = el('button', { class: 'primary', textContent: 'Применить', onclick: () => {
      const w = fromCm(width.value) ?? w0, h = fromCm(height.value) ?? h0;
      const cr = cs.map((c) => Math.min(0.9, Math.max(0, (parseFloat(c.value.replace(',', '.')) || 0) / 100)));
      if (cr[0] + cr[2] > 0.95 || cr[1] + cr[3] > 0.95) { toast('Обрезка слишком большая.'); return; }
      let float: Float | null = null;
      const mode = wrap.value;
      const cm = (s: string) => { const v = parseFloat(s.replace(',', '.')); return Number.isFinite(v) ? Math.round(v * CM) : 0; };
      const base: Float = f0 ?? { mode: 'behind', hRel: 'page', hOff: 0, hAlign: null, vRel: 'page', vOff: 0, vAlign: null, dist: [0, 0, 8, 8] };
      if (mode === 'behind' || mode === 'front') {
        float = { ...base, mode, hRel: hRel.value as Float['hRel'], hAlign: (hAl.value || null) as Float['hAlign'], hOff: hAl.value ? null : cm(hOff.value), vRel: vRel.value as Float['vRel'], vAlign: (vAl.value || null) as Float['vAlign'], vOff: vAl.value ? null : cm(vOff.value) };
      } else if (mode === 'left' || mode === 'right') float = { ...base, mode: 'wrap', side: mode, dist: f0?.mode === 'wrap' ? f0.dist : [0, 4, 8, 8] };
      close();
      const cur = selectedImage();
      if (!cur) return;
      setAttrs(cur.pos, cur.node, { width: Math.round(w), height: Math.round(h), crop: cr.some((x) => x > 0) ? cr : null, float, alt: alt.value });
      host.relayout();
    } });
    box.append(
      el('h2', { textContent: 'Формат рисунка' }),
      row('Ширина, см', width, el('label', {}, lock, ' пропорционально')),
      row('Высота, см', height, natBtn),
      row('Обрезка, %', el('span', { class: 'img-crop-fields' }, 'слева ', cs[0], ' сверху ', cs[1], ' справа ', cs[2], ' снизу ', cs[3])),
      row('Обтекание', wrap),
      posBox,
      row('Описание', alt),
      el('div', { class: 'dlg-actions' }, el('button', { textContent: 'Отмена', onclick: close }), ok),
    );
  }, false);
}
