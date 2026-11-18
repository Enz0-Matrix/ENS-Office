// SPDX-License-Identifier: AGPL-3.0-or-later
// Формулы: ввод в стиле LaTeX («\frac{a}{b}», «x^2», «\sqrt{x}», «\alpha»), показ через MathML (есть в Chromium
// с версии 109, значит и в 128 на РЕД ОС 7) и запись в docx как формулы Word (OMML, m:oMath), чтение обратно.
//
// Дерево формулы общее для всех трёх представлений: LaTeX ⇄ дерево ⇄ MathML / OMML.

export type MNode =
  | { t: 'row'; items: MNode[] }
  | { t: 'id' | 'num' | 'op' | 'text'; v: string }
  | { t: 'frac'; num: MNode; den: MNode }
  | { t: 'sqrt'; body: MNode; deg?: MNode }
  | { t: 'sup'; base: MNode; sup: MNode }
  | { t: 'sub'; base: MNode; sub: MNode }
  | { t: 'subsup'; base: MNode; sub: MNode; sup: MNode }
  | { t: 'fence'; open: string; close: string; body: MNode };

const GREEK: Record<string, string> = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', varepsilon: 'ε', zeta: 'ζ', eta: 'η', theta: 'θ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν',
  xi: 'ξ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', upsilon: 'υ', phi: 'φ', varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
};
const OPS: Record<string, string> = {
  cdot: '·', times: '×', div: '÷', pm: '±', mp: '∓', le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠', approx: '≈', equiv: '≡', sim: '∼',
  infty: '∞', sum: '∑', prod: '∏', int: '∫', oint: '∮', partial: '∂', nabla: '∇', to: '→', rightarrow: '→', leftarrow: '←', Rightarrow: '⇒', Leftrightarrow: '⇔',
  in: '∈', notin: '∉', subset: '⊂', supset: '⊃', cup: '∪', cap: '∩', forall: '∀', exists: '∃', emptyset: '∅', degree: '°', angle: '∠', perp: '⊥', parallel: '∥',
  ldots: '…', cdots: '⋯', lim: 'lim', sin: 'sin', cos: 'cos', tg: 'tg', ctg: 'ctg', tan: 'tan', log: 'log', ln: 'ln', lg: 'lg', exp: 'exp', max: 'max', min: 'min',
};
const FUNCS = new Set(['lim', 'sin', 'cos', 'tg', 'ctg', 'tan', 'log', 'ln', 'lg', 'exp', 'max', 'min']);
const LARGE = new Set(['∑', '∏', '∫', '∮']);
const OP_CHARS = '+-−=<>±×÷·≤≥≠≈≡∼∞∑∏∫∮∂∇→←⇒⇔∈∉⊂⊃∪∩∀∃∅∠⊥∥…⋯,;:!|/*';

/** Разбор LaTeX-подобной записи. Ошибки не бросаем: непонятное остаётся текстом. */
export function parseTex(src: string): MNode {
  let i = 0;
  const peek = () => src[i];
  const skip = () => { while (i < src.length && src[i] === ' ') i++; };
  const group = (): MNode => {
    skip();
    if (peek() === '{') { i++; const r = row('}'); i++; return r; }
    return atom() ?? { t: 'row', items: [] };
  };
  const atom = (): MNode | null => {
    skip();
    const c = peek();
    if (c === undefined || c === '}' ) return null;
    if (c === '{') { i++; const r = row('}'); i++; return r; }
    if (c === '\\') {
      i++;
      const m = /^[A-Za-z]+|^./.exec(src.slice(i))?.[0];
      if (!m) return { t: 'text', v: '\\' }; // «\» в конце: набор ещё не закончен
      i += m.length;
      if (m === 'frac') { const num = group(); const den = group(); return { t: 'frac', num, den }; }
      if (m === 'sqrt') {
        skip();
        let deg: MNode | undefined;
        if (peek() === '[') { i++; const s = i; while (i < src.length && src[i] !== ']') i++; deg = parseTex(src.slice(s, i)); i++; }
        return { t: 'sqrt', body: group(), deg };
      }
      if (m === 'text' || m === 'mathrm') { skip(); if (peek() === '{') { i++; const s = i; while (i < src.length && src[i] !== '}') i++; const v = src.slice(s, i); i++; return { t: 'text', v }; } }
      if (m === 'left' || m === 'right') return atom();
      if (GREEK[m]) return { t: 'id', v: GREEK[m] };
      if (OPS[m]) return FUNCS.has(m) ? { t: 'text', v: OPS[m] } : { t: 'op', v: OPS[m] };
      if (m === ',' || m === ';' || m === ' ') return { t: 'op', v: ' ' };
      if (m === '{' || m === '}') return { t: 'op', v: m };
      return { t: 'text', v: '\\' + m };
    }
    if (c === '(' || c === '[' || c === '|') {
      const close = c === '(' ? ')' : c === '[' ? ']' : '|';
      const at = i;
      i++;
      const body = row(close);
      if (peek() === close) { i++; return { t: 'fence', open: c, close, body }; }
      i = at + 1;
      return { t: 'op', v: c };
    }
    if (/[0-9.,]/.test(c) && /[0-9]/.test(c)) { const m = /^[0-9]+(?:[.,][0-9]+)?/.exec(src.slice(i))![0]; i += m.length; return { t: 'num', v: m }; }
    i++;
    if (c === '-') return { t: 'op', v: '−' };
    if (c === '*') return { t: 'op', v: '·' };
    if (OP_CHARS.includes(c) || c === ')' || c === ']') return { t: 'op', v: c };
    return { t: 'id', v: c };
  };
  const row = (until?: string): MNode => {
    const items: MNode[] = [];
    for (;;) {
      skip();
      if (i >= src.length || (until && peek() === until)) break;
      let a = atom();
      if (!a) { if (peek() === '}') { i++; continue; } break; }
      // Индексы: x_1^2, x^2_1.
      for (;;) {
        skip();
        if (peek() === '^' || peek() === '_') {
          const k = src[i++];
          const g = group();
          if (a.t === 'sub' && k === '^') a = { t: 'subsup', base: a.base, sub: a.sub, sup: g };
          else if (a.t === 'sup' && k === '_') a = { t: 'subsup', base: a.base, sub: g, sup: a.sup };
          else a = k === '^' ? { t: 'sup', base: a, sup: g } : { t: 'sub', base: a, sub: g };
        } else break;
      }
      items.push(a);
    }
    return items.length === 1 ? items[0] : { t: 'row', items };
  };
  return row();
}

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export function toMathML(n: MNode): string {
  const r = (x: MNode): string => {
    switch (x.t) {
      case 'row': return `<mrow>${x.items.map(r).join('')}</mrow>`;
      case 'id': return `<mi>${esc(x.v)}</mi>`;
      case 'num': return `<mn>${esc(x.v)}</mn>`;
      case 'op': return LARGE.has(x.v) ? `<mo largeop="true">${esc(x.v)}</mo>` : `<mo>${esc(x.v)}</mo>`;
      case 'text': return `<mi mathvariant="normal">${esc(x.v)}</mi>`;
      case 'frac': return `<mfrac>${r(x.num)}${r(x.den)}</mfrac>`;
      case 'sqrt': return x.deg ? `<mroot>${r(x.body)}${r(x.deg)}</mroot>` : `<msqrt>${r(x.body)}</msqrt>`;
      case 'sup': return `<msup>${r(x.base)}${r(x.sup)}</msup>`;
      case 'sub': return `<msub>${r(x.base)}${r(x.sub)}</msub>`;
      case 'subsup': return x.base.t === 'op' && LARGE.has(x.base.v) ? `<munderover>${r(x.base)}${r(x.sub)}${r(x.sup)}</munderover>` : `<msubsup>${r(x.base)}${r(x.sub)}${r(x.sup)}</msubsup>`;
      case 'fence': return `<mrow><mo>${esc(x.open)}</mo>${r(x.body)}<mo>${esc(x.close)}</mo></mrow>`;
    }
  };
  return `<math xmlns="http://www.w3.org/1998/Math/MathML">${r(n)}</math>`;
}

/** OMML (формулы Word). Обычный текст — m:r/m:t; дроби, корни, степени, индексы, скобки, n-арные операторы. */
export function toOMML(n: MNode): string {
  const run = (v: string, plain = false) => `<m:r>${plain ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : ''}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t xml:space="preserve">${esc(v)}</m:t></m:r>`;
  const e = (x: MNode) => `<m:e>${r(x)}</m:e>`;
  const r = (x: MNode): string => {
    switch (x.t) {
      case 'row': return x.items.map(r).join('');
      case 'id': case 'num': case 'op': return run(x.v, x.t !== 'id');
      case 'text': return run(x.v, true);
      case 'frac': return `<m:f><m:num>${r(x.num)}</m:num><m:den>${r(x.den)}</m:den></m:f>`;
      case 'sqrt': return x.deg ? `<m:rad><m:deg>${r(x.deg)}</m:deg>${e(x.body)}</m:rad>` : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${e(x.body)}</m:rad>`;
      case 'sup': return `<m:sSup>${e(x.base)}<m:sup>${r(x.sup)}</m:sup></m:sSup>`;
      case 'sub': return `<m:sSub>${e(x.base)}<m:sub>${r(x.sub)}</m:sub></m:sSub>`;
      case 'subsup':
        if (x.base.t === 'op' && LARGE.has(x.base.v)) return `<m:nary><m:naryPr><m:chr m:val="${esc(x.base.v)}"/></m:naryPr><m:sub>${r(x.sub)}</m:sub><m:sup>${r(x.sup)}</m:sup><m:e/></m:nary>`;
        return `<m:sSubSup>${e(x.base)}<m:sub>${r(x.sub)}</m:sub><m:sup>${r(x.sup)}</m:sup></m:sSubSup>`;
      case 'fence': return `<m:d><m:dPr><m:begChr m:val="${esc(x.open)}"/><m:endChr m:val="${esc(x.close)}"/></m:dPr>${e(x.body)}</m:d>`;
    }
  };
  return `<m:oMath>${r(n)}</m:oMath>`;
}

/** OMML → LaTeX-подобная запись (для правки и показа). Неизвестные элементы читаются по тексту. */
export function ommlToTex(el: Element): string {
  const kids = (e: Element, name?: string) => [...e.children].filter((c) => !name || c.localName === name);
  const kid = (e: Element, name: string) => kids(e, name)[0];
  const val = (e: Element | undefined, name: string) => (e ? kid(e, name)?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/math', 'val') ?? kid(e, name)?.getAttribute('m:val') ?? null : null);
  const back: Record<string, string> = {};
  for (const [k, v] of Object.entries({ ...GREEK, ...OPS })) if (!back[v] && !FUNCS.has(k)) back[v] = '\\' + k;
  const text = (s: string) => [...s].map((ch) => (ch === '−' ? '-' : back[ch] ? back[ch] + ' ' : ch)).join('');
  const g = (s: string) => (s.length === 1 ? s : `{${s}}`);
  const w = (e: Element | undefined): string => (e ? kids(e).map(r).join('') : '');
  const r = (e: Element): string => {
    switch (e.localName) {
      case 'r': return text(kids(e, 't').map((t) => t.textContent ?? '').join(''));
      case 'f': return `\\frac{${w(kid(e, 'num'))}}{${w(kid(e, 'den'))}}`;
      case 'rad': {
        const deg = w(kid(e, 'deg'));
        return deg ? `\\sqrt[${deg}]{${w(kid(e, 'e'))}}` : `\\sqrt{${w(kid(e, 'e'))}}`;
      }
      case 'sSup': return `${g(w(kid(e, 'e')))}^${g(w(kid(e, 'sup')))}`;
      case 'sSub': return `${g(w(kid(e, 'e')))}_${g(w(kid(e, 'sub')))}`;
      case 'sSubSup': return `${g(w(kid(e, 'e')))}_${g(w(kid(e, 'sub')))}^${g(w(kid(e, 'sup')))}`;
      case 'nary': {
        const ch = val(kid(e, 'naryPr'), 'chr') ?? '∫';
        const body = w(kid(e, 'e'));
        return `${text(ch)}_${g(w(kid(e, 'sub')))}^${g(w(kid(e, 'sup')))}${body ? ' ' + body : ''}`;
      }
      case 'd': {
        const pr = kid(e, 'dPr');
        const open = val(pr, 'begChr') ?? '(', close = val(pr, 'endChr') ?? ')';
        return `${open}${kids(e, 'e').map(w).join(', ')}${close}`;
      }
      case 'func': return `${w(kid(e, 'fName'))} ${w(kid(e, 'e'))}`;
      case 'rPr': case 'ctrlPr': case 'oMathParaPr': return '';
      default: return w(e);
    }
  };
  return r(el).replace(/\s+/g, ' ').trim();
}
