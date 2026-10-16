// Обезличенный синтетический лист, похожий по структуре на «ВСЕ Филиалы»:
// ключ-«госномер», филиал, числа, ВПР в справочник, произведение, ЕСЛИ, ОКРУГЛ, итоговая СУММ.
// Одинаковые данные для обоих прототипов (детерминированный генератор).

export const LOOKUP_ROWS = 2000;
export const HEADERS = ['№', 'Ключ', 'Филиал', 'Кол-во', 'Цена', 'Ставка (ВПР)', 'Сумма', 'Размер', 'Итог', 'ВСЕГО'];
export const COLS = HEADERS.length;

function rng(seed) { // mulberry32
  return () => {
    seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

// Excel-ОКРУГЛ: половина — от нуля (2,5 → 3), с поправкой на двоичную погрешность.
export function excelRound(x, d) {
  const m = 10 ** d;
  const s = Math.sign(x), a = Math.abs(x) * m;
  return (s * Math.round(Number(a.toPrecision(15)))) / m;
}

/** rows — число строк данных. Возвращает «сырые» значения и формулы по ячейкам (0-based). */
export function makeDataset(rows, useLookup = true) {
  const r = rng(42);
  const lookup = [];
  for (let i = 0; i < LOOKUP_ROWS; i++) lookup.push([`К${String(i).padStart(4, '0')}`, Math.round(r() * 2000) / 100]);
  const rate = new Map(lookup);
  const data = []; // [row][col] = {v} | {f}
  data.push(HEADERS.map(h => ({ v: h })));
  let expectedTotal = 0;
  for (let i = 0; i < rows; i++) {
    const n = i + 2; // номер строки Excel
    const li = Math.floor(r() * LOOKUP_ROWS), key = lookup[li][0];
    const qty = 1 + Math.floor(r() * 50);
    const price = Math.round(r() * 10000) / 100;
    const sum = qty * price;
    expectedTotal += excelRound(sum * rate.get(key) / 100, 2);
    data.push([
      { v: i + 1 }, { v: key }, { v: 'ДЛЗХ'[i % 4] }, { v: qty }, { v: price },
      useLookup ? { f: `=VLOOKUP(B${n},Lookup!$A$1:$B$${LOOKUP_ROWS},2,0)` } : { f: `=Lookup!B${li + 1}` },
      { f: `=D${n}*E${n}` },
      { f: `=IF(G${n}>500,"крупн","мелк")` },
      { f: `=ROUND(G${n}*F${n}/100,2)` },
      i === 0 ? { f: `=SUM(I2:I${rows + 1})` } : null,
    ]);
  }
  // Для теста правки берём ключ из первой строки данных — он гарантированно используется.
  const editLookupRow = lookup.findIndex(([k]) => k === data[1][1].v);
  return { data, lookup, editLookupRow, expectedTotal: excelRound(expectedTotal, 2) };
}

// Правка ставки в справочнике: пересчитываются строки с этим ключом (~rows/2000) и ВСЕГО.
export const EDIT_VALUE = 99.99;

// Уступить управление без таймера: в фоновой вкладке setTimeout замедляется до 1 раза в секунду.
const mc = new MessageChannel(); const q = [];
mc.port1.onmessage = () => q.shift()?.();
export const yieldNow = () => new Promise(r => { q.push(r); mc.port2.postMessage(0); });

export function memMB() {
  return performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1048576) : null;
}

/** Программная прокрутка: frames шагов. stepMs — синхронное время шага (для своей сетки это и есть
 *  отрисовка кадра). frameMs — интервал между кадрами браузера; меряется, только пока вкладка видима
 *  (в фоне браузер кадры не рисует), поэтому ждём кадр не дольше 250 мс. */
export async function measureScroll(step, frames = 240) {
  const stepT = [], frameT = [];
  for (let i = 0; i < frames; i++) {
    const s = performance.now();
    step(i);
    stepT.push(performance.now() - s);
    if (document.visibilityState === 'visible') {
      const got = await Promise.race([new Promise(r => requestAnimationFrame(() => r(true))), new Promise(r => setTimeout(() => r(false), 250))]);
      if (got) frameT.push(performance.now() - s);
    } else await yieldNow();
  }
  const stat = (t) => { if (!t.length) return 'нет данных (вкладка в фоне)'; t.sort((a, b) => a - b); return `ср ${(t.reduce((a, b) => a + b, 0) / t.length).toFixed(1)} / p95 ${t[Math.floor(t.length * .95)].toFixed(1)} / макс ${t.at(-1).toFixed(1)} мс`; };
  return { scrollStep: stat(stepT), scrollFrame: stat(frameT), framesSeen: frameT.length };
}

export function report(name, res) {
  window.__bench = res;
  const el = document.getElementById('report');
  el.textContent = `${name}\n` + Object.entries(res).map(([k, v]) => `${k}: ${v}`).join('\n');
  console.log('BENCH', name, JSON.stringify(res));
}

export function lookupParam() {
  return new URLSearchParams(location.search).get('lookup') !== '0';
}

export function rowsParam() {
  return Number(new URLSearchParams(location.search).get('rows') || 50000);
}
