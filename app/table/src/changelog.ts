// Журнал изменений книги: кто, когда и что менял (п. 8). Хранится в самом
// файле (xlsx, customXml — Excel его сохраняет), поэтому виден всем, кто
// открывает книгу с общего диска. При открытии показывается, что изменилось
// с вашего прошлого открытия, и замечается правка в другой программе
// (по «отпечаткам» содержимого листов).

import { columnName, type Model } from './engine';
import { saveSettings, settings } from './settings';
import { APP_VERSION } from './version';

export interface LogEntry {
  /** Кто сохранил. */
  u: string;
  /** Когда (мс). */
  t: number;
  /** Чем: «ENS-Table 1.0.0». */
  app: string;
  /** Что: листы и диапазоны, заметки (удалённые листы, строки…). */
  ch: { sheets: [string, string][]; notes: string[] };
  /** Отпечатки содержимого листов на момент сохранения. */
  h: Record<string, string>;
}

const MAX_ENTRIES = 300;
const MAX_LAST_OPENED = 200;

export function readLog(m: Model): LogEntry[] {
  try {
    const v = JSON.parse(m.getChangeLog() || '[]');
    return Array.isArray(v) ? v.filter((e) => e && typeof e.t === 'number') : [];
  } catch { return []; }
}

function hashes(m: Model): Record<string, string> {
  const out: Record<string, string> = {};
  m.getWorksheetsProperties().forEach((p, i) => { try { out[p.name] = m.sheetContentHash(i); } catch { /* пропустить */ } });
  return out;
}

/** Листы, изменённые не в ENS-Table после последней записи журнала. */
function changedOutside(m: Model, last: LogEntry | undefined): string[] {
  if (!last?.h) return [];
  const now = hashes(m);
  const names = new Set([...Object.keys(now), ...Object.keys(last.h)]);
  return [...names].filter((n) => now[n] !== last.h[n]);
}

export function userName(os: string): string {
  return settings.userName.trim() || os || 'Пользователь';
}

/** Перед сохранением: дописать запись (кто, что, отпечатки). Возвращает прежний журнал,
 *  чтобы вернуть его, если сохранить не удалось. */
export function beforeSave(m: Model, user: string): string {
  const old = m.getChangeLog();
  try {
    const log = readLog(m);
    const ch = JSON.parse(m.getEnsChanges()) as LogEntry['ch'];
    const outside = changedOutside(m, log[log.length - 1]);
    // Правку в другой программе тоже отмечаем: кто сохранил файл там — из свойств файла.
    if (outside.length && log.length) {
      const by = m.getLastModifiedBy();
      log.push({ u: by && by !== 'IronCalc User' ? by : '?', t: Date.now() - 1, app: 'другая программа', ch: { sheets: outside.map((n) => [n, '']), notes: [] }, h: {} });
    }
    const h = hashes(m);
    if (ch.sheets.length || ch.notes.length || !log.length) {
      log.push({ u: user, t: Date.now(), app: `ENS-Table ${APP_VERSION}`, ch, h });
    } else {
      log[log.length - 1].h = h; // ничего не меняли (только сохранили) — обновить отпечатки
    }
    m.setChangeLog(JSON.stringify(log.slice(-MAX_ENTRIES)));
    m.setLastModifiedBy(user);
  } catch { /* журнал — не повод не сохранить книгу */ }
  return old;
}

export function afterSave(m: Model, path: string): void {
  m.clearEnsChanges();
  markOpened(path);
}

function markOpened(path: string): void {
  if (!path) return;
  const all = { ...settings.lastOpened, [path]: Date.now() };
  const keys = Object.keys(all).sort((a, b) => all[b] - all[a]).slice(0, MAX_LAST_OPENED);
  settings.lastOpened = Object.fromEntries(keys.map((k) => [k, all[k]]));
  saveSettings();
}

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);
const when = (t: number) => new Date(t).toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: '2-digit', hour: '2-digit', minute: '2-digit' });

export interface ChangesHost {
  goTo(sheet: string, area: string): void;
  /** Подсветить области на листах (null — снять). */
  mark?(items: [string, string][] | null): void;
  /** Выбрать другую версию файла и показать отличия. */
  compare?(): void;
}

export interface BookDiff {
  sheets: [string, string][];
  notes: string[];
  cells: number;
}

/** Сравнение двух книг по содержимому ячеек (формулы и значения): где отличаются. */
export function diffBooks(a: Model, b: Model): BookDiff {
  const pa = a.getWorksheetsProperties(), pb = b.getWorksheetsProperties();
  const out: BookDiff = { sheets: [], notes: [], cells: 0 };
  const nb = new Map(pb.map((p, i) => [p.name, i]));
  pa.forEach((p, i) => {
    const j = nb.get(p.name);
    if (j === undefined) { out.notes.push(`Лист «${p.name}» есть только в открытой книге`); return; }
    const da = a.getSheetDimension(i), db = b.getSheetDimension(j);
    const r2 = Math.max(da[2], db[2]), c2 = Math.max(da[3], db[3]);
    const rows = new Map<number, [number, number][]>();
    let n = 0;
    for (let r = 1; r <= r2; r++) {
      let run: [number, number] | null = null;
      for (let c = 1; c <= c2 + 1; c++) {
        const diff = c <= c2 && a.getCellContent(i, r, c) !== b.getCellContent(j, r, c);
        if (diff) { n++; if (run) run[1] = c; else run = [c, c]; }
        else if (run) { (rows.get(r) ?? rows.set(r, []).get(r)!).push(run); run = null; }
      }
    }
    if (!n) return;
    out.cells += n;
    // Соседние строки с одинаковыми столбцами сливаются в один прямоугольник.
    const open = new Map<string, [number, number, number, number]>();
    const areas: [number, number, number, number][] = [];
    for (let r = 1; r <= r2 + 1; r++) {
      const here = new Set((rows.get(r) ?? []).map(([x, y]) => `${x}:${y}`));
      for (const [k, v] of [...open]) if (!here.has(k)) { areas.push(v); open.delete(k); }
      for (const [x, y] of rows.get(r) ?? []) {
        const k = `${x}:${y}`, o = open.get(k);
        if (o) o[2] = r; else open.set(k, [r, x, r, y]);
      }
    }
    const name = (q: [number, number, number, number]) => `${columnName(q[1])}${q[0]}` + (q[0] === q[2] && q[1] === q[3] ? '' : `:${columnName(q[3])}${q[2]}`);
    out.sheets.push([p.name, areas.sort((x, y) => x[0] - y[0] || x[1] - y[1]).map(name).join(' ')]);
  });
  pb.forEach((p) => { if (!pa.some((q) => q.name === p.name)) out.notes.push(`Лист «${p.name}» есть только в другой версии`); });
  return out;
}

/** Показать результат сравнения: вместо записей журнала — отличия, с подсветкой на листах. */
export function showDiff(host: ChangesHost, other: string, d: BookDiff): void {
  const entry: LogEntry = { u: 'Отличия от', t: Date.now(), app: other, ch: { sheets: d.sheets, notes: d.notes }, h: {} };
  show(host, 'Сравнение версий', d.cells || d.notes.length
    ? `Отличается ячеек: <b>${d.cells}</b>. Подсвечены на листах; кнопка с адресом — переход.`
    : 'Содержимое ячеек одинаково.', [entry]);
  host.mark?.(d.sheets);
}

function entryHtml(e: LogEntry, i: number): string {
  const parts = [
    ...(e.ch?.sheets ?? []).map(([s, areas]) => `<div class="chg-sheet"><b>${esc(s)}</b>${areas
      ? areas.split(' ').slice(0, 12).map((a) => ` <button class="chg-area" data-sheet="${esc(s)}" data-area="${esc(a)}">${esc(a)}</button>`).join('') + (areas.split(' ').length > 12 ? ' …' : '')
      : ' — изменён'}</div>`),
    ...(e.ch?.notes ?? []).map((n) => `<div class="chg-note">${esc(n)}</div>`),
  ];
  const mk = (e.ch?.sheets ?? []).some(([, a]) => a) ? ` <button class="chg-mark" data-i="${i}" title="Подсветить изменённые ячейки на листах">подсветить</button>` : '';
  return `<li><div class="chg-who">${esc(e.u)} · ${when(e.t)}${mk}${e.app && !e.app.startsWith('ENS-Table') ? ` · <i>${esc(e.app)}</i>` : ''}</div>${parts.join('') || '<div class="chg-note">сохранено без изменений</div>'}</li>`;
}

let panel: HTMLElement | null = null;

function show(host: ChangesHost, title: string, intro: string, entries: LogEntry[]): void {
  panel?.remove();
  panel = document.createElement('div');
  panel.className = 'changes-panel';
  panel.innerHTML = `<div class="chg-head"><b>${esc(title)}</b><span><button class="chg-clear" title="Снять подсветку">снять подсветку</button>${host.compare ? '<button class="chg-compare" title="Сравнить с другой версией файла">сравнить с файлом…</button>' : ''}<button class="chg-close" title="Закрыть">✕</button></span></div>
    ${intro ? `<p class="chg-intro">${intro}</p>` : ''}
    <ul class="chg-list">${entries.map((e, i) => entryHtml(e, i)).reverse().join('')}</ul>`;
  document.body.append(panel);
  panel.addEventListener('click', (e) => {
    const t = e.target as HTMLElement;
    if (t.closest('.chg-close')) { panel?.remove(); panel = null; host.mark?.(null); return; }
    if (t.closest('.chg-clear')) { host.mark?.(null); return; }
    const mk = t.closest<HTMLElement>('.chg-mark');
    if (mk) { host.mark?.(entries[Number(mk.dataset.i)].ch.sheets.filter(([, a]) => a)); return; }
    if (t.closest('.chg-compare')) { host.compare?.(); return; }
    const a = t.closest<HTMLElement>('.chg-area');
    if (a) host.goTo(a.dataset.sheet!, a.dataset.area!);
  });
}

/** При открытии файла: что изменилось с вашего прошлого открытия и правки в другой программе. */
export function onOpened(m: Model, path: string, user: string, host: ChangesHost): void {
  const log = readLog(m);
  const since = settings.lastOpened[path] ?? 0;
  markOpened(path);
  const fresh = since ? log.filter((e) => e.t > since && e.u !== user) : [];
  const outside = changedOutside(m, log[log.length - 1]);
  if (!fresh.length && !outside.length) return;
  const entries = [...fresh];
  let intro = since ? `С вашего прошлого открытия (${when(since)}):` : '';
  if (outside.length) {
    const by = m.getLastModifiedBy();
    entries.push({ u: by && by !== 'IronCalc User' ? by : 'неизвестно', t: log[log.length - 1].t, app: 'другая программа (Excel?)', ch: { sheets: outside.map((n) => [n, '']), notes: [] }, h: {} });
    intro += `${intro ? ' ' : ''}Файл менялся и в другой программе после последнего сохранения в ENS-Table.`;
  }
  show(host, 'Изменения в книге', intro, entries);
}

/** «Файл → Сведения → Журнал изменений»: вся история книги. */
export function showLog(m: Model, host: ChangesHost): void {
  const log = readLog(m);
  const pending = JSON.parse(m.getEnsChanges() || '{"sheets":[],"notes":[]}') as LogEntry['ch'];
  const entries = [...log];
  if (pending.sheets.length || pending.notes.length) entries.push({ u: 'вы (не сохранено)', t: Date.now(), app: '', ch: pending, h: {} });
  show(host, 'Журнал изменений', entries.length ? '' : 'Журнал пуст: записи появляются при сохранении книги в ENS-Table.', entries.slice(-100));
}
